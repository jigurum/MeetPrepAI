import express, { Request, Response } from 'express';
import { createServer as createViteServer } from 'vite';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import { GoogleGenAI } from '@google/genai';
import mammoth from 'mammoth';
import JSZip from 'jszip';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = Number(process.env.PORT) || 3000;

// Support large file uploads in base64
app.use(express.json({ limit: '40mb' }));
app.use(express.urlencoded({ extended: true, limit: '40mb' }));

// Initialize Gemini client strictly as per skill guidelines
const ai = new GoogleGenAI({
  apiKey: process.env.GEMINI_API_KEY,
  httpOptions: {
    headers: {
      'User-Agent': 'aistudio-build',
    },
  },
});

/**
 * Text extractor helper for PPTX files
 */
async function extractPptxText(buffer: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(buffer);
  const slideTexts: string[] = [];
  
  // Find all slide XML files
  const slideFiles = Object.keys(zip.files).filter((fileName) =>
    fileName.match(/^ppt\/slides\/slide\d+\.xml$/i)
  );

  // Sort slides numerically
  slideFiles.sort((a, b) => {
    const numA = parseInt(a.replace(/\D/g, ''), 10) || 0;
    const numB = parseInt(b.replace(/\D/g, ''), 10) || 0;
    return numA - numB;
  });

  for (let i = 0; i < slideFiles.length; i++) {
    const xml = await zip.files[slideFiles[i]].async('text');
    // Extract text inside <a:t> tags
    const matches = xml.match(/<a:t[^>]*>(.*?)<\/a:t>/g);
    if (matches && matches.length > 0) {
      const texts = matches.map((m) => m.replace(/<[^>]+>/g, '').trim()).filter(Boolean);
      if (texts.length > 0) {
        slideTexts.push(`[슬라이드 ${i + 1}]\n` + texts.join(' '));
      }
    }
  }

  return slideTexts.join('\n\n');
}

/**
 * Endpoint to extract text from uploaded files (DOCX, PPTX, TXT/MD, etc.)
 */
app.post('/api/extract-text', async (req: Request, res: Response): Promise<void> => {
  try {
    const { fileName, fileType, base64Data } = req.body;
    if (!base64Data) {
      res.status(400).json({ error: '파일 데이터가 누락되었습니다.' });
      return;
    }

    const buffer = Buffer.from(base64Data, 'base64');
    const ext = (fileName || '').split('.').pop()?.toLowerCase();

    let extractedText = '';

    if (ext === 'docx') {
      const result = await mammoth.extractRawText({ buffer });
      extractedText = result.value.trim();
    } else if (ext === 'pptx') {
      extractedText = await extractPptxText(buffer);
    } else if (ext === 'txt' || ext === 'md' || ext === 'markdown' || ext === 'csv' || ext === 'json') {
      extractedText = buffer.toString('utf-8').trim();
    } else if (ext === 'pdf') {
      // PDF will be sent directly to Gemini as inlineData multimodal part
      res.json({
        isPdf: true,
        message: 'PDF 파일은 Gemini 비전/문서 분석 엔진으로 직접 심층 분석됩니다.',
        textPreview: '[PDF 바이너리 문서 분석 대기]',
      });
      return;
    } else {
      res.status(400).json({
        error: `지원하지 않는 파일 형식입니다: .${ext}. PDF, DOCX, PPTX, TXT 파일을 사용해 주세요.`,
      });
      return;
    }

    if (!extractedText) {
      res.status(400).json({
        error: '파일에서 텍스트를 추출할 수 없습니다. 텍스트가 포함된 정상적인 문서인지 확인해 주세요.',
      });
      return;
    }

    res.json({
      success: true,
      fileName,
      extractedText,
      charCount: extractedText.length,
    });
  } catch (error: any) {
    console.error('Text extraction error:', error);
    res.status(500).json({
      error: `파일 텍스트 추출 중 오류가 발생했습니다: ${error.message || '알 수 없는 오류'}`,
    });
  }
});

/**
 * Main AI Analysis Endpoint
 * Analyzes meeting materials using Gemini 3.8 Flash
 */
app.post('/api/analyze', async (req: Request, res: Response): Promise<void> => {
  try {
    const {
      title,
      level, // 'beginner' | 'practitioner'
      inputType, // 'text' | 'file'
      rawText,
      fileData, // { mimeType, base64, fileName }
    } = req.body;

    if (!title) {
      res.status(400).json({ error: '자료명을 입력해 주세요.' });
      return;
    }

    if (!rawText && !fileData?.base64) {
      res.status(400).json({ error: '회의 자료 내용 또는 파일을 입력해 주세요.' });
      return;
    }

    const isBeginner = level === 'beginner';

    const systemInstruction = `
당신은 한국 기업 회의 준비를 돕는 최고의 회의 분석 전문가 "MeetPrep AI"입니다.
목표: 회의 자료에 포함된 전문용어와 배경지식을 이해하기 어려운 직장인(신입사원 및 비전문 분야 참석자)이 회의에 자신 있게 참석하도록 돕습니다.

[설명 수준 모드: ${isBeginner ? '입문자 모드 (Beginner)' : '실무자 모드 (Practitioner)'}]
${
  isBeginner
    ? `* 입문자 모드 핵심 원칙:
- 배경지식이 전혀 없는 신입사원도 즉시 이해할 수 있도록 쉬운 비유와 일상 예시를 적극 활용하세요.
- 불필요한 또 다른 기술 전문용어의 나열을 엄격히 배제하세요.
- 친절하고 직관적인 어조로 설명하세요.`
    : `* 실무자 모드 핵심 원칙:
- 실무 엔지니어 및 프로젝트 관리자를 위해 정확한 기술적 정의와 아키텍처 맥락을 제공하세요.
- 실제 업무 워크플로, 시스템 연계, 유사/경쟁 기술과의 차이점, 실무 주의사항을 포함하세요.
- 전문적이고 밀도 높은 실무형 어조로 설명하세요.`
}

[필수 작성 규칙 (Strict Rules)]:
1. 자료에서 확인한 사실과 AI가 보충한 일반 지식을 명확히 구분하세요.
2. 자료에 없는 내용을 사실처럼 지어내지 마세요(Hallucination 금지).
3. 회의 목적 등 자료에서 명시적으로 확인되지 않는 항목은 반드시 "자료에서 확인되지 않음"으로 표기하세요.
4. 중요 설명 및 논의 사항에는 반드시 자료의 실제 원문 문장(original quote)을 연결해 주세요.
5. 전문용어는 오직 실제 입력된 자료에 직접 등장한 용어/약어/개념만 추출하세요(최소 4개 ~ 최대 12개).
6. 예상 질문 및 답변(Should 기능): 회의 시 실제 질문받을 가능성이 높은 문항 3~4개를 작성하고, 각 답변이 '자료 근거(document)'인지 '실무 추론(inference)'인지 명확히 분류하세요.
7. 모든 출력은 한국어로 정중하고 명확하게 작성하며, 제공된 JSON 스키마를 엄격히 준수하세요.
`;

    // Construct user content parts
    const contentParts: any[] = [];

    if (fileData?.mimeType === 'application/pdf' && fileData?.base64) {
      contentParts.push({
        inlineData: {
          mimeType: 'application/pdf',
          data: fileData.base64,
        },
      });
      contentParts.push({
        text: `다음 첨부된 회의 자료(PDF)를 정밀 분석해 주세요.\n자료명: "${title}"\n입력 방식: 파일 업로드(PDF)\n설명 수준: ${isBeginner ? '입문자 모드' : '실무자 모드'}`
      });
    } else {
      contentParts.push({
        text: `다음 회의 자료를 분석해 주세요.\n자료명: "${title}"\n입력 방식: ${inputType === 'file' ? `파일 업로드 (${fileData?.fileName || '문서'})` : '텍스트 직접 입력'}\n설명 수준: ${isBeginner ? '입문자 모드' : '실무자 모드'}\n\n[회의 자료 본문 시작]\n${rawText}\n[회의 자료 본문 끝]`
      });
    }

    const promptInstruction = `
위 자료를 분석하여 아래 JSON 구조로만 정확하게 반환하세요.
{
  "topic": "회의 주제 요약 (1~2문장)",
  "purpose": "회의 목적 (자료에 명시된 경우 상세 서술, 자료에 없으면 정확히 '자료에서 확인되지 않음'으로 작성)",
  "purposeConfirmed": true 또는 false (자료에서 목적이 직접 확인되면 true, 없으면 false),
  "keyDiscussionItems": [
    {
      "title": "논의 주제",
      "description": "상세 논의 내용",
      "originalQuote": "자료에 등장하는 관련 원문 문장 (자료에 정확히 일치하는 인용구)",
      "isFromDocument": true
    }
  ],
  "technicalConcepts": [
    {
      "name": "기술/업무 개념명",
      "role": "자료 내에서의 역할과 중요성"
    }
  ],
  "keyPoints": [
    {
      "point": "이해해야 할 핵심 포인트",
      "details": "구체적 설명",
      "isFact": true
    }
  ],
  "backgroundKnowledge": {
    "basicConcepts": [
      {
        "concept": "기초 개념명",
        "explanation": "회의 이해에 필수적인 기초 지식 설명",
        "relevance": "이 회의와 어떻게 연결되는지"
      }
    ],
    "workProcesses": [
      {
        "stage": "단계명 (예: 기획/설계, 구현/연동, 테스트/배포)",
        "processDescription": "해당 업무 프로세스 흐름 설명"
      }
    ],
    "techRelationships": [
      {
        "entityA": "개념 A",
        "relationship": "관계 설명 (예: 하위 컴포넌트로 동작, 입력 데이터를 전달)",
        "entityB": "개념 B"
      }
    ],
    "recommendedPrep": [
      "사전에 미리 숙지하면 좋은 팁이나 확인 사항"
    ]
  },
  "terminology": [
    {
      "term": "자료에 등장한 전문용어 또는 약어",
      "definition": "간단 명료한 표준 정의 (1문장)",
      "easyExplanation": "${isBeginner ? '신입사원 눈높이의 쉬운 비유와 직관적 설명' : '실무 스펙 기반의 정밀 설명 및 기술 맥락'}",
      "contextualMeaning": "이번 회의 문맥 및 본 자료에서 의미하는 바",
      "practicalUseCases": "실제 현업 실무에서 쓰이는 사례 또는 업무 예시",
      "sourceQuote": "자료에서 이 용어가 사용된 실제 문장 또는 문맥",
      "levelNote": "${isBeginner ? '이해를 돕는 일상적 비유 한 줄' : '관련 유사 기술과의 차이점 또는 실무 팁'}"
    }
  ],
  "expectedQA": [
    {
      "question": "회의 중 예상되는 실전 질문",
      "suggestedAnswer": "모범 답변 또는 검토 방향",
      "answerBasis": "document" 또는 "inference",
      "sourceExcerpt": "자료 근거일 경우 원문 발췌, 추론일 경우 추론 배경"
    }
  ]
}
`;

    contentParts.push({ text: promptInstruction });

    // Call Gemini with automatic retry for transient errors (e.g., 503 UNAVAILABLE or 429)
    let responseText: string | undefined;
    const modelsToTry = ['gemini-3.8-flash', 'gemini-flash-latest'];
    let lastError: any = null;

    for (const modelName of modelsToTry) {
      for (let attempt = 1; attempt <= 3; attempt++) {
        try {
          const response = await ai.models.generateContent({
            model: modelName,
            contents: contentParts,
            config: {
              systemInstruction,
              responseMimeType: 'application/json',
            },
          });
          responseText = response.text;
          if (responseText) break;
        } catch (err: any) {
          lastError = err;
          console.warn(`[Gemini API] Attempt ${attempt} on ${modelName} failed:`, err?.message || err);
          // If 503 or 429, wait before retrying
          if (attempt < 3) {
            await new Promise((r) => setTimeout(r, 1000 * attempt));
          }
        }
      }
      if (responseText) break;
    }

    if (!responseText) {
      throw lastError || new Error('Gemini API로부터 응답 텍스트를 수신하지 못했습니다.');
    }

    let parsedResult: any;
    try {
      parsedResult = JSON.parse(responseText);
    } catch (parseError) {
      // Attempt cleanup if wrapped in markdown code blocks
      const cleanJson = responseText.replace(/```json/g, '').replace(/```/g, '').trim();
      parsedResult = JSON.parse(cleanJson);
    }

    // Generate unique IDs
    const analysisId = 'MP-' + Date.now().toString(36).toUpperCase() + '-' + Math.random().toString(36).substring(2, 6).toUpperCase();
    const analyzedAt = new Date().toISOString();

    // Enrich terminology with IDs
    const enrichedTerminology = (parsedResult.terminology || []).map((term: any, idx: number) => ({
      termId: `${analysisId}-T${String(idx + 1).padStart(2, '0')}`,
      analysisId,
      term: term.term || '미지정 용어',
      definition: term.definition || '',
      easyExplanation: term.easyExplanation || '',
      contextualMeaning: term.contextualMeaning || '',
      practicalUseCases: term.practicalUseCases || '',
      sourceQuote: term.sourceQuote || '',
      levelNote: term.levelNote || '',
      bookmarked: false,
    }));

    const finalResult = {
      analysisId,
      analyzedAt,
      title,
      inputType: inputType === 'file' ? `파일 (${fileData?.fileName || '문서'})` : '직접 입력',
      level,
      levelLabel: isBeginner ? '입문자' : '실무자',
      summary: {
        topic: parsedResult.topic || title,
        purpose: parsedResult.purpose || '자료에서 확인되지 않음',
        purposeConfirmed: Boolean(parsedResult.purposeConfirmed),
        keyDiscussionItems: parsedResult.keyDiscussionItems || [],
        technicalConcepts: parsedResult.technicalConcepts || [],
        keyPoints: parsedResult.keyPoints || [],
      },
      backgroundKnowledge: parsedResult.backgroundKnowledge || {
        basicConcepts: [],
        workProcesses: [],
        techRelationships: [],
        recommendedPrep: [],
      },
      terminology: enrichedTerminology,
      expectedQA: parsedResult.expectedQA || [],
      sourcePreview: rawText ? rawText.substring(0, 300) + (rawText.length > 300 ? '...' : '') : `${fileData?.fileName || '파일'} 분석 완료`,
    };

    res.json(finalResult);
  } catch (error: any) {
    console.error('Gemini Analysis Error:', error);

    let clientMsg = 'AI 분석 처리 중 오류가 발생했습니다. 잠시 후 다시 시도해 주세요.';
    const errMsg = error?.message || String(error);

    if (errMsg.includes('429') || errMsg.includes('RESOURCE_EXHAUSTED')) {
      clientMsg =
        'Gemini API 무료 티어 분당 호출 한도(RPM)에 도달했습니다. 약 15초 후 다시 시도해 주세요.';
    } else if (errMsg.includes('503') || errMsg.includes('UNAVAILABLE')) {
      clientMsg =
        'Gemini 서버 트래픽이 일시적으로 급증했습니다. 몇 초 후 다시 분석을 시도해 주세요.';
    } else if (errMsg.includes('403') || errMsg.includes('PERMISSION_DENIED')) {
      clientMsg = 'API 키 권한 오류가 발생했습니다. 유효한 API 키인지 확인해 주세요.';
    } else if (errMsg) {
      try {
        const parsed = JSON.parse(errMsg);
        if (parsed?.error?.message) {
          clientMsg = parsed.error.message;
        }
      } catch {
        clientMsg = errMsg.slice(0, 200);
      }
    }

    res.status(500).json({ error: clientMsg });
  }
});

/**
 * Proxy endpoint to communicate with Google Apps Script Web App
 * Bypasses browser CORS limitations and follows 302 redirects automatically
 */
app.post('/api/sheets-proxy', async (req: Request, res: Response): Promise<void> => {
  try {
    const { scriptUrl, action, payload } = req.body;

    if (!scriptUrl) {
      res.status(400).json({ error: 'Google Apps Script 웹 앱 URL이 필요합니다.' });
      return;
    }

    // Validate URL format
    if (!scriptUrl.startsWith('https://script.google.com/macros/s/')) {
      res.status(400).json({
        error: '유효한 Google Apps Script 웹 앱 URL이 아닙니다. ("https://script.google.com/macros/s/..." 형식이어야 합니다.)',
      });
      return;
    }

    const requestBody = JSON.stringify({
      action: action || 'ping',
      ...payload,
    });

    const gasResponse = await fetch(scriptUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: requestBody,
      redirect: 'follow', // Automatically follow GAS 302 redirects
    });

    if (!gasResponse.ok) {
      const errText = await gasResponse.text();
      res.status(gasResponse.status).json({
        error: `Apps Script 응답 오류 (${gasResponse.status}): ${errText.slice(0, 300)}`,
      });
      return;
    }

    const responseText = await gasResponse.text();
    let responseData: any;
    try {
      responseData = JSON.parse(responseText);
    } catch {
      responseData = { message: responseText };
    }

    res.json({
      success: true,
      data: responseData,
    });
  } catch (error: any) {
    console.error('GAS Proxy Error:', error);
    res.status(500).json({
      error: `Google Apps Script 연동 실패: ${error.message || '연결할 수 없습니다. URL 및 배포 권한("모든 사용자")을 확인해 주세요.'}`,
    });
  }
});

// Vite middleware for development, or static hosting for production
if (process.env.NODE_ENV !== 'production') {
  const vite = await createViteServer({
    server: { middlewareMode: true },
    appType: 'spa',
  });
  app.use(vite.middlewares);
} else {
  app.use(express.static(path.resolve(__dirname, 'dist')));
  app.get('*', (_req: Request, res: Response) => {
    res.sendFile(path.resolve(__dirname, 'dist', 'index.html'));
  });
}

app.listen(PORT, '0.0.0.0', () => {
  console.log(`[MeetPrep AI Server] Listening on http://0.0.0.0:${PORT}`);
});
