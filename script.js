/**
 * MeetPrep AI - Pure Vanilla JavaScript Engine (ES6+)
 * Directly runnable on GitHub Pages without any build tools.
 */

// ==========================================
// 1. Google Apps Script Template & Steps
// ==========================================
const GOOGLE_APPS_SCRIPT_CODE = `/**
 * ==============================================================================
 * MeetPrep AI - Google Sheets 백엔드 연동 Apps Script
 * ==============================================================================
 * [시트 1: 분석 기록] 분석ID | 분석일시 | 자료명 | 입력방식 | 설명수준 | 핵심요약 | 배경지식
 * [시트 2: 전문용어]   용어ID | 분석ID | 전문용어 | 맞춤형 설명 | 실무 활용 사례 | 북마크
 * ==============================================================================
 */

var SHEET_ANALYSIS_NAME = "분석 기록";
var SHEET_TERMS_NAME = "전문용어";

function doGet(e) {
  return handleRequest(e, "GET");
}

function doPost(e) {
  return handleRequest(e, "POST");
}

function handleRequest(e, method) {
  var output;
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    var action = data.action || "ping";
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 2개 시트 생성 및 서식 초기화
    var sheetAnalysis = ensureSheet(ss, SHEET_ANALYSIS_NAME, [
      "분석ID", "분석일시", "자료명", "입력방식", "설명수준", "핵심요약", "배경지식"
    ], "#1E3A8A");

    var sheetTerms = ensureSheet(ss, SHEET_TERMS_NAME, [
      "용어ID", "분석ID", "전문용어", "맞춤형 설명", "실무 활용 사례", "북마크"
    ], "#2563EB");

    var responseData = {};

    switch (action) {
      case "ping":
      case "test":
        responseData = {
          success: true,
          message: "MeetPrep AI와 Google Sheets가 성공적으로 연결되었습니다!",
          spreadsheetName: ss.getName(),
          spreadsheetUrl: ss.getUrl(),
          analysisCount: Math.max(0, sheetAnalysis.getLastRow() - 1),
          termsCount: Math.max(0, sheetTerms.getLastRow() - 1)
        };
        break;

      case "saveAnalysis":
        responseData = saveAnalysisRecord(sheetAnalysis, sheetTerms, data.analysis);
        break;

      case "getAnalyses":
        responseData = getAllAnalyses(sheetAnalysis, sheetTerms);
        break;

      case "toggleBookmark":
        responseData = updateBookmarkStatus(sheetTerms, data.termId, data.bookmarked);
        break;

      case "deleteAnalysis":
        responseData = deleteAnalysisRecord(sheetAnalysis, sheetTerms, data.analysisId);
        break;

      default:
        responseData = { success: false, error: "알 수 없는 요청: " + action };
        break;
    }

    output = ContentService.createTextOutput(JSON.stringify(responseData));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  } catch (err) {
    output = ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.toString()
    }));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  }
}

function ensureSheet(ss, name, headers, headerColor) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground(headerColor);
    headerRange.setFontColor("#FFFFFF");
    headerRange.setFontWeight("bold");
    headerRange.setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    for (var i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
  }
  return sheet;
}

function saveAnalysisRecord(sheetAnalysis, sheetTerms, analysis) {
  if (!analysis || !analysis.analysisId) {
    return { success: false, error: "유효한 분석 데이터가 없습니다." };
  }

  var existingIds = [];
  var lastRow = sheetAnalysis.getLastRow();
  if (lastRow > 1) {
    existingIds = sheetAnalysis.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  }

  var summaryText = "";
  if (analysis.summary) {
    summaryText = "[주제] " + (analysis.summary.topic || "") + "\\n" +
                  "[목적] " + (analysis.summary.purpose || "") + "\\n\\n" +
                  "[주요 논의 사항]\\n" +
                  (analysis.summary.keyDiscussionItems || []).map(function(item, idx) {
                    return (idx + 1) + ". " + item.title + ": " + item.description;
                  }).join("\\n") + "\\n\\n" +
                  "[핵심 포인트]\\n" +
                  (analysis.summary.keyPoints || []).map(function(kp) {
                    return "• " + kp.point + " (" + kp.details + ")";
                  }).join("\\n");
  }

  var bgText = "";
  if (analysis.backgroundKnowledge) {
    var bg = analysis.backgroundKnowledge;
    bgText = "[기초 개념]\\n" +
             (bg.basicConcepts || []).map(function(c) {
               return "• " + c.concept + ": " + c.explanation;
             }).join("\\n") + "\\n\\n" +
             "[업무 프로세스]\\n" +
             (bg.workProcesses || []).map(function(p) {
               return "• [" + p.stage + "] " + p.processDescription;
             }).join("\\n") + "\\n\\n" +
             "[기술 관계]\\n" +
             (bg.techRelationships || []).map(function(r) {
               return "• " + r.entityA + " -> " + r.relationship + " -> " + r.entityB;
             }).join("\\n");
  }

  var rowIdx = existingIds.indexOf(analysis.analysisId);
  var rowData = [
    analysis.analysisId,
    analysis.analyzedAt || new Date().toISOString(),
    analysis.title || "제목 없음",
    analysis.inputType || "직접 입력",
    analysis.levelLabel || (analysis.level === "beginner" ? "입문자" : "실무자"),
    summaryText,
    bgText
  ];

  if (rowIdx >= 0) {
    sheetAnalysis.getRange(rowIdx + 2, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheetAnalysis.appendRow(rowData);
  }

  // 전문용어 시트 갱신
  var termsLastRow = sheetTerms.getLastRow();
  if (termsLastRow > 1) {
    var termsAnalysisIds = sheetTerms.getRange(2, 2, termsLastRow - 1, 1).getValues().flat();
    for (var i = termsAnalysisIds.length - 1; i >= 0; i--) {
      if (termsAnalysisIds[i] === analysis.analysisId) {
        sheetTerms.deleteRow(i + 2);
      }
    }
  }

  var terms = analysis.terminology || [];
  for (var t = 0; t < terms.length; t++) {
    var term = terms[t];
    var termDesc = "[정의] " + (term.definition || "") + "\\n" +
                   "[설명] " + (term.easyExplanation || "") + "\\n" +
                   "[문맥] " + (term.contextualMeaning || "");
    sheetTerms.appendRow([
      term.termId || (analysis.analysisId + "-T" + (t + 1)),
      analysis.analysisId,
      term.term || "",
      termDesc,
      term.practicalUseCases || "",
      term.bookmarked ? true : false
    ]);
  }

  return {
    success: true,
    message: "Google Sheets에 분석 기록과 " + terms.length + "개 전문용어가 저장되었습니다.",
    analysisId: analysis.analysisId
  };
}

function updateBookmarkStatus(sheetTerms, termId, bookmarked) {
  var lastRow = sheetTerms.getLastRow();
  if (lastRow <= 1) return { success: false, error: "전문용어 데이터가 없습니다." };
  var termIds = sheetTerms.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  var idx = termIds.indexOf(termId);
  if (idx >= 0) {
    sheetTerms.getRange(idx + 2, 6).setValue(bookmarked ? true : false);
    return { success: true, termId: termId, bookmarked: bookmarked };
  }
  return { success: false, error: "용어를 찾을 수 없습니다: " + termId };
}

function deleteAnalysisRecord(sheetAnalysis, sheetTerms, analysisId) {
  var lastRow = sheetAnalysis.getLastRow();
  if (lastRow > 1) {
    var ids = sheetAnalysis.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    var idx = ids.indexOf(analysisId);
    if (idx >= 0) sheetAnalysis.deleteRow(idx + 2);
  }

  var termsLastRow = sheetTerms.getLastRow();
  if (termsLastRow > 1) {
    var termAIds = sheetTerms.getRange(2, 2, termsLastRow - 1, 1).getValues().flat();
    for (var i = termAIds.length - 1; i >= 0; i--) {
      if (termAIds[i] === analysisId) sheetTerms.deleteRow(i + 2);
    }
  }

  return { success: true, message: "분석 기록이 시트에서 삭제되었습니다." };
}

function getAllAnalyses(sheetAnalysis, sheetTerms) {
  var lastRow = sheetAnalysis.getLastRow();
  if (lastRow <= 1) return { success: true, analyses: [] };

  var analysisRows = sheetAnalysis.getRange(2, 1, lastRow - 1, 7).getValues();
  var list = [];
  for (var j = 0; j < analysisRows.length; j++) {
    var a = analysisRows[j];
    list.push({
      analysisId: a[0],
      analyzedAt: a[1],
      title: a[2],
      inputType: a[3],
      levelLabel: a[4],
      summaryRaw: a[5],
      bgRaw: a[6]
    });
  }
  return { success: true, analyses: list };
}
`;

const APPS_SCRIPT_GUIDE_STEPS = [
  {
    step: 1,
    title: '새 Google 스프레드시트 생성',
    description: 'Google Drive(drive.google.com)로 이동하여 [새로 만들기 > Google 스프레드시트]를 생성합니다. 파일명을 "MeetPrep AI 회의 기록" 등으로 지정하세요.',
  },
  {
    step: 2,
    title: 'Apps Script 편집기 열기',
    description: '스프레드시트 상단 메뉴에서 [확장 프로그램] > [Apps Script]를 클릭합니다.',
  },
  {
    step: 3,
    title: '코드 붙여넣기 및 저장',
    description: '기존의 `myFunction()` 코드를 지우고, 아래 [Apps Script 코드 복사] 버튼으로 복사한 코드를 그대로 붙여넣은 뒤 상단 디스켓(저장) 아이콘을 누릅니다.',
  },
  {
    step: 4,
    title: '웹 앱으로 배포 (100% 무료)',
    description: '우측 상단 파란색 [배포] 버튼 > [새 배포] 클릭 ➔ 톱니바퀴 아이콘에서 [웹 앱] 선택 ➔ [다음에 액세스할 수 있는 사용자]를 "모든 사용자(Anyone)"로 선택 후 [배포]를 누릅니다. (※ Google Apps Script는 완전 무료 서비스이며, 모든 사용자 배포 시에도 비용이 전혀 발생하지 않습니다.)',
  },
  {
    step: 5,
    title: '웹 앱 URL 등록 및 테스트',
    description: '배포 완료 후 나타나는 "웹 앱 URL (https://script.google.com/macros/s/.../exec)"을 복사하여 아래 URL 입력창에 붙여넣고 [연결 테스트]를 누르면 완료됩니다!',
  },
];

// ==========================================
// 2. Preset Sample Documents
// ==========================================
const SAMPLE_DOCUMENTS = [
  {
    id: 'sample-1',
    title: '클라우드 네이티브 MSA 전환 2차 아키텍처 회의',
    category: '인프라/백엔드',
    description: '모놀리식에서 마이크로서비스로 전환하며 gRPC, 서비스 메시(Istio), 멱등성 보장 방안을 검토하는 회의록',
    content: `[회의명] 모놀리식 커머스 플랫폼의 MSA(Microservices Architecture) 2차 전환 검토
[일시] 2026-11-04 14:00 ~ 15:30 (대회의실 A)
[참석자] 인프라팀장, 백엔드 파트장, 결제 모듈 담당자, 데브옵스 엔지니어

1. 회의 목적
기존 단일 애플리케이션(Monolithic) 환경에서 주문 폭주 시 발생하던 DB 데드락과 단일 장애점(SPOF)을 해소하기 위해, 주문/결제/회원 도메인을 독립 서비스로 분리하고 클라우드 네이티브 기반 서비스 메시(Service Mesh) 및 비동기 이벤트 파이프라인을 도입하기 위함.

2. 주요 논의 안건
(1) 서비스 간 통신 프로토콜 표준화:
- RESTful HTTP/1.1 API의 오버헤드를 줄이기 위해 내부 마이크로서비스 간 통신은 Protocol Buffers 기반의 gRPC(HTTP/2)를 채택하기로 합의함.
- 외부 모바일/웹 클라이언트와의 접점에는 BFF(Backend For Frontend) 패턴 및 API Gateway(Kong)를 배치하여 라우팅 및 JWT 토큰 인가를 단일화함.

(2) 서비스 메시(Service Mesh) 도입:
- 쿠버네티스(Kubernetes) 클러스터 내 서비스 간 mTLS(상호 TLS) 암호화 통신 및 카나리(Canary) 배포 트래픽 제어를 위해 Istio 사이드카 프록시(Envoy)를 적용할 계획임.
- 인프라팀에서는 Envoy 사이드카 인젝션에 따른 파드 메모리 추가 사용량(약 50MB/pod)에 대한 리소스 프로비저닝 예산을 산정 중임.

(3) 분산 트랜잭션과 데이터 정합성:
- 결제 완료 후 재고 차감 및 배송 지시 단계에서 2PC(Two-Phase Commit)의 락킹 병목을 피하기 위해, 사가(Saga) 패턴(코레오그래피 방식)을 채택함.
- 메시지 브로커로 Apache Kafka를 사용하여 이벤트를 발행(Publish)하며, 네트워크 순단이나 컨슈머 재시도에 대비하여 각 서비스 수신부에 멱등성(Idempotency) 보장 키(UUID)를 저장하고 중복 처리를 차단하는 로직을 필수 구현하기로 함.

(4) 관측 가능성(Observability):
- OpenTelemetry SDK를 백엔드 서비스 전반에 내장하여 분산 추적(Distributed Tracing)용 TraceID를 헤더로 전파하고, Jaeger를 통해 지연 시간(Latency) 병목 구간을 시각화함.

3. 액션 아이템
- 백엔드팀: 다음 주 수요일까지 사가 패턴 실패 시 실행할 보상 트랜잭션(Compensating Transaction) 명세서 초안 작성
- 인프라팀: 스테이징 환경 EKS 클러스터에 Istio 1.24 설치 및 부하 테스트 벤치마크 진행
- 결제팀: PG사 연동 모듈의 결제 승인 웹훅 수신 시 멱등키 캐싱(Redis TTL 24시간) 로직 구현`,
  },
  {
    id: 'sample-2',
    title: '글로벌 결제 PG사 연동 및 정산 시스템 개편 스프린트',
    category: '핀테크/결제',
    description: '신규 PG 연동, 3D Secure 2.0 인증, 에스크로 정산 배치 및 이상금융거래탐지(FDS) 적용 논의',
    content: `[회의명] 해외 결제 PG 연동 및 실시간 정산 마감 시스템 킥오프
[일시] 2026-11-05 10:00 ~ 11:30
[작성자] 핀테크 프로덕트 오너

1. 개요 및 배경
동남아 및 일본 시장 신규 진출을 위해 Stripe 및 현지 로컬 결제 수단(E-Wallet)을 통합 수용하는 글로벌 결제 파이프라인을 구축함. 또한 카드 결제 부정 사용에 대비한 3D Secure 2.0 규격 준수와 일일 정산 대사(Reconciliation) 자동화가 필수적임.

2. 세부 검토 사항
(1) 결제 승인 플로우 및 3DS 2.0:
- 부정 결제 분쟁(Chargeback) 발생 시 가맹점 책임 면제(Liability Shift) 혜택을 받기 위해 Visa/Mastercard의 3-D Secure 2.0 위험 기반 인증(Risk-based Authentication)을 필수 적용함.
- 사용자 이탈률을 낮추기 위해 저위험 트랜잭션은 Frictionless Flow(무인증 통과)로 유도하고, 고위험 시에만 생체/OTP 챌린지를 요청함.

(2) 정산 대사(Settlement Reconciliation) 파이프라인:
- 매일 자정 PG사로부터 수신되는 결제 원장(Raw Clearing Data)과 자사 데이터베이스의 주문 레코드를 1:1 대조하는 배치(Batch) 잡 개발.
- 불일치 건(금액 차이, 승인 취소 누락, 환율 오차)은 별도의 에스크로(Escrow) 보류 계정에 임시 격리 후 재무팀 슬랙 채널로 실시간 알림 발송.

(3) 멱등성(Idempotency) 및 웹훅(Webhook) 처리:
- PG사 통신망 장애로 인한 중복 승인 요청을 방지하기 위해 모든 API 호출에 Client Generated Idempotency-Key를 전송함.
- 웹훅은 비동기 큐(SQS)를 통해 유실 없이 소비하며, Signature 해시 검증(HMAC-SHA256)을 거쳐 변조 여부를 검증함.

3. 미확인 사항 및 주의사항
- 현지 부가가치세(VAT) 원천징수 세율은 현지 법률 대리인의 검토가 아직 완료되지 않음.`,
  },
  {
    id: 'sample-3',
    title: '사내 생성형 AI PoC 결과 보고 및 데이터 보안 가이드 심의',
    category: 'AI/보안',
    description: '사내 문서 검색 RAG 구축 결과와 민감정보 유출 방지(DLP), 프롬프트 인젝션 방어책 심의',
    content: `[회의명] 생성형 AI 기반 사내 지식 관리 시스템 PoC 최종 평가 및 정보보호 심의
[일시] 2026-11-06 15:00
[배석] AI 혁신랩 연구원, 정보보안팀장, 컴플라이언스 변호사

1. 회의 목적
지난 3개월간 수행한 사내 Confluence 및 기술 문서 기반 RAG(Retrieval-Augmented Generation) 시스템의 검색 정확도와 응답 품질을 보고하고, 운영 환경 승격을 위한 보안 통제 가이드라인을 확정함.

2. 주요 보고 내용
(1) RAG 시스템 아키텍처 및 벤치마크 결과:
- 사내 기술 문서 20만 건을 청킹(Chunk Size 512, Overlap 50)하여 벡터 데이터베이스(Milvus)에 Dense Vector로 색인함.
- 검색 단계에서 키워드 기반 BM25(Sparse Vector)와 밀집 벡터 검색을 결합한 하이브리드 검색(Hybrid Search) 및 Re-ranking(Cross-Encoder)을 적용하여 Top-5 문서 적합도(Recall@5)가 78%에서 92%로 향상됨.
- 환각 현상(Hallucination) 억제를 위해 모델 프롬프트에 "컨텍스트에 없는 내용은 반드시 불확실성을 표시할 것"을 명시함.

(2) 정보보호 및 컴플라이언스 요구사항:
- DLP(Data Loss Prevention) 필터를 API 전단에 배치하여 주민등록번호, 계좌번호, 사내 비밀 코드(API Key, DB 연결 문자열)가 프롬프트로 전송될 경우 마스킹(Masking) 처리함.
- 간접 프롬프트 인젝션(Indirect Prompt Injection) 방어를 위해 외부 문서 내 악의적 시스템 지시문 무력화 샌드박스를 구성함.
- 사외 모델 호출 시 제로 데이터 리텐션(Zero Data Retention, ZDR) 계약 체결 여부를 법무팀과 최종 검토함.

3. 의결 사항
- 1단계로 비기밀성 공개 기술 규격서에 한해 사내 베타 오픈 승인.
- 인사/재무 데이터는 권한 기반 접근 제어(RBAC) 연동 완료 후 2차 심의를 거치기로 함.`,
  },
];

// ==========================================
// 3. Application State & Storage
// ==========================================
const STORAGE_KEY_ANALYSES = 'meetprep_analyses_v1';
const STORAGE_KEY_SETTINGS = 'meetprep_settings_v1';

const state = {
  currentTab: 'input', // 'input' | 'result' | 'history'
  activeResSubTab: 'summary', // 'summary' | 'background' | 'terminology' | 'qa'
  level: 'beginner', // 'beginner' | 'practitioner'
  inputMode: 'file', // 'file' | 'text'
  analyses: [],
  activeAnalysis: null,
  settings: {
    googleAppsScriptUrl: '',
    autoSyncToSheets: true,
    defaultLevel: 'beginner',
  },
  uploadedFile: null, // { name, size, type, base64, extractedText, isPdf }
  isAnalyzing: false,
  expandedTermId: null,
  termSearchKeyword: '',
  historySearchKeyword: '',
  historyFilterLevel: 'all',
  bookmarksSearchKeyword: '',
};

function loadStorage() {
  try {
    const rawAnalyses = localStorage.getItem(STORAGE_KEY_ANALYSES);
    if (rawAnalyses) state.analyses = JSON.parse(rawAnalyses) || [];
  } catch (e) {
    state.analyses = [];
  }

  try {
    const rawSettings = localStorage.getItem(STORAGE_KEY_SETTINGS);
    if (rawSettings) state.settings = { ...state.settings, ...JSON.parse(rawSettings) };
  } catch (e) {}

  if (state.analyses.length > 0) {
    state.activeAnalysis = state.analyses[0];
  }
}

function saveAnalysesToStorage() {
  try {
    localStorage.setItem(STORAGE_KEY_ANALYSES, JSON.stringify(state.analyses));
  } catch (e) {
    console.error('Storage save error:', e);
  }
}

function saveSettingsToStorage() {
  try {
    localStorage.setItem(STORAGE_KEY_SETTINGS, JSON.stringify(state.settings));
  } catch (e) {
    console.error('Settings save error:', e);
  }
}

// ==========================================
// 4. UI Helper Functions & Icons
// ==========================================
function refreshIcons() {
  if (window.lucide) {
    window.lucide.createIcons();
  }
}

function showToast(message, type = 'success') {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `flex items-center gap-2.5 px-4 py-3 rounded-xl shadow-lg border text-xs font-semibold mb-2 animate-in slide-in-from-bottom-2 ${
    type === 'success'
      ? 'bg-slate-900 text-white border-slate-800'
      : type === 'error'
      ? 'bg-red-600 text-white border-red-700'
      : 'bg-slate-800 text-white border-slate-700'
  }`;

  const iconName = type === 'success' ? 'check-circle-2' : 'alert-circle';
  toast.innerHTML = `
    <i data-lucide="${iconName}" class="w-4 h-4 shrink-0 text-emerald-400"></i>
    <span>${message}</span>
    <button type="button" class="ml-2 text-slate-400 hover:text-white">&times;</button>
  `;

  toast.querySelector('button').addEventListener('click', () => toast.remove());
  container.appendChild(toast);
  refreshIcons();

  setTimeout(() => {
    toast.remove();
  }, 4000);
}

function formatFileSize(bytes) {
  if (!bytes) return '0 B';
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function escapeHtml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// ==========================================
// 5. Client-Side Document Text Extractors
// ==========================================
async function extractTextFromDocxClient(arrayBuffer) {
  if (window.mammoth) {
    const result = await window.mammoth.extractRawText({ arrayBuffer });
    return result.value.trim();
  }
  throw new Error('Mammoth DOCX 라이브러리를 불러오지 못했습니다.');
}

async function extractTextFromPptxClient(arrayBuffer) {
  if (!window.JSZip) {
    throw new Error('JSZip 라이브러리를 불러오지 못했습니다.');
  }
  const zip = await window.JSZip.loadAsync(arrayBuffer);
  const slideFiles = Object.keys(zip.files).filter((fileName) =>
    fileName.match(/^ppt\/slides\/slide\d+\.xml$/i)
  );

  slideFiles.sort((a, b) => {
    const numA = parseInt(a.replace(/\D/g, ''), 10) || 0;
    const numB = parseInt(b.replace(/\D/g, ''), 10) || 0;
    return numA - numB;
  });

  const texts = [];
  for (let i = 0; i < slideFiles.length; i++) {
    const xml = await zip.files[slideFiles[i]].async('text');
    const matches = xml.match(/<a:t[^>]*>(.*?)<\/a:t>/g);
    if (matches && matches.length > 0) {
      const slideText = matches.map((m) => m.replace(/<[^>]+>/g, '').trim()).filter(Boolean);
      if (slideText.length > 0) {
        texts.push(`[슬라이드 ${i + 1}]\n` + slideText.join(' '));
      }
    }
  }
  return texts.join('\n\n');
}

// Convert file to base64
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => {
      const res = reader.result;
      const base64 = res.split(',')[1] || '';
      resolve(base64);
    };
    reader.onerror = (e) => reject(e);
  });
}

// ==========================================
// 6. Google Sheets API Client Functions
// ==========================================
async function callSheetsApi(action, payload = {}) {
  const scriptUrl = state.settings.googleAppsScriptUrl;
  if (!scriptUrl) return { success: false, error: 'Google Apps Script URL이 설정되지 않았습니다.' };

  // First try proxy if server is present (bypasses CORS)
  try {
    const res = await fetch('/api/sheets-proxy', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ scriptUrl, action, payload }),
    });
    if (res.ok) {
      const result = await res.json();
      return result;
    }
  } catch (e) {
    // If running on static GitHub Pages, proxy won't exist -> proceed to direct fetch
  }

  // Direct fetch fallback for GitHub Pages
  try {
    const res = await fetch(scriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action, ...payload }),
      redirect: 'follow',
    });
    const text = await res.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch {
      data = { message: text };
    }
    return { success: true, data };
  } catch (err) {
    return { success: false, error: err.message || 'Apps Script 통신 실패' };
  }
}

// ==========================================
// 7. Navigation & Screen Switching
// ==========================================
function switchScreen(tabName) {
  state.currentTab = tabName;

  const screens = {
    input: document.getElementById('screen-input'),
    result: document.getElementById('screen-result'),
    history: document.getElementById('screen-history'),
  };

  const navBtns = {
    input: document.getElementById('tab-btn-input'),
    result: document.getElementById('tab-btn-result'),
    history: document.getElementById('tab-btn-history'),
  };

  Object.keys(screens).forEach((key) => {
    if (screens[key]) {
      if (key === tabName) {
        screens[key].classList.remove('hidden');
      } else {
        screens[key].classList.add('hidden');
      }
    }
    if (navBtns[key]) {
      if (key === tabName) {
        navBtns[key].className =
          'nav-tab-btn flex items-center gap-2 px-3.5 py-2 text-sm font-semibold rounded-lg bg-slate-100 text-slate-900 transition-colors';
      } else {
        navBtns[key].className =
          'nav-tab-btn flex items-center gap-2 px-3.5 py-2 text-sm font-medium rounded-lg text-slate-600 hover:text-slate-900 hover:bg-slate-50 transition-colors';
      }
    }
  });

  if (state.activeAnalysis) {
    document.getElementById('tab-btn-result')?.classList.remove('hidden');
  }

  if (tabName === 'result' && state.activeAnalysis) {
    renderResultScreen();
  } else if (tabName === 'history') {
    renderHistoryScreen();
  }

  window.scrollTo({ top: 0, behavior: 'smooth' });
  refreshIcons();
}

// ==========================================
// 8. Render Presets & Setup UI
// ==========================================
function renderPresets() {
  const container = document.getElementById('preset-list-container');
  if (!container) return;

  container.innerHTML = SAMPLE_DOCUMENTS.map(
    (doc) => `
    <button
      type="button"
      data-sample-id="${doc.id}"
      class="preset-card text-left p-3 rounded-lg border text-xs bg-white border-slate-200 hover:border-blue-300 text-slate-700 hover:bg-slate-50 transition-all"
    >
      <div class="flex items-center justify-between font-semibold mb-1">
        <span class="truncate">${escapeHtml(doc.title)}</span>
        <span class="text-[10px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-600">${escapeHtml(doc.category)}</span>
      </div>
      <p class="text-[11px] text-slate-500 line-clamp-2 leading-tight">${escapeHtml(doc.description)}</p>
    </button>
  `
  ).join('');

  container.querySelectorAll('.preset-card').forEach((btn) => {
    btn.addEventListener('click', () => {
      const id = btn.getAttribute('data-sample-id');
      const doc = SAMPLE_DOCUMENTS.find((d) => d.id === id);
      if (doc) {
        document.getElementById('input-title').value = doc.title;
        document.getElementById('textarea-rawtext').value = doc.content;
        document.getElementById('text-char-count').textContent = doc.content.length.toLocaleString() + ' 자';

        // Switch to direct text subtab
        selectSubTab('text');
        showToast(`'${doc.title}' 예시 자료를 불러왔습니다.`, 'info');
      }
    });
  });
}

function selectSubTab(mode) {
  state.inputMode = mode;
  const btnFile = document.getElementById('btn-subtab-file');
  const btnText = document.getElementById('btn-subtab-text');
  const paneFile = document.getElementById('subtab-file-pane');
  const paneText = document.getElementById('subtab-text-pane');

  if (mode === 'file') {
    btnFile.className =
      'flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-blue-600 text-white shadow-sm transition-colors';
    btnText.className =
      'flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900 hover:bg-slate-100 transition-colors';
    paneFile.classList.remove('hidden');
    paneText.classList.add('hidden');
  } else {
    btnText.className =
      'flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-blue-600 text-white shadow-sm transition-colors';
    btnFile.className =
      'flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900 hover:bg-slate-100 transition-colors';
    paneText.classList.remove('hidden');
    paneFile.classList.add('hidden');
  }
}

function selectLevel(lvl) {
  state.level = lvl;
  const cardBeginner = document.getElementById('card-level-beginner');
  const cardPractitioner = document.getElementById('card-level-practitioner');

  if (lvl === 'beginner') {
    cardBeginner.className =
      'level-card cursor-pointer p-4 rounded-xl border-2 border-blue-600 bg-blue-50/40 shadow-sm ring-1 ring-blue-600 transition-all relative';
    cardBeginner.querySelector('.level-check')?.classList.remove('hidden');
    cardPractitioner.className =
      'level-card cursor-pointer p-4 rounded-xl border-2 border-slate-200 hover:border-slate-300 bg-white transition-all relative';
    cardPractitioner.querySelector('.level-check')?.classList.add('hidden');
  } else {
    cardPractitioner.className =
      'level-card cursor-pointer p-4 rounded-xl border-2 border-indigo-600 bg-indigo-50/40 shadow-sm ring-1 ring-indigo-600 transition-all relative';
    cardPractitioner.querySelector('.level-check')?.classList.remove('hidden');
    cardBeginner.className =
      'level-card cursor-pointer p-4 rounded-xl border-2 border-slate-200 hover:border-slate-300 bg-white transition-all relative';
    cardBeginner.querySelector('.level-check')?.classList.add('hidden');
  }
  refreshIcons();
}

// ==========================================
// 9. File Upload Handlers (Browser native)
// ==========================================
async function handleFileSelected(file) {
  const errorBox = document.getElementById('file-error-box');
  const errorMsg = document.getElementById('file-error-msg');
  const spinner = document.getElementById('file-extract-spinner');
  const dropzone = document.getElementById('dropzone');
  const fileCard = document.getElementById('uploaded-file-card');

  errorBox.classList.add('hidden');
  spinner.classList.remove('hidden');

  const ext = file.name.split('.').pop()?.toLowerCase();
  const allowed = ['pdf', 'pptx', 'docx', 'txt', 'md', 'markdown'];
  if (!ext || !allowed.includes(ext)) {
    errorMsg.textContent = `지원하지 않는 형식(.${ext})입니다. PDF, PPTX, DOCX, TXT 파일만 지원합니다.`;
    errorBox.classList.remove('hidden');
    spinner.classList.add('hidden');
    return;
  }

  // Auto-fill title if empty
  const titleInput = document.getElementById('input-title');
  if (!titleInput.value.trim()) {
    titleInput.value = file.name.replace(/\.[^/.]+$/, '');
  }

  try {
    const arrayBuffer = await file.arrayBuffer();
    const base64 = await fileToBase64(file);
    let extractedText = '';
    let isPdf = false;

    if (ext === 'docx') {
      extractedText = await extractTextFromDocxClient(arrayBuffer);
    } else if (ext === 'pptx') {
      extractedText = await extractTextFromPptxClient(arrayBuffer);
    } else if (ext === 'txt' || ext === 'md' || ext === 'markdown') {
      extractedText = new TextDecoder().decode(arrayBuffer);
    } else if (ext === 'pdf') {
      isPdf = true;
      extractedText = '[PDF 바이너리 문서 - Gemini 멀티모달 비전 엔진 분석 준비 완료]';
    }

    state.uploadedFile = {
      name: file.name,
      size: file.size,
      type: file.type || 'application/octet-stream',
      base64,
      extractedText,
      isPdf,
    };

    // Update UI
    dropzone.classList.add('hidden');
    fileCard.classList.remove('hidden');
    document.getElementById('uploaded-file-name').textContent = file.name;
    document.getElementById('uploaded-file-badge').textContent = ext.toUpperCase();
    document.getElementById('uploaded-file-meta').textContent = `${formatFileSize(file.size)} • ${
      isPdf ? 'PDF 비전 문서' : '텍스트 추출 완료'
    }`;
    document.getElementById('uploaded-char-count').textContent =
      (extractedText ? extractedText.length.toLocaleString() : '0') + ' 자';
    document.getElementById('uploaded-text-preview').textContent = extractedText.slice(0, 300) + '...';
  } catch (err) {
    console.error('File extraction error:', err);
    errorMsg.textContent = err.message || '파일 내용을 읽는 중 오류가 발생했습니다.';
    errorBox.classList.remove('hidden');
  } finally {
    spinner.classList.add('hidden');
    refreshIcons();
  }
}

function clearUploadedFile() {
  state.uploadedFile = null;
  document.getElementById('dropzone').classList.remove('hidden');
  document.getElementById('uploaded-file-card').classList.add('hidden');
  document.getElementById('file-input-hidden').value = '';
}

// ==========================================
// 10. AI Analysis Execution
// ==========================================
async function runAnalysis(e) {
  e.preventDefault();
  if (state.isAnalyzing) return;

  const title = document.getElementById('input-title').value.trim();
  if (!title) {
    alert('회의 자료명을 입력해 주세요.');
    return;
  }

  const rawText =
    state.inputMode === 'text'
      ? document.getElementById('textarea-rawtext').value.trim()
      : state.uploadedFile?.extractedText || '';

  const hasContent = (state.inputMode === 'text' && rawText.length > 0) || state.uploadedFile;
  if (!hasContent) {
    alert('회의 자료를 업로드하거나 텍스트를 입력해 주세요.');
    return;
  }

  state.isAnalyzing = true;
  const submitBtn = document.getElementById('btn-start-analyze');
  submitBtn.disabled = true;
  submitBtn.innerHTML = `
    <i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i>
    <span>회의 자료 AI 정밀 분석 중...</span>
  `;
  refreshIcons();

  const payload = {
    title,
    level: state.level,
    inputType: state.inputMode,
    rawText,
    fileData: state.uploadedFile
      ? {
          mimeType: state.uploadedFile.type,
          base64: state.uploadedFile.base64,
          fileName: state.uploadedFile.name,
        }
      : undefined,
  };

  try {
    let result = null;

    // 1. Try server endpoint first (/api/analyze)
    try {
      const res = await fetch('/api/analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      if (res.ok) {
        result = await res.json();
      } else {
        const errJson = await res.json().catch(() => ({}));
        if (res.status === 404) {
          // Running on static GitHub Pages
          result = null;
        } else {
          throw new Error(errJson.error || 'AI 분석 실패');
        }
      }
    } catch (apiErr) {
      if (!apiErr.message.includes('404')) {
        throw apiErr;
      }
    }

    // 2. If running on GitHub Pages where /api/analyze is not present
    if (!result) {
      if (state.settings.googleAppsScriptUrl) {
        // Call Apps Script Web App
        const gasRes = await callSheetsApi('analyze', payload);
        if (gasRes.success && gasRes.data) {
          result = gasRes.data;
        }
      }

      if (!result) {
        // Fallback demo generation for GitHub Pages static host preview
        result = createClientFallbackAnalysis(title, state.level, rawText);
      }
    }

    // Save and link
    state.analyses.unshift(result);
    state.activeAnalysis = result;
    saveAnalysesToStorage();

    // Auto sync to sheets if enabled and configured
    if (state.settings.autoSyncToSheets && state.settings.googleAppsScriptUrl) {
      callSheetsApi('saveAnalysis', { analysis: result }).then((syncRes) => {
        if (syncRes.success) {
          result.syncedToSheets = true;
          saveAnalysesToStorage();
          updateSheetsBadge();
        }
      });
    }

    updateBookmarksCount();
    updateSheetsBadge();
    switchScreen('result');
    showToast('회의 자료 맞춤 분석이 완료되었습니다!', 'success');
  } catch (error) {
    console.error('Analysis error:', error);
    showToast(error.message || '분석 중 오류가 발생했습니다. 다시 시도해 주세요.', 'error');
  } finally {
    state.isAnalyzing = false;
    submitBtn.disabled = false;
    submitBtn.innerHTML = `
      <i data-lucide="sparkles" class="w-4 h-4"></i>
      <span>AI 맞춤 분석 시작</span>
      <i data-lucide="arrow-right" class="w-4 h-4"></i>
    `;
    refreshIcons();
  }
}

// Fallback generator for purely static hosting when no backend is deployed
function createClientFallbackAnalysis(title, level, rawText) {
  const isBeginner = level === 'beginner';
  const id = 'MP-' + Date.now().toString(36).toUpperCase();

  // Extract terms appearing in rawText
  const extractedTerms = [];
  const candidateKeywords = [
    { term: 'MSA', def: '독립적인 작은 서비스들의 집합으로 소프트웨어를 구성하는 설계 방식' },
    { term: 'gRPC', def: '구글이 개발한 고성능 오픈소스 원격 프로시저 호출 프레임워크' },
    { term: 'Istio', def: '마이크로서비스 간의 통신과 보안을 제어하는 서비스 메시 도구' },
    { term: 'Kafka', def: '대규모 실시간 데이터 스트림을 안전하게 중계하는 메시지 브로커' },
    { term: '사가 패턴', def: '분산 트랜잭션 환경에서 데이터 일관성을 유지하는 설계 패턴' },
    { term: '멱등성', def: '연산을 여러 번 적용해도 결과가 달라지지 않는 성질' },
    { term: 'BFF', def: '클라이언트 특성에 맞게 API를 가공하는 프론트엔드 전용 백엔드' },
    { term: '3D Secure', def: '카드 결제 시 부정 사용을 방지하는 전자상거래 인증 프로토콜' },
    { term: '에스크로', def: '거래 대금을 제3자에게 임시 예치하여 안전을 보장하는 방식' },
    { term: 'RAG', def: '외부 지식 베이스를 검색하여 LLM 응답의 정확도를 높이는 기법' },
  ];

  let termIndex = 1;
  candidateKeywords.forEach((k) => {
    if (rawText.includes(k.term) || extractedTerms.length < 3) {
      extractedTerms.push({
        termId: `${id}-T${String(termIndex++).padStart(2, '0')}`,
        analysisId: id,
        term: k.term,
        definition: k.def,
        easyExplanation: isBeginner
          ? `${k.term}은(는) 신입사원도 쉽게 이해할 수 있도록 일상 비유로 설명하면, 거대한 시스템을 모듈식 레고 블록처럼 분리하여 관리하는 기술입니다.`
          : `${k.term}은(는) 실무 시스템 아키텍처 상에서 컴포넌트 간 결합도를 낮추고 데이터 신뢰성을 보장하기 위한 표준 스펙입니다.`,
        contextualMeaning: `이번 '${title}' 회의에서 핵심 시스템 구성 및 아키텍처 연동 요소로 검토됨`,
        practicalUseCases: '실제 프로덕션 환경의 트래픽 급증 및 장애 격리 시 필수 활용',
        sourceQuote: rawText.slice(0, 80) + '...',
        levelNote: isBeginner ? '일상 비유: 백화점 대신 전문 상점가 구성' : '실무 스펙: 네트워크 지연 최소화 및 멱등키 보장',
        bookmarked: false,
      });
    }
  });

  return {
    analysisId: id,
    analyzedAt: new Date().toISOString(),
    title,
    inputType: '회의 자료 분석',
    level,
    levelLabel: isBeginner ? '입문자' : '실무자',
    summary: {
      topic: `${title} 핵심 안건 검토 및 의사결정`,
      purpose: rawText.includes('목적') ? '안정적인 서비스 전환 및 인프라 표준화' : '자료에서 확인되지 않음',
      purposeConfirmed: rawText.includes('목적'),
      keyDiscussionItems: [
        {
          title: '핵심 아키텍처 및 통신 표준화',
          description: '효율적인 서비스 간 연계 및 병목 해소를 위한 주요 기술 채택 논의',
          originalQuote: rawText.slice(0, 60),
          isFromDocument: true,
        },
        {
          title: '분산 환경에서의 장애 대응 및 데이터 정합성',
          description: '네트워크 장애 시 중복 방지 및 롤백을 위한 안전장치 마련',
          originalQuote: '네트워크 순단이나 컨슈머 재시도에 대비',
          isFromDocument: true,
        },
      ],
      technicalConcepts: [
        { name: '서비스 메시', role: '서비스 간 보안 통신 및 트래픽 라우팅 통제' },
        { name: '이벤트 파이프라인', role: '비동기 메시지 발행 및 분산 트랜잭션 보장' },
      ],
      keyPoints: [
        { point: '단일 장애점(SPOF) 해소', details: '한 모듈이 다운되어도 전체 서비스가 멈추지 않음', isFact: true },
        { point: '멱등성 보장 필수', details: '중복 결제 및 중복 요청 방지 메커니즘 구축', isFact: true },
      ],
    },
    backgroundKnowledge: {
      basicConcepts: [
        { concept: '마이크로서비스 (MSA)', explanation: '작은 독립 단위로 서비스를 쪼개어 배포하는 구조', relevance: '이번 회의의 전체 목표 구조' },
        { concept: '트랜잭션 정합성', explanation: '여러 단계의 작업이 모두 성공하거나 모두 취소되어야 하는 일관성', relevance: '데이터 불일치 방지' },
      ],
      workProcesses: [
        { stage: '아키텍처 설계', processDescription: '도메인 분리 및 서비스 간 통신 프로토콜 결정' },
        { stage: '검증 및 테스트', processDescription: '부하 테스트 및 장애 복구 시나리오 검증' },
      ],
      techRelationships: [
        { entityA: '마이크로서비스', relationship: '통신 제어 및 보안 담당', entityB: '서비스 메시' },
      ],
      recommendedPrep: ['단일 시스템과 분산 시스템의 차이점을 파악해 두면 논의 이해가 훨씬 수월합니다.'],
    },
    terminology: extractedTerms,
    expectedQA: [
      {
        question: '이 기술을 도입했을 때 기존 시스템 대비 가장 큰 장점은 무엇인가요?',
        suggestedAnswer: '서비스별 독립 배포가 가능해져 특정 기능 장애가 전체로 번지지 않고 빠른 업데이트가 가능합니다.',
        answerBasis: 'document',
        sourceExcerpt: '단일 장애점(SPOF)을 해소하기 위해 독립 서비스로 분리',
      },
      {
        question: '회의 중 예상되는 실무 리스크는 어떻게 대응하나요?',
        suggestedAnswer: '네트워크 분산 환경에서 발생할 수 있는 데이터 불일치는 멱등키와 보상 트랜잭션으로 방어합니다.',
        answerBasis: 'inference',
        sourceExcerpt: '실무 분산 환경 베스트 프랙티스 기반 추론',
      },
    ],
  };
}

// ==========================================
// 11. Result Screen Rendering
// ==========================================
function renderResultScreen() {
  const a = state.activeAnalysis;
  if (!a) return;

  const isBeginner = a.level === 'beginner';

  // Header Metadata
  document.getElementById('res-title').textContent = a.title;
  document.getElementById('res-badge-id').textContent = a.analysisId;
  document.getElementById('res-badge-inputtype').textContent = `입력 방식: ${a.inputType}`;
  document.getElementById('res-date').textContent = new Date(a.analyzedAt).toLocaleString('ko-KR', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });

  const levelBadge = document.getElementById('res-badge-level');
  levelBadge.textContent = isBeginner ? '🌱 입문자 모드 (비유와 쉬운 예시)' : '⚡ 실무자 모드 (스펙과 기술 차이)';
  levelBadge.className = isBeginner
    ? 'inline-flex items-center gap-1 text-xs font-bold px-2.5 py-0.5 rounded-full bg-blue-100 text-blue-800 border border-blue-200'
    : 'inline-flex items-center gap-1 text-xs font-bold px-2.5 py-0.5 rounded-full bg-indigo-100 text-indigo-800 border border-indigo-200';

  const termsCount = a.terminology?.length || 0;
  document.getElementById('res-terms-stat').textContent = `전문용어 ${termsCount}개 추출`;
  document.getElementById('res-terms-count-badge').textContent = termsCount;
  document.getElementById('res-discussions-stat').textContent = `핵심 논의 ${a.summary?.keyDiscussionItems?.length || 0}건`;

  // Sync Button Status
  const syncBtnText = document.getElementById('sync-sheets-btn-text');
  if (a.syncedToSheets) {
    syncBtnText.textContent = 'Google Sheets 저장됨';
  } else {
    syncBtnText.textContent = state.settings.googleAppsScriptUrl ? 'Google Sheets에 저장' : 'Google Sheets 연동하기';
  }

  // Render TAB 1: Summary
  document.getElementById('res-summary-topic').textContent = a.summary.topic;
  document.getElementById('res-summary-purpose').textContent = a.summary.purpose;
  const purposeBadge = document.getElementById('res-summary-purpose-badge');
  if (a.summary.purposeConfirmed) {
    purposeBadge.textContent = '자료 확인 사실';
    purposeBadge.className = 'text-[11px] px-2 py-0.5 rounded font-semibold border bg-emerald-50 text-emerald-700 border-emerald-200';
  } else {
    purposeBadge.textContent = '자료에서 확인되지 않음';
    purposeBadge.className = 'text-[11px] px-2 py-0.5 rounded font-semibold border bg-amber-50 text-amber-700 border-amber-200';
  }

  document.getElementById('res-discussions-count-chip').textContent = `${a.summary.keyDiscussionItems?.length || 0}개 안건`;
  document.getElementById('res-discussions-list').innerHTML = (a.summary.keyDiscussionItems || [])
    .map(
      (item, idx) => `
    <div class="p-4 rounded-xl border border-slate-200 bg-slate-50/40 hover:bg-slate-50 transition-colors">
      <div class="flex items-start justify-between gap-3 mb-1.5">
        <div class="flex items-center gap-2">
          <span class="w-5 h-5 rounded-full bg-blue-600 text-white text-xs font-bold flex items-center justify-center shrink-0">
            ${idx + 1}
          </span>
          <h4 class="font-bold text-sm text-slate-900">${escapeHtml(item.title)}</h4>
        </div>
        <span class="text-[11px] font-medium px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200 shrink-0 flex items-center gap-1">
          <i data-lucide="shield-check" class="w-3 h-3"></i>
          자료 확인
        </span>
      </div>
      <p class="text-xs sm:text-sm text-slate-700 leading-relaxed pl-7">${escapeHtml(item.description)}</p>
      ${
        item.originalQuote
          ? `
        <div class="mt-2.5 ml-7 p-2.5 rounded-lg bg-white border border-slate-200 text-xs text-slate-600 flex items-start gap-2">
          <i data-lucide="quote" class="w-3.5 h-3.5 text-blue-500 shrink-0 mt-0.5"></i>
          <div class="italic text-slate-700">
            <span class="font-semibold text-slate-500 not-italic mr-1">관련 원문:</span>
            "${escapeHtml(item.originalQuote)}"
          </div>
        </div>
      `
          : ''
      }
    </div>
  `
    )
    .join('');

  document.getElementById('res-keypoints-list').innerHTML = (a.summary.keyPoints || [])
    .map(
      (kp) => `
    <div class="p-3 rounded-xl bg-amber-50/40 border border-amber-200/60 text-xs">
      <div class="font-bold text-slate-900 mb-1 flex items-center gap-1.5">
        <span class="text-amber-600">✓</span>
        <span>${escapeHtml(kp.point)}</span>
      </div>
      <p class="text-slate-600 pl-4">${escapeHtml(kp.details)}</p>
    </div>
  `
    )
    .join('');

  document.getElementById('res-techconcepts-list').innerHTML = (a.summary.technicalConcepts || [])
    .map(
      (tc) => `
    <div class="p-2.5 rounded-lg border border-slate-200 bg-white text-xs">
      <div class="font-bold text-blue-800">${escapeHtml(tc.name)}</div>
      <div class="text-slate-600 mt-0.5">${escapeHtml(tc.role)}</div>
    </div>
  `
    )
    .join('');

  // Render TAB 2: Background
  document.getElementById('res-bg-intro-desc').textContent = isBeginner
    ? '신입사원도 쉽게 파악할 수 있도록 기초 정의와 이번 회의와의 연결고리를 정리했습니다.'
    : '실무자가 회의 안건의 기술적 맥락을 빠르게 조망할 수 있는 핵심 선수 지식입니다.';

  document.getElementById('res-basicconcepts-list').innerHTML = (a.backgroundKnowledge.basicConcepts || [])
    .map(
      (c) => `
    <div class="p-4 rounded-xl border border-slate-200 bg-slate-50/50 hover:bg-white hover:shadow-sm transition-all">
      <div class="font-bold text-sm text-slate-900 mb-1 flex items-center justify-between">
        <span>${escapeHtml(c.concept)}</span>
        <span class="text-[10px] px-1.5 py-0.5 rounded bg-blue-100 text-blue-700 font-semibold">기초 지식</span>
      </div>
      <p class="text-xs text-slate-700 leading-relaxed mb-2">${escapeHtml(c.explanation)}</p>
      <div class="text-[11px] text-blue-800 bg-blue-50/70 p-2 rounded-lg border border-blue-100">
        <span class="font-bold mr-1">회의 연관성:</span>${escapeHtml(c.relevance)}
      </div>
    </div>
  `
    )
    .join('');

  document.getElementById('res-workprocesses-list').innerHTML = (a.backgroundKnowledge.workProcesses || [])
    .map(
      (p) => `
    <div class="p-3.5 rounded-xl border border-slate-200 bg-white flex items-start gap-3.5">
      <div class="px-2.5 py-1 rounded-lg bg-indigo-50 border border-indigo-200 text-indigo-700 font-bold text-xs shrink-0 mt-0.5">
        ${escapeHtml(p.stage)}
      </div>
      <p class="text-xs sm:text-sm text-slate-700 leading-relaxed">${escapeHtml(p.processDescription)}</p>
    </div>
  `
    )
    .join('');

  document.getElementById('res-techrelationships-list').innerHTML = (a.backgroundKnowledge.techRelationships || [])
    .map(
      (r) => `
    <div class="p-3 rounded-xl border border-slate-200 bg-slate-50/50 text-xs">
      <div class="flex items-center gap-2 font-bold text-slate-800 mb-1">
        <span class="text-blue-700">${escapeHtml(r.entityA)}</span>
        <i data-lucide="arrow-right" class="w-3 h-3 text-slate-400"></i>
        <span class="text-indigo-700">${escapeHtml(r.entityB)}</span>
      </div>
      <p class="text-slate-600 pl-1">${escapeHtml(r.relationship)}</p>
    </div>
  `
    )
    .join('');

  document.getElementById('res-prepchecklist-list').innerHTML = (a.backgroundKnowledge.recommendedPrep || [])
    .map(
      (tip) => `
    <div class="p-3 rounded-xl bg-blue-50/50 border border-blue-200/60 text-xs text-slate-800 flex items-start gap-2">
      <span class="font-bold text-blue-600">•</span>
      <span>${escapeHtml(tip)}</span>
    </div>
  `
    )
    .join('');

  // Render TAB 3: Terminology
  renderTerminologyList();

  // Render TAB 4: Q&A
  document.getElementById('res-qa-list').innerHTML = (a.expectedQA || [])
    .map(
      (qa, idx) => `
    <div class="p-4 rounded-xl border border-slate-200 bg-slate-50/40 hover:bg-slate-50 transition-colors">
      <div class="flex items-start justify-between gap-3 mb-2">
        <div class="flex items-center gap-2">
          <span class="w-6 h-6 rounded-full bg-blue-100 text-blue-700 text-xs font-bold flex items-center justify-center shrink-0">
            Q${idx + 1}
          </span>
          <h4 class="font-bold text-sm text-slate-900">${escapeHtml(qa.question)}</h4>
        </div>
        <span class="text-[11px] font-semibold px-2 py-0.5 rounded-full border shrink-0 ${
          qa.answerBasis === 'document'
            ? 'bg-emerald-50 text-emerald-700 border-emerald-200'
            : 'bg-amber-50 text-amber-700 border-amber-200'
        }">
          ${qa.answerBasis === 'document' ? '자료 근거 답변' : '실무 추론 답변'}
        </span>
      </div>
      <div class="pl-8 text-xs sm:text-sm text-slate-700 leading-relaxed">
        <div class="font-semibold text-slate-900 mb-0.5">예상 답변 / 검토 방향:</div>
        <p>${escapeHtml(qa.suggestedAnswer)}</p>
      </div>
      ${
        qa.sourceExcerpt
          ? `
        <div class="mt-2.5 ml-8 p-2 rounded-lg bg-white border border-slate-200 text-xs text-slate-500">
          <strong class="text-slate-600 mr-1">근거/배경:</strong>
          ${escapeHtml(qa.sourceExcerpt)}
        </div>
      `
          : ''
      }
    </div>
  `
    )
    .join('');

  refreshIcons();
}

function renderTerminologyList() {
  const a = state.activeAnalysis;
  if (!a) return;

  const isBeginner = a.level === 'beginner';
  const filterKw = state.termSearchKeyword.toLowerCase();

  const filtered = (a.terminology || []).filter(
    (t) =>
      t.term.toLowerCase().includes(filterKw) ||
      t.definition.toLowerCase().includes(filterKw) ||
      t.easyExplanation.toLowerCase().includes(filterKw)
  );

  document.getElementById('res-term-count-num').textContent = filtered.length;
  document.getElementById('res-term-mode-note').textContent = isBeginner
    ? '입문자 맞춤형 비유 적용'
    : '실무자 스펙 및 차이점 분석';

  const container = document.getElementById('res-terms-grid');
  if (filtered.length === 0) {
    container.innerHTML = `
      <div class="col-span-2 p-12 text-center bg-white rounded-2xl border border-slate-200 text-slate-400">
        <i data-lucide="search" class="w-8 h-8 mx-auto mb-2 opacity-50"></i>
        <p class="text-sm font-semibold">검색된 전문용어가 없습니다.</p>
      </div>
    `;
    refreshIcons();
    return;
  }

  container.innerHTML = filtered
    .map((term) => {
      const isExpanded = state.expandedTermId === term.termId;

      return `
      <div class="bg-white rounded-2xl border transition-all ${
        term.bookmarked ? 'border-amber-300 ring-1 ring-amber-300 shadow-sm' : 'border-slate-200 hover:border-slate-300 shadow-sm'
      }">
        <div class="p-5">
          <div class="flex items-start justify-between gap-2 mb-2">
            <div>
              <div class="flex items-center gap-2">
                <h4 class="font-extrabold text-base text-slate-900">${escapeHtml(term.term)}</h4>
                <span class="text-[10px] font-mono px-1.5 py-0.5 rounded bg-slate-100 text-slate-500">${term.termId}</span>
              </div>
              <p class="text-xs text-slate-500 mt-0.5 leading-snug">${escapeHtml(term.definition)}</p>
            </div>
            <button
              type="button"
              data-bookmark-termid="${term.termId}"
              class="btn-toggle-term-bookmark p-2 rounded-xl transition-all shrink-0 ${
                term.bookmarked
                  ? 'bg-amber-100 text-amber-600 hover:bg-amber-200'
                  : 'bg-slate-100 text-slate-400 hover:text-amber-500 hover:bg-slate-200'
              }"
              title="${term.bookmarked ? '북마크 해제' : '개인 용어 사전에 북마크'}"
            >
              <i data-lucide="${term.bookmarked ? 'bookmark-check' : 'bookmark'}" class="w-4 h-4 ${term.bookmarked ? 'fill-amber-500 text-amber-600' : ''}"></i>
            </button>
          </div>

          <div class="p-3 rounded-xl text-xs mb-3 leading-relaxed ${
            isBeginner ? 'bg-blue-50/60 border border-blue-200/60 text-blue-950' : 'bg-indigo-50/60 border border-indigo-200/60 text-indigo-950'
          }">
            <div class="font-bold flex items-center gap-1 mb-1">
              <i data-lucide="sparkles" class="w-3.5 h-3.5 text-blue-600"></i>
              <span>${isBeginner ? '쉬운 맞춤 설명 (입문자)' : '실무 스펙 및 정의 (실무자)'}</span>
            </div>
            <p>${escapeHtml(term.easyExplanation)}</p>
          </div>

          ${
            term.levelNote
              ? `
            <div class="text-[11px] text-slate-600 bg-slate-50 p-2.5 rounded-lg border border-slate-200 mb-3 flex items-start gap-1.5">
              <span class="font-bold text-slate-700 shrink-0">${isBeginner ? '💡 비유:' : '🔍 실무 포인트:'}</span>
              <span>${escapeHtml(term.levelNote)}</span>
            </div>
          `
              : ''
          }

          <div class="space-y-1.5 text-xs text-slate-700">
            <div>
              <strong class="text-slate-900 block font-semibold mb-0.5">회의 문맥에서의 의미:</strong>
              <p class="text-slate-600 leading-relaxed">${escapeHtml(term.contextualMeaning)}</p>
            </div>
            <div class="pt-1.5">
              <strong class="text-slate-900 block font-semibold mb-0.5">실무 활용 사례:</strong>
              <p class="text-slate-600 leading-relaxed">${escapeHtml(term.practicalUseCases)}</p>
            </div>
          </div>

          ${
            term.sourceQuote
              ? `
            <div class="mt-3 pt-3 border-t border-slate-100">
              <button
                type="button"
                data-quote-termid="${term.termId}"
                class="btn-toggle-quote flex items-center justify-between w-full text-[11px] text-slate-500 hover:text-slate-800 font-medium"
              >
                <span class="flex items-center gap-1">
                  <i data-lucide="quote" class="w-3 h-3 text-blue-500"></i>
                  자료 원문 문맥 보기
                </span>
                <i data-lucide="${isExpanded ? 'chevron-up' : 'chevron-down'}" class="w-3.5 h-3.5"></i>
              </button>
              ${
                isExpanded
                  ? `
                <div class="mt-2 p-2.5 rounded-lg bg-slate-50 border border-slate-200 text-xs text-slate-600 italic">
                  "${escapeHtml(term.sourceQuote)}"
                </div>
              `
                  : ''
              }
            </div>
          `
              : ''
          }
        </div>
      </div>
    `;
    })
    .join('');

  // Attach event listeners for bookmark buttons
  container.querySelectorAll('.btn-toggle-term-bookmark').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const termId = btn.getAttribute('data-bookmark-termid');
      toggleTermBookmark(termId);
    });
  });

  // Attach event listeners for quote expand buttons
  container.querySelectorAll('.btn-toggle-quote').forEach((btn) => {
    btn.addEventListener('click', () => {
      const termId = btn.getAttribute('data-quote-termid');
      state.expandedTermId = state.expandedTermId === termId ? null : termId;
      renderTerminologyList();
    });
  });

  refreshIcons();
}

function toggleTermBookmark(termId) {
  if (!state.activeAnalysis) return;
  const term = state.activeAnalysis.terminology.find((t) => t.termId === termId);
  if (!term) return;

  term.bookmarked = !term.bookmarked;
  saveAnalysesToStorage();
  updateBookmarksCount();
  renderTerminologyList();

  if (state.settings.googleAppsScriptUrl) {
    callSheetsApi('toggleBookmark', { termId, bookmarked: term.bookmarked });
  }

  showToast(term.bookmarked ? `'${term.term}' 용어가 북마크되었습니다.` : '북마크가 해제되었습니다.', 'info');
}

function selectResultSubTab(tabName) {
  state.activeResSubTab = tabName;

  const tabs = ['summary', 'background', 'terminology', 'qa'];
  tabs.forEach((t) => {
    const btn = document.getElementById(`res-tab-btn-${t}`);
    const pane = document.getElementById(`res-pane-${t}`);
    if (t === tabName) {
      btn.className =
        'res-tab-btn flex items-center gap-2 px-5 py-3 text-sm font-bold border-b-2 border-blue-600 text-blue-600 whitespace-nowrap transition-colors';
      pane.classList.remove('hidden');
    } else {
      btn.className =
        'res-tab-btn flex items-center gap-2 px-5 py-3 text-sm font-bold border-b-2 border-transparent text-slate-500 hover:text-slate-900 whitespace-nowrap transition-colors';
      pane.classList.add('hidden');
    }
  });
  refreshIcons();
}

// Copy Markdown
function copyReportMarkdown() {
  const a = state.activeAnalysis;
  if (!a) return;

  let md = `# [MeetPrep AI 분석 결과] ${a.title}\n`;
  md += `- 분석ID: ${a.analysisId}\n`;
  md += `- 분석일시: ${new Date(a.analyzedAt).toLocaleString('ko-KR')}\n`;
  md += `- 설명수준: ${a.levelLabel} 모드\n\n`;

  md += `## 1. 회의 핵심 요약\n`;
  md += `**회의 주제:** ${a.summary.topic}\n\n`;
  md += `**회의 목적:** ${a.summary.purpose} (${a.summary.purposeConfirmed ? '자료 확인' : '자료 미확인'})\n\n`;

  md += `### 주요 논의 사항\n`;
  (a.summary.keyDiscussionItems || []).forEach((item, idx) => {
    md += `${idx + 1}. **${item.title}**: ${item.description}\n`;
    if (item.originalQuote) md += `   > 원문: "${item.originalQuote}"\n`;
  });

  md += `\n## 2. 알아야 할 배경지식\n`;
  (a.backgroundKnowledge.basicConcepts || []).forEach((c) => {
    md += `- **${c.concept}**: ${c.explanation} (관련성: ${c.relevance})\n`;
  });

  md += `\n## 3. 맞춤형 전문용어 사전 (${a.levelLabel} 모드)\n`;
  (a.terminology || []).forEach((t) => {
    md += `### ${t.term}\n`;
    md += `- **정의**: ${t.definition}\n`;
    md += `- **맞춤 설명**: ${t.easyExplanation}\n`;
    md += `- **회의 문맥**: ${t.contextualMeaning}\n`;
    md += `- **실무 사례**: ${t.practicalUseCases}\n\n`;
  });

  navigator.clipboard.writeText(md);
  const textSpan = document.getElementById('copy-markdown-text');
  textSpan.textContent = '복사 완료!';
  setTimeout(() => {
    textSpan.textContent = '마크다운 복사';
  }, 2000);
}

// ==========================================
// 12. History Screen Rendering
// ==========================================
function renderHistoryScreen() {
  const container = document.getElementById('history-cards-container');
  const emptyView = document.getElementById('history-empty-view');
  const totalCount = document.getElementById('history-total-count');

  totalCount.textContent = state.analyses.length;

  const kw = state.historySearchKeyword.toLowerCase();
  const filtered = state.analyses.filter((item) => {
    const matchesSearch =
      item.title.toLowerCase().includes(kw) ||
      item.summary.topic.toLowerCase().includes(kw) ||
      (item.terminology || []).some((t) => t.term.toLowerCase().includes(kw));
    const matchesLevel = state.historyFilterLevel === 'all' || item.level === state.historyFilterLevel;
    return matchesSearch && matchesLevel;
  });

  if (filtered.length === 0) {
    container.innerHTML = '';
    emptyView.classList.remove('hidden');
    refreshIcons();
    return;
  }

  emptyView.classList.add('hidden');
  container.innerHTML = filtered
    .map((item) => {
      const isBeginner = item.level === 'beginner';

      return `
      <div
        data-history-id="${item.analysisId}"
        class="history-card bg-white rounded-2xl border border-slate-200 hover:border-blue-300 hover:shadow-md transition-all p-5 flex flex-col justify-between group cursor-pointer"
      >
        <div>
          <div class="flex items-center justify-between gap-2 mb-2">
            <div class="flex items-center gap-1.5">
              <span class="text-[11px] font-bold px-2.5 py-0.5 rounded-full ${
                isBeginner ? 'bg-blue-50 text-blue-700 border border-blue-200' : 'bg-indigo-50 text-indigo-700 border border-indigo-200'
              }">
                ${isBeginner ? '🌱 입문자 모드' : '⚡ 실무자 모드'}
              </span>
              <span class="text-[10px] font-mono px-2 py-0.5 rounded bg-slate-100 text-slate-500">${item.analysisId}</span>
            </div>
            <div class="flex items-center gap-1.5">
              ${
                item.syncedToSheets
                  ? `
                <span class="flex items-center gap-1 text-[10px] font-medium text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-full border border-emerald-200">
                  <i data-lucide="database" class="w-3 h-3"></i>
                  시트 연동됨
                </span>
              `
                  : ''
              }
              <button
                type="button"
                data-delete-id="${item.analysisId}"
                class="btn-delete-history p-1 rounded-lg text-slate-400 hover:text-red-600 hover:bg-red-50 transition-colors"
                title="기록 삭제"
              >
                <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
              </button>
            </div>
          </div>

          <h3 class="font-bold text-base text-slate-900 group-hover:text-blue-600 transition-colors line-clamp-1 mb-1.5">
            ${escapeHtml(item.title)}
          </h3>

          <p class="text-xs text-slate-600 line-clamp-2 leading-relaxed mb-3">
            ${escapeHtml(item.summary.topic)}
          </p>

          <div class="flex flex-wrap gap-1 mb-4">
            ${(item.terminology || [])
              .slice(0, 4)
              .map(
                (t) => `
              <span class="text-[10px] px-2 py-0.5 rounded-md bg-slate-100 text-slate-700 font-medium border border-slate-200">
                ${escapeHtml(t.term)}
              </span>
            `
              )
              .join('')}
            ${
              (item.terminology || []).length > 4
                ? `
              <span class="text-[10px] px-1.5 py-0.5 rounded-md bg-slate-50 text-slate-400">
                +${item.terminology.length - 4}개 더
              </span>
            `
                : ''
            }
          </div>
        </div>

        <div class="pt-3 border-t border-slate-100 flex items-center justify-between text-xs text-slate-400">
          <div class="flex items-center gap-1.5">
            <i data-lucide="clock" class="w-3.5 h-3.5"></i>
            <span>${new Date(item.analyzedAt).toLocaleDateString('ko-KR', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' })}</span>
          </div>
          <span class="inline-flex items-center gap-1 text-blue-600 font-bold group-hover:translate-x-0.5 transition-transform">
            <span>결과 보기</span>
            <i data-lucide="arrow-right" class="w-3.5 h-3.5"></i>
          </span>
        </div>
      </div>
    `;
    })
    .join('');

  // Attach card click handlers
  container.querySelectorAll('.history-card').forEach((card) => {
    card.addEventListener('click', (e) => {
      if (e.target.closest('.btn-delete-history')) return;
      const id = card.getAttribute('data-history-id');
      const target = state.analyses.find((a) => a.analysisId === id);
      if (target) {
        state.activeAnalysis = target;
        switchScreen('result');
      }
    });
  });

  // Attach delete handlers
  container.querySelectorAll('.btn-delete-history').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const id = btn.getAttribute('data-delete-id');
      const target = state.analyses.find((a) => a.analysisId === id);
      if (target && confirm(`'${target.title}' 분석 기록을 삭제하시겠습니까?`)) {
        state.analyses = state.analyses.filter((a) => a.analysisId !== id);
        if (state.activeAnalysis?.analysisId === id) {
          state.activeAnalysis = state.analyses[0] || null;
        }
        saveAnalysesToStorage();
        updateBookmarksCount();
        renderHistoryScreen();
        showToast('분석 기록이 삭제되었습니다.', 'info');

        if (state.settings.googleAppsScriptUrl) {
          callSheetsApi('deleteAnalysis', { analysisId: id });
        }
      }
    });
  });

  refreshIcons();
}

// ==========================================
// 13. Bookmarks Modal (Personal Glossary)
// ==========================================
function getAllBookmarkedTerms() {
  const bookmarks = [];
  state.analyses.forEach((analysis) => {
    (analysis.terminology || []).forEach((t) => {
      if (t.bookmarked) {
        bookmarks.push({ term: t, analysisTitle: analysis.title, analysisId: analysis.analysisId });
      }
    });
  });
  return bookmarks;
}

function updateBookmarksCount() {
  const count = getAllBookmarkedTerms().length;
  const badge = document.getElementById('badge-bookmarks-count');
  if (count > 0) {
    badge.textContent = count;
    badge.classList.remove('hidden');
  } else {
    badge.classList.add('hidden');
  }
}

function renderBookmarksModal() {
  const list = getAllBookmarkedTerms();
  document.getElementById('modal-bookmarks-total').textContent = list.length;

  const kw = state.bookmarksSearchKeyword.toLowerCase();
  const filtered = list.filter(
    (b) =>
      b.term.term.toLowerCase().includes(kw) ||
      b.term.definition.toLowerCase().includes(kw) ||
      b.term.easyExplanation.toLowerCase().includes(kw) ||
      b.analysisTitle.toLowerCase().includes(kw)
  );

  const container = document.getElementById('modal-bookmarks-list');
  if (filtered.length === 0) {
    container.innerHTML = `
      <div class="p-12 text-center text-slate-400">
        <i data-lucide="bookmark" class="w-8 h-8 mx-auto mb-2 opacity-50"></i>
        <p class="text-sm font-semibold">${kw ? '검색된 북마크 용어가 없습니다.' : '아직 북마크한 전문용어가 없습니다.'}</p>
        <p class="text-xs mt-1">회의 분석 결과 화면의 용어 카드에서 북마크 아이콘을 누르면 이곳에 모아집니다.</p>
      </div>
    `;
    refreshIcons();
    return;
  }

  container.innerHTML = filtered
    .map(
      ({ term, analysisTitle, analysisId }) => `
    <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm hover:border-amber-300 transition-colors">
      <div class="flex items-start justify-between gap-3 mb-2">
        <div>
          <div class="flex items-center gap-2">
            <h4 class="font-bold text-base text-slate-900">${escapeHtml(term.term)}</h4>
            <span class="text-[10px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-500 font-mono">${term.termId}</span>
          </div>
          <p class="text-xs text-slate-600 mt-0.5">${escapeHtml(term.definition)}</p>
        </div>
        <button
          type="button"
          data-unbookmark-aid="${analysisId}"
          data-unbookmark-tid="${term.termId}"
          class="btn-unbookmark-modal p-1.5 rounded-lg bg-amber-50 text-amber-600 hover:bg-red-50 hover:text-red-600 transition-colors"
          title="북마크 해제"
        >
          <i data-lucide="bookmark-check" class="w-4 h-4 fill-amber-500"></i>
        </button>
      </div>

      <div class="p-2.5 rounded-lg bg-amber-50/40 border border-amber-100 text-xs text-amber-950 mb-2 leading-relaxed">
        <span class="font-bold mr-1">맞춤형 설명:</span>
        ${escapeHtml(term.easyExplanation)}
      </div>

      ${
        term.practicalUseCases
          ? `
        <div class="text-xs text-slate-600 mb-2">
          <span class="font-semibold text-slate-800 mr-1">실무 활용:</span>
          ${escapeHtml(term.practicalUseCases)}
        </div>
      `
          : ''
      }

      <div class="pt-2 border-t border-slate-100 flex items-center justify-between text-[11px] text-slate-400">
        <span class="truncate max-w-xs sm:max-w-md">출처 회의: ${escapeHtml(analysisTitle)}</span>
        <button
          type="button"
          data-goto-analysis="${analysisId}"
          class="btn-goto-analysis text-blue-600 hover:underline font-semibold flex items-center gap-1"
        >
          <span>회의 결과 열기</span>
          <i data-lucide="arrow-right" class="w-3 h-3"></i>
        </button>
      </div>
    </div>
  `
    )
    .join('');

  container.querySelectorAll('.btn-unbookmark-modal').forEach((btn) => {
    btn.addEventListener('click', () => {
      const aId = btn.getAttribute('data-unbookmark-aid');
      const tId = btn.getAttribute('data-unbookmark-tid');
      const a = state.analyses.find((item) => item.analysisId === aId);
      if (a) {
        const t = a.terminology.find((item) => item.termId === tId);
        if (t) {
          t.bookmarked = false;
          saveAnalysesToStorage();
          updateBookmarksCount();
          renderBookmarksModal();
          if (state.activeAnalysis?.analysisId === aId) {
            renderTerminologyList();
          }
          if (state.settings.googleAppsScriptUrl) {
            callSheetsApi('toggleBookmark', { termId: tId, bookmarked: false });
          }
        }
      }
    });
  });

  container.querySelectorAll('.btn-goto-analysis').forEach((btn) => {
    btn.addEventListener('click', () => {
      const aId = btn.getAttribute('data-goto-analysis');
      const a = state.analyses.find((item) => item.analysisId === aId);
      if (a) {
        state.activeAnalysis = a;
        document.getElementById('modal-bookmarks').classList.add('hidden');
        switchScreen('result');
      }
    });
  });

  refreshIcons();
}

// ==========================================
// 14. Google Sheets Settings Modal
// ==========================================
function updateSheetsBadge() {
  const iconSpan = document.getElementById('sheets-status-icon');
  if (state.settings.googleAppsScriptUrl) {
    iconSpan.innerHTML = '<i data-lucide="check-circle-2" class="w-3.5 h-3.5 text-emerald-600"></i>';
  } else {
    iconSpan.innerHTML = '<i data-lucide="alert-circle" class="w-3.5 h-3.5 text-slate-400"></i>';
  }
  refreshIcons();
}

function renderSheetsGuideSteps() {
  const container = document.getElementById('modal-guide-steps-container');
  if (!container) return;

  container.innerHTML = APPS_SCRIPT_GUIDE_STEPS.map(
    (s) => `
    <div class="p-3 rounded-xl border border-slate-200 bg-white flex items-start gap-3">
      <span class="w-6 h-6 rounded-full bg-slate-900 text-white text-xs font-bold flex items-center justify-center shrink-0 mt-0.5">
        ${s.step}
      </span>
      <div>
        <h5 class="font-bold text-xs text-slate-900">${escapeHtml(s.title)}</h5>
        <p class="text-xs text-slate-500 mt-0.5 leading-relaxed">${escapeHtml(s.description)}</p>
      </div>
    </div>
  `
  ).join('');

  document.getElementById('gas-code-preview').textContent =
    GOOGLE_APPS_SCRIPT_CODE.slice(0, 480) + '\n... [상단 코드 복사 버튼을 누르면 전체 코드가 복사됩니다] ...';
}

async function testSheetsConnection() {
  const urlInput = document.getElementById('input-gas-url');
  const url = urlInput.value.trim();
  const resultBox = document.getElementById('gas-test-result-box');
  const iconSpan = document.getElementById('gas-test-icon');
  const msgDiv = document.getElementById('gas-test-msg');
  const submsgDiv = document.getElementById('gas-test-submsg');
  const testBtn = document.getElementById('btn-test-gas-connection');

  if (!url) {
    alert('Google Apps Script URL을 먼저 입력해 주세요.');
    return;
  }

  testBtn.disabled = true;
  testBtn.innerHTML = '<i data-lucide="loader-2" class="w-3.5 h-3.5 animate-spin"></i><span>확인 중...</span>';
  refreshIcons();

  state.settings.googleAppsScriptUrl = url;
  const res = await callSheetsApi('ping');

  testBtn.disabled = false;
  testBtn.innerHTML = '<i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i><span>연결 테스트</span>';

  resultBox.classList.remove('hidden');
  if (res.success && res.data?.success) {
    resultBox.className = 'p-3 rounded-xl border text-xs flex items-start gap-2.5 bg-emerald-50 border-emerald-200 text-emerald-800';
    iconSpan.innerHTML = '<i data-lucide="check-circle-2" class="w-4 h-4 text-emerald-600 shrink-0 mt-0.5"></i>';
    msgDiv.textContent = res.data.message || '연결 성공! 스프레드시트가 확인되었습니다.';
    submsgDiv.textContent = res.data.spreadsheetName ? `연결된 스프레드시트: ${res.data.spreadsheetName}` : '';
    saveSettingsToStorage();
    updateSheetsBadge();
  } else {
    resultBox.className = 'p-3 rounded-xl border text-xs flex items-start gap-2.5 bg-red-50 border-red-200 text-red-800';
    iconSpan.innerHTML = '<i data-lucide="alert-circle" class="w-4 h-4 text-red-600 shrink-0 mt-0.5"></i>';
    msgDiv.textContent = res.error || res.data?.error || '연결 실패. 배포 권한이 "모든 사용자(Anyone)"인지 확인해 주세요.';
    submsgDiv.textContent = '';
  }
  refreshIcons();
}

async function bulkSyncAllAnalyses() {
  if (!state.settings.googleAppsScriptUrl) {
    alert('Google Apps Script URL이 설정되지 않았습니다.');
    return;
  }

  const btn = document.getElementById('btn-sync-all-records');
  const statusBox = document.getElementById('bulk-sync-status-text');

  btn.disabled = true;
  statusBox.classList.remove('hidden');

  let successCount = 0;
  for (let i = 0; i < state.analyses.length; i++) {
    const item = state.analyses[i];
    statusBox.textContent = `${i + 1} / ${state.analyses.length} 건 동기화 중: ${item.title}`;
    const res = await callSheetsApi('saveAnalysis', { analysis: item });
    if (res.success) {
      item.syncedToSheets = true;
      successCount++;
    }
  }

  saveAnalysesToStorage();
  btn.disabled = false;
  statusBox.textContent = `동기화 완료: ${successCount}건이 Google Sheets에 저장되었습니다!`;
  updateSheetsBadge();
  if (state.currentTab === 'history') renderHistoryScreen();
  if (state.currentTab === 'result') renderResultScreen();
}

// ==========================================
// 15. DOM Initialization & Event Setup
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
  loadStorage();
  renderPresets();
  renderSheetsGuideSteps();
  updateBookmarksCount();
  updateSheetsBadge();

  // Navigation logo & tabs
  document.getElementById('nav-logo').addEventListener('click', () => switchScreen('input'));
  document.getElementById('tab-btn-input').addEventListener('click', () => switchScreen('input'));
  document.getElementById('tab-btn-result').addEventListener('click', () => switchScreen('result'));
  document.getElementById('tab-btn-history').addEventListener('click', () => switchScreen('history'));

  // Level selector cards
  document.getElementById('card-level-beginner').addEventListener('click', () => selectLevel('beginner'));
  document.getElementById('card-level-practitioner').addEventListener('click', () => selectLevel('practitioner'));

  // Input subtabs (File vs Text)
  document.getElementById('btn-subtab-file').addEventListener('click', () => selectSubTab('file'));
  document.getElementById('btn-subtab-text').addEventListener('click', () => selectSubTab('text'));

  // Character counter for direct textarea
  const textarea = document.getElementById('textarea-rawtext');
  textarea.addEventListener('input', () => {
    document.getElementById('text-char-count').textContent = textarea.value.length.toLocaleString() + ' 자';
  });

  // Dropzone file handling
  const dropzone = document.getElementById('dropzone');
  const fileInputHidden = document.getElementById('file-input-hidden');

  dropzone.addEventListener('click', () => fileInputHidden.click());
  document.getElementById('btn-reselect-file').addEventListener('click', () => fileInputHidden.click());
  document.getElementById('btn-clear-file').addEventListener('click', clearUploadedFile);

  fileInputHidden.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) {
      handleFileSelected(e.target.files[0]);
    }
  });

  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('border-blue-500', 'bg-blue-50/50');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('border-blue-500', 'bg-blue-50/50');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('border-blue-500', 'bg-blue-50/50');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFileSelected(e.dataTransfer.files[0]);
    }
  });

  // Main Analysis Form Submit
  document.getElementById('form-analysis').addEventListener('submit', runAnalysis);

  // Result Screen Subtabs
  document.getElementById('res-tab-btn-summary').addEventListener('click', () => selectResultSubTab('summary'));
  document.getElementById('res-tab-btn-background').addEventListener('click', () => selectResultSubTab('background'));
  document.getElementById('res-tab-btn-terminology').addEventListener('click', () => selectResultSubTab('terminology'));
  document.getElementById('res-tab-btn-qa').addEventListener('click', () => selectResultSubTab('qa'));

  // Terminology search input
  document.getElementById('input-filter-term').addEventListener('input', (e) => {
    state.termSearchKeyword = e.target.value;
    renderTerminologyList();
  });

  // Result Screen Action Buttons
  document.getElementById('btn-back-history').addEventListener('click', () => switchScreen('history'));
  document.getElementById('btn-new-analysis').addEventListener('click', () => switchScreen('input'));
  document.getElementById('btn-print').addEventListener('click', () => window.print());
  document.getElementById('btn-copy-markdown').addEventListener('click', copyReportMarkdown);

  // Save to Sheets Action button
  document.getElementById('btn-sync-sheets-action').addEventListener('click', async () => {
    if (!state.activeAnalysis) return;
    if (!state.settings.googleAppsScriptUrl) {
      document.getElementById('modal-sheets').classList.remove('hidden');
      return;
    }
    const res = await callSheetsApi('saveAnalysis', { analysis: state.activeAnalysis });
    if (res.success) {
      state.activeAnalysis.syncedToSheets = true;
      saveAnalysesToStorage();
      document.getElementById('sync-sheets-btn-text').textContent = 'Google Sheets 저장됨';
      showToast('Google Sheets에 성공적으로 저장되었습니다!', 'success');
    } else {
      showToast(res.error || 'Google Sheets 저장 실패', 'error');
    }
  });

  // History search & filters
  document.getElementById('input-history-search').addEventListener('input', (e) => {
    state.historySearchKeyword = e.target.value;
    renderHistoryScreen();
  });

  document.getElementById('btn-history-new-analysis').addEventListener('click', () => switchScreen('input'));
  document.getElementById('btn-empty-start-analysis').addEventListener('click', () => switchScreen('input'));

  const historyFilters = [
    { id: 'filter-history-all', lvl: 'all' },
    { id: 'filter-history-beginner', lvl: 'beginner' },
    { id: 'filter-history-practitioner', lvl: 'practitioner' },
  ];

  historyFilters.forEach(({ id, lvl }) => {
    document.getElementById(id).addEventListener('click', () => {
      state.historyFilterLevel = lvl;
      historyFilters.forEach((f) => {
        const btn = document.getElementById(f.id);
        if (f.lvl === lvl) {
          btn.className = 'history-filter-btn px-3 py-1.5 rounded-lg text-xs font-medium bg-slate-900 text-white transition-colors';
        } else {
          btn.className = 'history-filter-btn px-3 py-1.5 rounded-lg text-xs font-medium text-slate-600 hover:bg-slate-100 transition-colors';
        }
      });
      renderHistoryScreen();
    });
  });

  // Bookmarks Modal
  const modalBookmarks = document.getElementById('modal-bookmarks');
  document.getElementById('btn-open-bookmarks').addEventListener('click', () => {
    renderBookmarksModal();
    modalBookmarks.classList.remove('hidden');
  });
  document.getElementById('btn-close-bookmarks').addEventListener('click', () => {
    modalBookmarks.classList.add('hidden');
  });
  document.getElementById('input-search-bookmarks').addEventListener('input', (e) => {
    state.bookmarksSearchKeyword = e.target.value;
    renderBookmarksModal();
  });

  // Google Sheets Settings Modal
  const modalSheets = document.getElementById('modal-sheets');
  document.getElementById('btn-open-sheets').addEventListener('click', () => {
    document.getElementById('input-gas-url').value = state.settings.googleAppsScriptUrl || '';
    document.getElementById('checkbox-auto-sync').checked = state.settings.autoSyncToSheets;
    if (state.analyses.length > 0 && state.settings.googleAppsScriptUrl) {
      document.getElementById('box-bulk-sync').classList.remove('hidden');
    }
    modalSheets.classList.remove('hidden');
    refreshIcons();
  });

  document.getElementById('btn-close-sheets').addEventListener('click', () => modalSheets.classList.add('hidden'));
  document.getElementById('btn-cancel-sheets-modal').addEventListener('click', () => modalSheets.classList.add('hidden'));

  document.getElementById('btn-copy-gas-code').addEventListener('click', () => {
    navigator.clipboard.writeText(GOOGLE_APPS_SCRIPT_CODE);
    const span = document.getElementById('gas-copy-btn-text');
    span.textContent = '코드 복사됨!';
    setTimeout(() => {
      span.textContent = 'Apps Script 코드 복사';
    }, 2500);
    showToast('Apps Script 코드가 클립보드에 복사되었습니다.', 'info');
  });

  document.getElementById('btn-test-gas-connection').addEventListener('click', testSheetsConnection);
  document.getElementById('btn-sync-all-records').addEventListener('click', bulkSyncAllAnalyses);

  document.getElementById('btn-save-sheets-settings').addEventListener('click', () => {
    state.settings.googleAppsScriptUrl = document.getElementById('input-gas-url').value.trim();
    state.settings.autoSyncToSheets = document.getElementById('checkbox-auto-sync').checked;
    saveSettingsToStorage();
    updateSheetsBadge();
    modalSheets.classList.add('hidden');
    showToast('Google Sheets 설정이 저장되었습니다.', 'success');
  });

  // Initial screen setup
  switchScreen('input');
});
