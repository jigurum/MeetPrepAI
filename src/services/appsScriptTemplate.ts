/**
 * Google Apps Script Deployment Code and Sync Services
 * Connects MeetPrep AI with Google Sheets using the 2-Sheet Architecture
 */

import { AnalysisRecord, TermItem } from '../types';

export const GOOGLE_APPS_SCRIPT_CODE = `/**
 * ==============================================================================
 * MeetPrep AI - Google Sheets 백엔드 연동 Apps Script
 * ==============================================================================
 * 
 * [시트 1: 분석 기록]
 * - 열 구조: 분석ID | 분석일시 | 자료명 | 입력방식 | 설명수준 | 핵심요약 | 배경지식
 * 
 * [시트 2: 전문용어]
 * - 열 구조: 용어ID | 분석ID | 전문용어 | 맞춤형 설명 | 실무 활용 사례 | 북마크
 * 
 * 두 시트는 '분석ID'로 1:N 자동 연결됩니다.
 * ==============================================================================
 */

// 시트 이름 정의
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

    // 2개 시트 생성 및 서식 초기화 확인
    var sheetAnalysis = ensureSheet(ss, SHEET_ANALYSIS_NAME, [
      "분석ID", "분석일시", "자료명", "입력방식", "설명수준", "핵심요약", "배경지식"
    ], "#1E3A8A"); // Navy header

    var sheetTerms = ensureSheet(ss, SHEET_TERMS_NAME, [
      "용어ID", "분석ID", "전문용어", "맞춤형 설명", "실무 활용 사례", "북마크"
    ], "#2563EB"); // Blue header

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
          termsCount: Math.max(0, sheetTerms.getLastRow() - 1),
          timestamp: new Date().toISOString()
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
        responseData = { success: false, error: "알 수 없는 작업 요청입니다: " + action };
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

/**
 * 시트 존재 확인 및 헤더 생성
 */
function ensureSheet(ss, name, headers, headerColor) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    // 기본 첫 시트가 빈 시트인 경우 정리 가능
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

/**
 * 분석 결과 저장 (시트 1: 분석 기록, 시트 2: 전문용어)
 */
function saveAnalysisRecord(sheetAnalysis, sheetTerms, analysis) {
  if (!analysis || !analysis.analysisId) {
    return { success: false, error: "유효한 분석 데이터가 전달되지 않았습니다." };
  }

  // 중복 저장 방지: 이미 존재하는 분석ID인지 확인
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
               return "• " + r.entityA + " ➔ " + r.relationship + " ➔ " + r.entityB;
             }).join("\\n");
  }

  // 시트 1: 분석 기록에 추가 (이미 있으면 업데이트)
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

  // 시트 2: 전문용어에 추가
  // 기존 해당 분석ID의 용어 제거 후 재삽입
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

/**
 * 북마크 체크박스 토글
 */
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

/**
 * 분석 기록 및 관련 전문용어 삭제
 */
function deleteAnalysisRecord(sheetAnalysis, sheetTerms, analysisId) {
  // 1. 분석 기록 삭제
  var lastRow = sheetAnalysis.getLastRow();
  if (lastRow > 1) {
    var ids = sheetAnalysis.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    var idx = ids.indexOf(analysisId);
    if (idx >= 0) {
      sheetAnalysis.deleteRow(idx + 2);
    }
  }

  // 2. 전문용어 시트에서 삭제
  var termsLastRow = sheetTerms.getLastRow();
  if (termsLastRow > 1) {
    var termAIds = sheetTerms.getRange(2, 2, termsLastRow - 1, 1).getValues().flat();
    for (var i = termAIds.length - 1; i >= 0; i--) {
      if (termAIds[i] === analysisId) {
        sheetTerms.deleteRow(i + 2);
      }
    }
  }

  return { success: true, message: "분석 기록이 시트에서 삭제되었습니다." };
}

/**
 * 저장된 모든 분석 및 용어 조회
 */
function getAllAnalyses(sheetAnalysis, sheetTerms) {
  var lastRow = sheetAnalysis.getLastRow();
  if (lastRow <= 1) {
    return { success: true, analyses: [] };
  }

  var analysisRows = sheetAnalysis.getRange(2, 1, lastRow - 1, 7).getValues();
  var termsRows = [];
  var termsLastRow = sheetTerms.getLastRow();
  if (termsLastRow > 1) {
    termsRows = sheetTerms.getRange(2, 1, termsLastRow - 1, 6).getValues();
  }

  // 그룹화
  var termsMap = {};
  for (var i = 0; i < termsRows.length; i++) {
    var r = termsRows[i];
    var aId = r[1];
    if (!termsMap[aId]) termsMap[aId] = [];
    termsMap[aId].push({
      termId: r[0],
      analysisId: r[1],
      term: r[2],
      easyExplanation: r[3],
      practicalUseCases: r[4],
      bookmarked: r[5] === true || r[5] === "TRUE"
    });
  }

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
      bgRaw: a[6],
      terminology: termsMap[a[0]] || []
    });
  }

  return { success: true, analyses: list };
}
`;

/**
 * Proxy call helper to communicate with Google Apps Script without CORS issues
 */
export async function callAppsScriptProxy(
  scriptUrl: string,
  action: string,
  payload: Record<string, any> = {}
): Promise<{ success: boolean; data?: any; error?: string }> {
  try {
    const res = await fetch('/api/sheets-proxy', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        scriptUrl,
        action,
        payload,
      }),
    });

    const result = await res.json();
    if (!res.ok || result.error) {
      return { success: false, error: result.error || 'Google Apps Script 통신 실패' };
    }

    return { success: true, data: result.data };
  } catch (error: any) {
    return { success: false, error: error.message || '네트워크 오류가 발생했습니다.' };
  }
}

/**
 * Test Apps Script connection
 */
export async function testAppsScriptConnection(scriptUrl: string) {
  return await callAppsScriptProxy(scriptUrl, 'ping');
}

/**
 * Save Analysis to Google Sheets
 */
export async function saveAnalysisToGoogleSheets(scriptUrl: string, analysis: AnalysisRecord) {
  return await callAppsScriptProxy(scriptUrl, 'saveAnalysis', { analysis });
}

/**
 * Toggle bookmark on Google Sheets
 */
export async function toggleBookmarkInGoogleSheets(
  scriptUrl: string,
  termId: string,
  bookmarked: boolean
) {
  return await callAppsScriptProxy(scriptUrl, 'toggleBookmark', { termId, bookmarked });
}

/**
 * Delete analysis from Google Sheets
 */
export async function deleteAnalysisFromGoogleSheets(scriptUrl: string, analysisId: string) {
  return await callAppsScriptProxy(scriptUrl, 'deleteAnalysis', { analysisId });
}

/**
 * Step-by-step setup guide texts
 */
export const APPS_SCRIPT_GUIDE_STEPS = [
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
