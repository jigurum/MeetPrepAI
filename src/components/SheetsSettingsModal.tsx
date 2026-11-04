import React, { useState } from 'react';
import {
  X,
  Database,
  Check,
  Copy,
  ExternalLink,
  AlertCircle,
  CheckCircle2,
  HelpCircle,
  Loader2,
  RefreshCw,
  FileSpreadsheet,
  Table,
  Upload,
} from 'lucide-react';
import { AppSettings, AnalysisRecord } from '../types';
import {
  GOOGLE_APPS_SCRIPT_CODE,
  APPS_SCRIPT_GUIDE_STEPS,
  testAppsScriptConnection,
  saveAnalysisToGoogleSheets,
} from '../services/appsScriptTemplate';

interface SheetsSettingsModalProps {
  isOpen: boolean;
  onClose: () => void;
  settings: AppSettings;
  onSaveSettings: (settings: AppSettings) => void;
  analyses: AnalysisRecord[];
  onAnalysesSynced: (updatedAnalyses: AnalysisRecord[]) => void;
}

export const SheetsSettingsModal: React.FC<SheetsSettingsModalProps> = ({
  isOpen,
  onClose,
  settings,
  onSaveSettings,
  analyses,
  onAnalysesSynced,
}) => {
  const [url, setUrl] = useState(settings.googleAppsScriptUrl || '');
  const [autoSync, setAutoSync] = useState(settings.autoSyncToSheets);
  const [copiedCode, setCopiedCode] = useState(false);
  const [testing, setTesting] = useState(false);
  const [testResult, setTestResult] = useState<{
    success: boolean;
    message: string;
    spreadsheetName?: string;
    spreadsheetUrl?: string;
  } | null>(null);
  const [syncingAll, setSyncingAll] = useState(false);
  const [syncAllStatus, setSyncAllStatus] = useState<string | null>(null);

  if (!isOpen) return null;

  const handleCopyCode = () => {
    navigator.clipboard.writeText(GOOGLE_APPS_SCRIPT_CODE);
    setCopiedCode(true);
    setTimeout(() => setCopiedCode(false), 2500);
  };

  const handleTestConnection = async () => {
    if (!url.trim()) {
      setTestResult({
        success: false,
        message: 'Google Apps Script 웹 앱 URL을 먼저 입력해 주세요.',
      });
      return;
    }

    setTesting(true);
    setTestResult(null);

    const result = await testAppsScriptConnection(url.trim());
    setTesting(false);

    if (result.success && result.data?.success) {
      setTestResult({
        success: true,
        message: result.data.message || '연결 성공! 스프레드시트가 확인되었습니다.',
        spreadsheetName: result.data.spreadsheetName,
        spreadsheetUrl: result.data.spreadsheetUrl,
      });
      // Save valid URL automatically
      onSaveSettings({
        ...settings,
        googleAppsScriptUrl: url.trim(),
        autoSyncToSheets: autoSync,
      });
    } else {
      setTestResult({
        success: false,
        message:
          result.error ||
          result.data?.error ||
          '연결에 실패했습니다. 배포 권한이 "모든 사용자(Anyone)"인지 확인해 주세요.',
      });
    }
  };

  const handleSave = () => {
    onSaveSettings({
      ...settings,
      googleAppsScriptUrl: url.trim(),
      autoSyncToSheets: autoSync,
    });
    onClose();
  };

  const handleSyncAllAnalyses = async () => {
    if (!url.trim()) {
      alert('Google Apps Script URL이 설정되지 않았습니다.');
      return;
    }

    setSyncingAll(true);
    setSyncAllStatus(`0 / ${analyses.length} 건 동기화 중...`);

    const updatedList = [...analyses];
    let successCount = 0;

    for (let i = 0; i < analyses.length; i++) {
      const item = analyses[i];
      setSyncAllStatus(`${i + 1} / ${analyses.length} 건 저장 중: ${item.title}`);
      const res = await saveAnalysisToGoogleSheets(url.trim(), item);
      if (res.success) {
        updatedList[i] = { ...item, syncedToSheets: true };
        successCount++;
      }
    }

    onAnalysesSynced(updatedList);
    setSyncingAll(false);
    setSyncAllStatus(`동기화 완료: ${successCount}건이 Google Sheets에 저장되었습니다!`);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm animate-in fade-in duration-200">
      <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl w-full max-w-3xl max-h-[90vh] flex flex-col overflow-hidden">
        {/* Header */}
        <div className="p-5 border-b border-slate-200 flex items-center justify-between bg-slate-50/50">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-xl bg-emerald-100 text-emerald-700 flex items-center justify-center">
              <Database className="w-5 h-5" />
            </div>
            <div>
              <h3 className="font-bold text-base text-slate-900">
                Google Sheets & Apps Script 연동 가이드
              </h3>
              <p className="text-xs text-slate-500">
                2개 시트 구조(분석 기록 + 전문용어)를 클라우드 스프레드시트에 영구 저장합니다.
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-200/60 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-6 overflow-y-auto space-y-6 flex-1 text-slate-700">
          {/* Architecture Overview */}
          <div className="bg-blue-50/50 p-4 rounded-xl border border-blue-200/80">
            <h4 className="text-xs font-bold text-blue-900 uppercase tracking-wider mb-2 flex items-center gap-1.5">
              <Table className="w-4 h-4 text-blue-700" />
              자동 생성되는 Google Sheets 2개 시트 데이터 구조
            </h4>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
              <div className="bg-white p-3 rounded-lg border border-blue-100 shadow-xs">
                <span className="font-bold text-slate-900 block mb-1">
                  1. [분석 기록] 시트
                </span>
                <p className="text-slate-500 text-[11px] leading-tight">
                  분석ID, 분석일시, 자료명, 입력방식, 설명수준, 핵심요약, 배경지식
                </p>
              </div>
              <div className="bg-white p-3 rounded-lg border border-blue-100 shadow-xs">
                <span className="font-bold text-slate-900 block mb-1">
                  2. [전문용어] 시트
                </span>
                <p className="text-slate-500 text-[11px] leading-tight">
                  용어ID, 분석ID, 전문용어, 맞춤형 설명, 실무 활용 사례, 북마크
                </p>
              </div>
            </div>
          </div>

          {/* 5-Step Visual Guide */}
          <div className="space-y-3">
            <h4 className="text-xs font-bold text-slate-900 uppercase tracking-wider">
              간단한 5단계 연동 절차
            </h4>
            <div className="space-y-2">
              {APPS_SCRIPT_GUIDE_STEPS.map((s) => (
                <div
                  key={s.step}
                  className="p-3 rounded-xl border border-slate-200 bg-white flex items-start gap-3"
                >
                  <span className="w-6 h-6 rounded-full bg-slate-900 text-white text-xs font-bold flex items-center justify-center shrink-0 mt-0.5">
                    {s.step}
                  </span>
                  <div>
                    <h5 className="font-bold text-xs text-slate-900">{s.title}</h5>
                    <p className="text-xs text-slate-500 mt-0.5 leading-relaxed">
                      {s.description}
                    </p>
                  </div>
                </div>
              ))}
            </div>

            {/* Cost & Quota Reassurance Banner */}
            <div className="p-3.5 rounded-xl bg-emerald-50/80 border border-emerald-200 text-xs text-emerald-950 flex items-start gap-2.5">
              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
              <div>
                <strong className="font-bold text-emerald-900 block mb-0.5">
                  비용 발생 여부 안내: 100% 완전 무료 (0원)
                </strong>
                <p className="text-[11px] text-emerald-800 leading-relaxed">
                  Google Apps Script는 신용카드 등록이 필요 없는 <strong>완전 무료 서비스</strong>입니다.
                  배포 시 [모든 사용자(Anyone)]를 선택하더라도 <strong>비용은 1원도 청구되지 않습니다.</strong>
                  Google의 일일 무료 쿼터(하루 수만 회 실행 가능)로 동작하며, 한도를 넘더라도 일시 대기 상태가 될 뿐 요금이 부과되지 않습니다.
                </p>
              </div>
            </div>
          </div>

          {/* Copy Script Code Section */}
          <div className="p-4 rounded-xl border border-slate-200 bg-slate-50 space-y-3">
            <div className="flex items-center justify-between">
              <div>
                <span className="text-xs font-bold text-slate-900 block">
                  Google Apps Script 배포 코드
                </span>
                <span className="text-[11px] text-slate-500">
                  Apps Script 편집기(Code.gs)에 복사하여 붙여넣으세요.
                </span>
              </div>
              <button
                type="button"
                onClick={handleCopyCode}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold shadow-sm transition-colors"
              >
                {copiedCode ? (
                  <>
                    <Check className="w-3.5 h-3.5" />
                    <span>코드 복사됨!</span>
                  </>
                ) : (
                  <>
                    <Copy className="w-3.5 h-3.5" />
                    <span>Apps Script 코드 복사</span>
                  </>
                )}
              </button>
            </div>
            <pre className="text-[10px] font-mono bg-slate-900 text-slate-200 p-3 rounded-lg max-h-36 overflow-y-auto leading-relaxed border border-slate-800">
              {GOOGLE_APPS_SCRIPT_CODE.slice(0, 500)}
              {'\n... [전체 코드 복사 버튼을 누르면 모든 함수가 클립보드에 복사됩니다] ...'}
            </pre>
          </div>

          {/* Web App URL Input & Connection Test */}
          <div className="p-4 rounded-xl border border-slate-200 bg-white space-y-4">
            <div>
              <label htmlFor="gas-url" className="block text-xs font-bold text-slate-900 mb-1">
                배포된 웹 앱 URL (Web App URL) <span className="text-red-500">*</span>
              </label>
              <div className="flex gap-2">
                <input
                  id="gas-url"
                  type="url"
                  value={url}
                  onChange={(e) => setUrl(e.target.value)}
                  placeholder="https://script.google.com/macros/s/AKfycb.../exec"
                  className="flex-1 px-3.5 py-2 text-xs rounded-xl border border-slate-300 font-mono focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600"
                />
                <button
                  type="button"
                  onClick={handleTestConnection}
                  disabled={testing}
                  className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-900 text-white text-xs font-bold flex items-center gap-1.5 transition-colors disabled:opacity-50"
                >
                  {testing ? (
                    <>
                      <Loader2 className="w-3.5 h-3.5 animate-spin" />
                      <span>확인 중...</span>
                    </>
                  ) : (
                    <>
                      <RefreshCw className="w-3.5 h-3.5" />
                      <span>연결 테스트</span>
                    </>
                  )}
                </button>
              </div>
              <p className="text-[11px] text-slate-400 mt-1">
                Apps Script 배포 완료 시 제공되는 URL(마지막이 <code className="font-mono">/exec</code>로 끝남)을 입력합니다.
              </p>
            </div>

            {/* Test Result Message */}
            {testResult && (
              <div
                className={`p-3 rounded-xl border text-xs flex items-start gap-2.5 ${
                  testResult.success
                    ? 'bg-emerald-50 border-emerald-200 text-emerald-800'
                    : 'bg-red-50 border-red-200 text-red-800'
                }`}
              >
                {testResult.success ? (
                  <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                ) : (
                  <AlertCircle className="w-4 h-4 text-red-600 shrink-0 mt-0.5" />
                )}
                <div>
                  <div className="font-bold">{testResult.message}</div>
                  {testResult.spreadsheetName && (
                    <div className="mt-1 text-[11px]">
                      연결된 스프레드시트: <strong>{testResult.spreadsheetName}</strong>
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* Auto sync checkbox */}
            <div className="flex items-center gap-2 pt-1">
              <input
                id="auto-sync"
                type="checkbox"
                checked={autoSync}
                onChange={(e) => setAutoSync(e.target.checked)}
                className="w-4 h-4 rounded text-blue-600 focus:ring-blue-500 border-slate-300"
              />
              <label htmlFor="auto-sync" className="text-xs font-medium text-slate-700 cursor-pointer">
                새 회의 분석 완료 시 자동으로 Google Sheets에 저장하기
              </label>
            </div>

            {/* Bulk sync existing analyses */}
            {analyses.length > 0 && url && (
              <div className="pt-3 border-t border-slate-100 flex items-center justify-between">
                <div>
                  <span className="text-xs font-semibold text-slate-800 block">
                    기존 분석 기록 일괄 전송
                  </span>
                  <span className="text-[11px] text-slate-400">
                    현재 보관 중인 {analyses.length}건의 분석 기록을 시트로 전송합니다.
                  </span>
                </div>
                <button
                  type="button"
                  onClick={handleSyncAllAnalyses}
                  disabled={syncingAll}
                  className="px-3 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs font-semibold flex items-center gap-1.5 transition-colors"
                >
                  {syncingAll ? (
                    <Loader2 className="w-3.5 h-3.5 animate-spin text-blue-600" />
                  ) : (
                    <Upload className="w-3.5 h-3.5" />
                  )}
                  <span>전체 동기화</span>
                </button>
              </div>
            )}
            {syncAllStatus && (
              <div className="text-xs font-medium text-blue-700 bg-blue-50 p-2.5 rounded-lg border border-blue-200">
                {syncAllStatus}
              </div>
            )}
          </div>
        </div>

        {/* Footer */}
        <div className="p-4 border-t border-slate-200 flex items-center justify-between bg-slate-50/50">
          <span className="text-[11px] text-slate-400">
            * 스프레드시트 URL이 없어도 모든 분석 및 북마크는 로컬 브라우저에 즉시 안전 보관됩니다.
          </span>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 rounded-xl text-xs font-medium text-slate-600 hover:bg-slate-200/60 transition-colors"
            >
              닫기
            </button>
            <button
              type="button"
              onClick={handleSave}
              className="px-5 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold shadow-sm transition-colors"
            >
              설정 저장
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
