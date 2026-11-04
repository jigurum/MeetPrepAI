/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { Header } from './components/Header';
import { InputScreen } from './components/InputScreen';
import { ResultScreen } from './components/ResultScreen';
import { HistoryScreen } from './components/HistoryScreen';
import { BookmarksModal } from './components/BookmarksModal';
import { SheetsSettingsModal } from './components/SheetsSettingsModal';
import { AnalysisRecord, AppSettings, ExplanationLevel } from './types';
import {
  getStoredAnalyses,
  saveAnalysis,
  deleteAnalysis,
  toggleTermBookmark,
  getAppSettings,
  saveAppSettings,
  getAllBookmarkedTerms,
} from './services/storage';
import { saveAnalysisToGoogleSheets } from './services/appsScriptTemplate';
import { CheckCircle2, AlertCircle, X } from 'lucide-react';

export default function App() {
  const [currentTab, setCurrentTab] = useState<'input' | 'history' | 'result'>('input');
  const [analyses, setAnalyses] = useState<AnalysisRecord[]>([]);
  const [activeAnalysis, setActiveAnalysis] = useState<AnalysisRecord | null>(null);
  const [settings, setSettings] = useState<AppSettings>(getAppSettings());
  const [isLoading, setIsLoading] = useState(false);
  const [isSyncingSheets, setIsSyncingSheets] = useState(false);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isBookmarksOpen, setIsBookmarksOpen] = useState(false);
  const [toast, setToast] = useState<{ type: 'success' | 'error' | 'info'; message: string } | null>(null);

  // Load stored analyses on mount
  useEffect(() => {
    const list = getStoredAnalyses();
    setAnalyses(list);
    if (list.length > 0 && !activeAnalysis) {
      setActiveAnalysis(list[0]);
    }
  }, []);

  const showToast = (message: string, type: 'success' | 'error' | 'info' = 'success') => {
    setToast({ type, message });
    setTimeout(() => {
      setToast(null);
    }, 4000);
  };

  // Start AI analysis
  const handleStartAnalysis = async (params: {
    title: string;
    level: ExplanationLevel;
    inputType: 'text' | 'file';
    rawText: string;
    fileData?: { mimeType: string; base64: string; fileName: string };
  }) => {
    if (isLoading) return; // Prevent duplicate clicks (PRD requirement)
    setIsLoading(true);

    try {
      const response = await fetch('/api/analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(params),
      });

      const data = await response.json();
      if (!response.ok || data.error) {
        throw new Error(data.error || '회의 자료 분석에 실패했습니다.');
      }

      const newRecord: AnalysisRecord = data;

      // Save to local storage and sync to Google Sheets if configured
      const { record: savedRecord, sheetSyncResult } = await saveAnalysis(
        newRecord,
        settings.autoSyncToSheets
      );

      // Update state
      setAnalyses((prev) => [savedRecord, ...prev.filter((a) => a.analysisId !== savedRecord.analysisId)]);
      setActiveAnalysis(savedRecord);
      setCurrentTab('result');

      if (sheetSyncResult?.success) {
        showToast('회의 자료 분석 완료 및 Google Sheets 저장 성공!', 'success');
      } else if (sheetSyncResult?.error) {
        showToast(`분석 완료 (시트 저장 실패: ${sheetSyncResult.error})`, 'info');
      } else {
        showToast('회의 자료 맞춤 분석이 완료되었습니다!', 'success');
      }
    } catch (error: any) {
      console.error('Analysis error:', error);
      showToast(error.message || '분석 중 오류가 발생했습니다. 다시 시도해 주세요.', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  // Toggle bookmark
  const handleToggleBookmark = async (termId: string, newState: boolean) => {
    if (!activeAnalysis) return;
    const updated = await toggleTermBookmark(activeAnalysis.analysisId, termId, newState);
    if (updated) {
      setActiveAnalysis({ ...updated });
      setAnalyses((prev) =>
        prev.map((a) => (a.analysisId === updated.analysisId ? { ...updated } : a))
      );
    }
  };

  // Toggle bookmark from modal
  const handleToggleBookmarkFromModal = async (
    analysisId: string,
    termId: string,
    newState: boolean
  ) => {
    const updated = await toggleTermBookmark(analysisId, termId, newState);
    if (updated) {
      if (activeAnalysis?.analysisId === analysisId) {
        setActiveAnalysis({ ...updated });
      }
      setAnalyses((prev) =>
        prev.map((a) => (a.analysisId === updated.analysisId ? { ...updated } : a))
      );
    }
  };

  // Sync active analysis to Google Sheets
  const handleSyncActiveAnalysisToSheets = async () => {
    if (!activeAnalysis) return;
    if (!settings.googleAppsScriptUrl) {
      setIsSettingsOpen(true);
      return;
    }

    setIsSyncingSheets(true);
    try {
      const res = await saveAnalysisToGoogleSheets(settings.googleAppsScriptUrl, activeAnalysis);
      if (res.success) {
        const updated = { ...activeAnalysis, syncedToSheets: true };
        setActiveAnalysis(updated);
        await saveAnalysis(updated, false);
        setAnalyses((prev) =>
          prev.map((a) => (a.analysisId === updated.analysisId ? updated : a))
        );
        showToast('Google Sheets에 성공적으로 저장되었습니다!', 'success');
      } else {
        showToast(res.error || 'Google Sheets 저장에 실패했습니다.', 'error');
      }
    } catch (e: any) {
      showToast(e.message || 'Google Sheets 통신 오류가 발생했습니다.', 'error');
    } finally {
      setIsSyncingSheets(false);
    }
  };

  // Delete analysis
  const handleDeleteAnalysis = async (analysisId: string) => {
    await deleteAnalysis(analysisId);
    const updated = analyses.filter((a) => a.analysisId !== analysisId);
    setAnalyses(updated);
    if (activeAnalysis?.analysisId === analysisId) {
      setActiveAnalysis(updated.length > 0 ? updated[0] : null);
    }
    showToast('분석 기록이 삭제되었습니다.', 'info');
  };

  // Save app settings
  const handleSaveSettings = (newSettings: AppSettings) => {
    setSettings(newSettings);
    saveAppSettings(newSettings);
    showToast('설정이 저장되었습니다.', 'success');
  };

  const bookmarksList = getAllBookmarkedTerms();

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col font-sans selection:bg-blue-500/20">
      {/* Header */}
      <Header
        currentTab={currentTab}
        onSelectTab={setCurrentTab}
        onOpenSettings={() => setIsSettingsOpen(true)}
        onOpenBookmarks={() => setIsBookmarksOpen(true)}
        bookmarkedCount={bookmarksList.length}
        settings={settings}
        hasActiveResult={Boolean(activeAnalysis)}
      />

      {/* Main Content Area */}
      <main className="flex-1 pb-16">
        {currentTab === 'input' && (
          <InputScreen
            onStartAnalysis={handleStartAnalysis}
            isLoading={isLoading}
            defaultLevel={settings.defaultLevel}
          />
        )}

        {currentTab === 'result' && activeAnalysis && (
          <ResultScreen
            analysis={activeAnalysis}
            onToggleBookmark={handleToggleBookmark}
            onSyncToSheets={handleSyncActiveAnalysisToSheets}
            isSyncingSheets={isSyncingSheets}
            sheetsUrlConfigured={Boolean(settings.googleAppsScriptUrl)}
            onBackToHistory={() => setCurrentTab('history')}
            onNewAnalysis={() => setCurrentTab('input')}
            onOpenSheetsSettings={() => setIsSettingsOpen(true)}
          />
        )}

        {currentTab === 'history' && (
          <HistoryScreen
            analyses={analyses}
            onSelectAnalysis={(item) => {
              setActiveAnalysis(item);
              setCurrentTab('result');
            }}
            onDeleteAnalysis={handleDeleteAnalysis}
            onStartNewAnalysis={() => setCurrentTab('input')}
          />
        )}
      </main>

      {/* Toast Notification */}
      {toast && (
        <div className="fixed bottom-5 right-5 z-50 animate-in slide-in-from-bottom-3 duration-200">
          <div
            className={`flex items-center gap-2.5 px-4 py-3 rounded-xl shadow-lg border text-xs font-semibold ${
              toast.type === 'success'
                ? 'bg-slate-900 text-white border-slate-800'
                : toast.type === 'error'
                ? 'bg-red-600 text-white border-red-700'
                : 'bg-slate-800 text-white border-slate-700'
            }`}
          >
            {toast.type === 'success' ? (
              <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
            ) : (
              <AlertCircle className="w-4 h-4 text-amber-300 shrink-0" />
            )}
            <span>{toast.message}</span>
            <button
              type="button"
              onClick={() => setToast(null)}
              className="ml-2 text-slate-400 hover:text-white"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>
      )}

      {/* Bookmarks Modal (Personal Glossary) */}
      <BookmarksModal
        isOpen={isBookmarksOpen}
        onClose={() => setIsBookmarksOpen(false)}
        bookmarks={bookmarksList}
        onToggleBookmark={handleToggleBookmarkFromModal}
        onSelectAnalysis={(analysisId) => {
          const item = analyses.find((a) => a.analysisId === analysisId);
          if (item) {
            setActiveAnalysis(item);
            setCurrentTab('result');
          }
        }}
      />

      {/* Google Sheets & Apps Script Settings Modal */}
      <SheetsSettingsModal
        isOpen={isSettingsOpen}
        onClose={() => setIsSettingsOpen(false)}
        settings={settings}
        onSaveSettings={handleSaveSettings}
        analyses={analyses}
        onAnalysesSynced={(updated) => {
          setAnalyses(updated);
          if (activeAnalysis) {
            const syncedActive = updated.find((a) => a.analysisId === activeAnalysis.analysisId);
            if (syncedActive) setActiveAnalysis(syncedActive);
          }
        }}
      />
    </div>
  );
}
