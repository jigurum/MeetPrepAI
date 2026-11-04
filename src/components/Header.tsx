import React from 'react';
import { Sparkles, FileText, History, BookmarkCheck, Settings, Database, CheckCircle2, AlertCircle } from 'lucide-react';
import { AppSettings } from '../types';

interface HeaderProps {
  currentTab: 'input' | 'history' | 'result';
  onSelectTab: (tab: 'input' | 'history' | 'result') => void;
  onOpenSettings: () => void;
  onOpenBookmarks: () => void;
  bookmarkedCount: number;
  settings: AppSettings;
  hasActiveResult: boolean;
}

export const Header: React.FC<HeaderProps> = ({
  currentTab,
  onSelectTab,
  onOpenSettings,
  onOpenBookmarks,
  bookmarkedCount,
  settings,
  hasActiveResult,
}) => {
  const isSheetsConnected = Boolean(settings.googleAppsScriptUrl);

  return (
    <header className="sticky top-0 z-30 bg-white/95 backdrop-blur-md border-b border-slate-200">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex items-center justify-between h-16">
          {/* Logo & Brand */}
          <div className="flex items-center gap-3 cursor-pointer" onClick={() => onSelectTab('input')}>
            <div className="w-9 h-9 rounded-xl bg-gradient-to-tr from-blue-700 via-blue-600 to-indigo-600 flex items-center justify-center text-white shadow-sm shadow-blue-500/20">
              <Sparkles className="w-5 h-5 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="font-bold text-lg tracking-tight text-slate-900">MeetPrep AI</span>
                <span className="text-[11px] font-semibold uppercase tracking-wider px-2 py-0.5 rounded-full bg-blue-50 text-blue-700 border border-blue-200">
                  회의 준비 도우미
                </span>
              </div>
              <p className="text-xs text-slate-500 hidden sm:block">
                전문용어 & 배경지식 맞춤 분석기
              </p>
            </div>
          </div>

          {/* Navigation Tabs */}
          <nav className="flex items-center gap-1 sm:gap-2">
            <button
              type="button"
              onClick={() => onSelectTab('input')}
              className={`flex items-center gap-2 px-3.5 py-2 text-sm font-medium rounded-lg transition-colors ${
                currentTab === 'input'
                  ? 'bg-slate-100 text-slate-900 font-semibold'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-50'
              }`}
            >
              <FileText className="w-4 h-4 text-blue-600" />
              <span>자료 분석</span>
            </button>

            {hasActiveResult && (
              <button
                type="button"
                onClick={() => onSelectTab('result')}
                className={`flex items-center gap-2 px-3.5 py-2 text-sm font-medium rounded-lg transition-colors ${
                  currentTab === 'result'
                    ? 'bg-blue-50 text-blue-800 font-semibold border border-blue-200/60'
                    : 'text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                }`}
              >
                <Sparkles className="w-4 h-4 text-blue-600" />
                <span>분석 결과</span>
              </button>
            )}

            <button
              type="button"
              onClick={() => onSelectTab('history')}
              className={`flex items-center gap-2 px-3.5 py-2 text-sm font-medium rounded-lg transition-colors ${
                currentTab === 'history'
                  ? 'bg-slate-100 text-slate-900 font-semibold'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-50'
              }`}
            >
              <History className="w-4 h-4 text-slate-500" />
              <span>분석 기록</span>
            </button>
          </nav>

          {/* Right Tools: Bookmarks & Sheets Integration */}
          <div className="flex items-center gap-2">
            {/* Bookmarks Glossary button */}
            <button
              type="button"
              onClick={onOpenBookmarks}
              className="relative p-2 text-slate-600 hover:text-slate-900 hover:bg-slate-100 rounded-lg transition-colors"
              title="북마크한 전문용어 모아보기"
            >
              <BookmarkCheck className="w-5 h-5 text-amber-500" />
              {bookmarkedCount > 0 && (
                <span className="absolute -top-1 -right-1 px-1.5 py-0.2 bg-amber-500 text-white text-[10px] font-bold rounded-full min-w-4 text-center">
                  {bookmarkedCount}
                </span>
              )}
            </button>

            {/* Google Sheets Status & Settings */}
            <button
              type="button"
              onClick={onOpenSettings}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium border transition-colors ${
                isSheetsConnected
                  ? 'bg-emerald-50 text-emerald-700 border-emerald-200 hover:bg-emerald-100'
                  : 'bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100'
              }`}
            >
              <Database className="w-3.5 h-3.5" />
              <span className="hidden md:inline">Google Sheets</span>
              {isSheetsConnected ? (
                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
              ) : (
                <AlertCircle className="w-3.5 h-3.5 text-slate-400" />
              )}
              <Settings className="w-3.5 h-3.5 ml-0.5 opacity-70" />
            </button>
          </div>
        </div>
      </div>
    </header>
  );
};
