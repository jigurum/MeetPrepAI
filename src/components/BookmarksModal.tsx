import React, { useState } from 'react';
import {
  X,
  BookmarkCheck,
  Bookmark,
  Search,
  BookOpen,
  Sparkles,
  ArrowRight,
  ExternalLink,
} from 'lucide-react';
import { TermItem } from '../types';

interface BookmarksModalProps {
  isOpen: boolean;
  onClose: () => void;
  bookmarks: { term: TermItem; analysisTitle: string; analysisId: string }[];
  onToggleBookmark: (analysisId: string, termId: string, newState: boolean) => void;
  onSelectAnalysis: (analysisId: string) => void;
}

export const BookmarksModal: React.FC<BookmarksModalProps> = ({
  isOpen,
  onClose,
  bookmarks,
  onToggleBookmark,
  onSelectAnalysis,
}) => {
  const [search, setSearch] = useState('');

  if (!isOpen) return null;

  const filtered = bookmarks.filter(
    (b) =>
      b.term.term.toLowerCase().includes(search.toLowerCase()) ||
      b.term.definition.toLowerCase().includes(search.toLowerCase()) ||
      b.term.easyExplanation.toLowerCase().includes(search.toLowerCase()) ||
      b.analysisTitle.toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm animate-in fade-in duration-200">
      <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl w-full max-w-3xl max-h-[85vh] flex flex-col overflow-hidden">
        {/* Header */}
        <div className="p-5 border-b border-slate-200 flex items-center justify-between bg-slate-50/50">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-amber-100 text-amber-600 flex items-center justify-center">
              <BookmarkCheck className="w-4 h-4 fill-amber-500" />
            </div>
            <div>
              <h3 className="font-bold text-base text-slate-900">개인 전문용어 보관함</h3>
              <p className="text-xs text-slate-500">
                각 회의 자료에서 북마크한 나만의 핵심 전문용어 모음 ({bookmarks.length}개)
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

        {/* Search */}
        <div className="p-4 border-b border-slate-100 bg-white">
          <div className="relative">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2.5" />
            <input
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="북마크한 용어, 정의, 회의명 검색..."
              className="w-full pl-9 pr-3 py-1.5 text-xs rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500"
            />
          </div>
        </div>

        {/* Content List */}
        <div className="p-5 overflow-y-auto space-y-3.5 flex-1 bg-slate-50/30">
          {filtered.length > 0 ? (
            filtered.map(({ term, analysisTitle, analysisId }) => (
              <div
                key={term.termId}
                className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm hover:border-amber-300 transition-colors"
              >
                <div className="flex items-start justify-between gap-3 mb-2">
                  <div>
                    <div className="flex items-center gap-2">
                      <h4 className="font-bold text-base text-slate-900">{term.term}</h4>
                      <span className="text-[10px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-500 font-mono">
                        {term.termId}
                      </span>
                    </div>
                    <p className="text-xs text-slate-600 mt-0.5">{term.definition}</p>
                  </div>

                  <button
                    type="button"
                    onClick={() => onToggleBookmark(analysisId, term.termId, false)}
                    className="p-1.5 rounded-lg bg-amber-50 text-amber-600 hover:bg-red-50 hover:text-red-600 transition-colors"
                    title="북마크 해제"
                  >
                    <BookmarkCheck className="w-4 h-4 fill-amber-500" />
                  </button>
                </div>

                <div className="p-2.5 rounded-lg bg-amber-50/40 border border-amber-100 text-xs text-amber-950 mb-2 leading-relaxed">
                  <span className="font-bold mr-1">맞춤형 설명:</span>
                  {term.easyExplanation}
                </div>

                {term.practicalUseCases && (
                  <div className="text-xs text-slate-600 mb-2">
                    <span className="font-semibold text-slate-800 mr-1">실무 활용:</span>
                    {term.practicalUseCases}
                  </div>
                )}

                <div className="pt-2 border-t border-slate-100 flex items-center justify-between text-[11px] text-slate-400">
                  <span className="truncate max-w-xs sm:max-w-md">출처 회의: {analysisTitle}</span>
                  <button
                    type="button"
                    onClick={() => {
                      onClose();
                      onSelectAnalysis(analysisId);
                    }}
                    className="text-blue-600 hover:underline font-semibold flex items-center gap-1"
                  >
                    <span>회의 결과 열기</span>
                    <ArrowRight className="w-3 h-3" />
                  </button>
                </div>
              </div>
            ))
          ) : (
            <div className="p-12 text-center text-slate-400">
              <Bookmark className="w-8 h-8 mx-auto mb-2 opacity-50" />
              <p className="text-sm font-semibold">
                {search ? '검색된 북마크 용어가 없습니다.' : '아직 북마크한 전문용어가 없습니다.'}
              </p>
              <p className="text-xs mt-1">
                회의 분석 결과 화면의 용어 카드에서 북마크 아이콘을 누르면 이곳에 모아집니다.
              </p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
