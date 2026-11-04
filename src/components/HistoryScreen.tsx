import React, { useState } from 'react';
import {
  Search,
  Calendar,
  Layers,
  ArrowRight,
  Trash2,
  Database,
  CheckCircle2,
  Clock,
  Sparkles,
  BookOpen,
  Filter,
} from 'lucide-react';
import { AnalysisRecord, ExplanationLevel } from '../types';

interface HistoryScreenProps {
  analyses: AnalysisRecord[];
  onSelectAnalysis: (analysis: AnalysisRecord) => void;
  onDeleteAnalysis: (analysisId: string) => void;
  onStartNewAnalysis: () => void;
}

export const HistoryScreen: React.FC<HistoryScreenProps> = ({
  analyses,
  onSelectAnalysis,
  onDeleteAnalysis,
  onStartNewAnalysis,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [filterLevel, setFilterLevel] = useState<'all' | ExplanationLevel>('all');

  // Filter analyses
  const filtered = analyses.filter((item) => {
    const matchesSearch =
      item.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
      item.summary.topic.toLowerCase().includes(searchTerm.toLowerCase()) ||
      (item.terminology || []).some((t) =>
        t.term.toLowerCase().includes(searchTerm.toLowerCase())
      );
    const matchesLevel = filterLevel === 'all' || item.level === filterLevel;
    return matchesSearch && matchesLevel;
  });

  return (
    <div className="max-w-6xl mx-auto py-8 px-4 sm:px-6 lg:px-8 space-y-6">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h1 className="text-2xl font-extrabold text-slate-900 tracking-tight">
            회의 자료 분석 기록
          </h1>
          <p className="text-xs sm:text-sm text-slate-500 mt-1">
            이전에 분석했던 회의 안건과 맞춤형 전문용어 목록을 언제든 다시 열람할 수 있습니다.
          </p>
        </div>
        <button
          type="button"
          onClick={onStartNewAnalysis}
          className="inline-flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold shadow-sm shadow-blue-500/20 transition-all self-start sm:self-auto"
        >
          <Sparkles className="w-4 h-4" />
          <span>새 회의 자료 분석하기</span>
        </button>
      </div>

      {/* Search and Filters */}
      <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col sm:flex-row items-center justify-between gap-3">
        <div className="relative w-full sm:w-96">
          <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2.5" />
          <input
            type="text"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            placeholder="자료명, 주제, 전문용어 검색..."
            className="w-full pl-9 pr-3 py-1.5 text-xs rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-colors"
          />
        </div>

        {/* Level filter tabs */}
        <div className="flex items-center gap-1.5 w-full sm:w-auto justify-end">
          <Filter className="w-3.5 h-3.5 text-slate-400 mr-1 hidden sm:inline" />
          <button
            type="button"
            onClick={() => setFilterLevel('all')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
              filterLevel === 'all'
                ? 'bg-slate-900 text-white'
                : 'text-slate-600 hover:bg-slate-100'
            }`}
          >
            전체 ({analyses.length})
          </button>
          <button
            type="button"
            onClick={() => setFilterLevel('beginner')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
              filterLevel === 'beginner'
                ? 'bg-blue-600 text-white'
                : 'text-slate-600 hover:bg-slate-100'
            }`}
          >
            입문자 모드
          </button>
          <button
            type="button"
            onClick={() => setFilterLevel('practitioner')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
              filterLevel === 'practitioner'
                ? 'bg-indigo-600 text-white'
                : 'text-slate-600 hover:bg-slate-100'
            }`}
          >
            실무자 모드
          </button>
        </div>
      </div>

      {/* Analysis Cards List */}
      {filtered.length > 0 ? (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          {filtered.map((item) => {
            const isBeginner = item.level === 'beginner';

            return (
              <div
                key={item.analysisId}
                className="bg-white rounded-2xl border border-slate-200 hover:border-blue-300 hover:shadow-md transition-all p-5 flex flex-col justify-between group cursor-pointer"
                onClick={() => onSelectAnalysis(item)}
              >
                <div>
                  {/* Top Badges */}
                  <div className="flex items-center justify-between gap-2 mb-2">
                    <div className="flex items-center gap-1.5">
                      <span
                        className={`text-[11px] font-bold px-2.5 py-0.5 rounded-full ${
                          isBeginner
                            ? 'bg-blue-50 text-blue-700 border border-blue-200'
                            : 'bg-indigo-50 text-indigo-700 border border-indigo-200'
                        }`}
                      >
                        {isBeginner ? '🌱 입문자 모드' : '⚡ 실무자 모드'}
                      </span>
                      <span className="text-[10px] font-mono px-2 py-0.5 rounded bg-slate-100 text-slate-500">
                        {item.analysisId}
                      </span>
                    </div>

                    <div className="flex items-center gap-1.5">
                      {item.syncedToSheets && (
                        <span
                          className="flex items-center gap-1 text-[10px] font-medium text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-full border border-emerald-200"
                          title="Google Sheets에 동기화 완료"
                        >
                          <Database className="w-3 h-3" />
                          시트 연동됨
                        </span>
                      )}

                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          if (confirm(`'${item.title}' 분석 기록을 삭제하시겠습니까?`)) {
                            onDeleteAnalysis(item.analysisId);
                          }
                        }}
                        className="p-1 rounded-lg text-slate-400 hover:text-red-600 hover:bg-red-50 transition-colors"
                        title="기록 삭제"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>

                  {/* Title */}
                  <h3 className="font-bold text-base text-slate-900 group-hover:text-blue-600 transition-colors line-clamp-1 mb-1.5">
                    {item.title}
                  </h3>

                  {/* Topic snippet */}
                  <p className="text-xs text-slate-600 line-clamp-2 leading-relaxed mb-3">
                    {item.summary.topic}
                  </p>

                  {/* Sample Term Chips */}
                  <div className="flex flex-wrap gap-1 mb-4">
                    {(item.terminology || []).slice(0, 4).map((t, i) => (
                      <span
                        key={i}
                        className="text-[10px] px-2 py-0.5 rounded-md bg-slate-100 text-slate-700 font-medium border border-slate-200"
                      >
                        {t.term}
                      </span>
                    ))}
                    {(item.terminology || []).length > 4 && (
                      <span className="text-[10px] px-1.5 py-0.5 rounded-md bg-slate-50 text-slate-400">
                        +{item.terminology.length - 4}개 더
                      </span>
                    )}
                  </div>
                </div>

                {/* Footer Metadata */}
                <div className="pt-3 border-t border-slate-100 flex items-center justify-between text-xs text-slate-400">
                  <div className="flex items-center gap-1.5">
                    <Clock className="w-3.5 h-3.5" />
                    <span>
                      {new Date(item.analyzedAt).toLocaleDateString('ko-KR', {
                        month: 'short',
                        day: 'numeric',
                        hour: '2-digit',
                        minute: '2-digit',
                      })}
                    </span>
                    <span>•</span>
                    <span>{item.inputType}</span>
                  </div>

                  <span className="inline-flex items-center gap-1 text-blue-600 font-bold group-hover:translate-x-0.5 transition-transform">
                    <span>결과 보기</span>
                    <ArrowRight className="w-3.5 h-3.5" />
                  </span>
                </div>
              </div>
            );
          })}
        </div>
      ) : (
        <div className="bg-white rounded-2xl border border-slate-200 p-12 text-center shadow-sm">
          <BookOpen className="w-10 h-10 text-slate-300 mx-auto mb-3" />
          <h3 className="text-base font-bold text-slate-800">
            {searchTerm ? '검색 결과가 없습니다.' : '아직 저장된 분석 기록이 없습니다.'}
          </h3>
          <p className="text-xs text-slate-500 mt-1 max-w-sm mx-auto">
            {searchTerm
              ? '다른 검색어를 입력하거나 필터를 해제해 보세요.'
              : '어려운 회의 자료를 업로드하여 첫 번째 맞춤형 분석을 시작해 보세요!'}
          </p>
          <button
            type="button"
            onClick={onStartNewAnalysis}
            className="mt-5 inline-flex items-center gap-2 px-4 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold shadow-sm transition-colors"
          >
            <Sparkles className="w-3.5 h-3.5" />
            <span>회의 자료 분석 시작하기</span>
          </button>
        </div>
      )}
    </div>
  );
};
