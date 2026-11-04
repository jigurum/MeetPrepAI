import React, { useState } from 'react';
import {
  FileText,
  Calendar,
  Layers,
  Sparkles,
  BookOpen,
  Bookmark,
  BookmarkCheck,
  CheckCircle2,
  AlertCircle,
  HelpCircle,
  Share2,
  Copy,
  Check,
  Workflow,
  ArrowRight,
  Quote,
  ShieldCheck,
  ExternalLink,
  ChevronDown,
  ChevronUp,
  Search,
  Database,
  ArrowLeft,
  Printer,
  Info,
} from 'lucide-react';
import { AnalysisRecord, TermItem } from '../types';

interface ResultScreenProps {
  analysis: AnalysisRecord;
  onToggleBookmark: (termId: string, newState: boolean) => void;
  onSyncToSheets: () => Promise<void>;
  isSyncingSheets: boolean;
  sheetsUrlConfigured: boolean;
  onBackToHistory: () => void;
  onNewAnalysis: () => void;
  onOpenSheetsSettings: () => void;
}

export const ResultScreen: React.FC<ResultScreenProps> = ({
  analysis,
  onToggleBookmark,
  onSyncToSheets,
  isSyncingSheets,
  sheetsUrlConfigured,
  onBackToHistory,
  onNewAnalysis,
  onOpenSheetsSettings,
}) => {
  const [activeTab, setActiveTab] = useState<'summary' | 'background' | 'terminology' | 'qa'>('summary');
  const [termSearch, setTermSearch] = useState('');
  const [copiedMarkdown, setCopiedMarkdown] = useState(false);
  const [expandedTermId, setExpandedTermId] = useState<string | null>(null);

  const isBeginner = analysis.level === 'beginner';

  // Filtered terminology
  const filteredTerms = (analysis.terminology || []).filter(
    (t) =>
      t.term.toLowerCase().includes(termSearch.toLowerCase()) ||
      t.definition.toLowerCase().includes(termSearch.toLowerCase()) ||
      t.easyExplanation.toLowerCase().includes(termSearch.toLowerCase())
  );

  // Copy full report as Markdown
  const handleCopyMarkdown = () => {
    let md = `# [MeetPrep AI 분석 결과] ${analysis.title}\n`;
    md += `- 분석ID: ${analysis.analysisId}\n`;
    md += `- 분석일시: ${new Date(analysis.analyzedAt).toLocaleString('ko-KR')}\n`;
    md += `- 설명수준: ${analysis.levelLabel} 모드\n\n`;

    md += `## 1. 회의 핵심 요약\n`;
    md += `**회의 주제:** ${analysis.summary.topic}\n\n`;
    md += `**회의 목적:** ${analysis.summary.purpose} (${analysis.summary.purposeConfirmed ? '자료 확인 사실' : '자료 미확인'})\n\n`;

    md += `### 주요 논의 사항\n`;
    analysis.summary.keyDiscussionItems.forEach((item, idx) => {
      md += `${idx + 1}. **${item.title}**: ${item.description}\n`;
      if (item.originalQuote) md += `   > 원문: "${item.originalQuote}"\n`;
    });

    md += `\n## 2. 알아야 할 배경지식\n`;
    analysis.backgroundKnowledge.basicConcepts.forEach((c) => {
      md += `- **${c.concept}**: ${c.explanation} (관련성: ${c.relevance})\n`;
    });

    md += `\n## 3. 맞춤형 전문용어 사전 (${analysis.levelLabel} 모드)\n`;
    analysis.terminology.forEach((t) => {
      md += `### ${t.term}\n`;
      md += `- **정의**: ${t.definition}\n`;
      md += `- **맞춤 설명**: ${t.easyExplanation}\n`;
      md += `- **회의 문맥**: ${t.contextualMeaning}\n`;
      md += `- **실무 사례**: ${t.practicalUseCases}\n\n`;
    });

    navigator.clipboard.writeText(md);
    setCopiedMarkdown(true);
    setTimeout(() => setCopiedMarkdown(false), 2000);
  };

  const handlePrint = () => {
    window.print();
  };

  return (
    <div className="max-w-6xl mx-auto py-6 px-4 sm:px-6 lg:px-8 space-y-6">
      {/* Top Breadcrumb & Quick Actions */}
      <div className="flex flex-wrap items-center justify-between gap-3 text-xs text-slate-500">
        <button
          type="button"
          onClick={onBackToHistory}
          className="inline-flex items-center gap-1.5 hover:text-slate-900 font-medium transition-colors"
        >
          <ArrowLeft className="w-3.5 h-3.5" />
          <span>전체 분석 목록으로</span>
        </button>

        <div className="flex items-center gap-2">
          {/* Print Button */}
          <button
            type="button"
            onClick={handlePrint}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 font-medium transition-colors"
          >
            <Printer className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">인쇄 / PDF</span>
          </button>

          {/* Copy Markdown */}
          <button
            type="button"
            onClick={handleCopyMarkdown}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 font-medium transition-colors"
          >
            {copiedMarkdown ? (
              <>
                <Check className="w-3.5 h-3.5 text-emerald-600" />
                <span className="text-emerald-700 font-semibold">복사 완료!</span>
              </>
            ) : (
              <>
                <Copy className="w-3.5 h-3.5" />
                <span>마크다운 복사</span>
              </>
            )}
          </button>

          {/* Sync to Google Sheets Button */}
          {sheetsUrlConfigured ? (
            <button
              type="button"
              onClick={onSyncToSheets}
              disabled={isSyncingSheets}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold border transition-all ${
                analysis.syncedToSheets
                  ? 'bg-emerald-50 text-emerald-700 border-emerald-300 hover:bg-emerald-100'
                  : 'bg-blue-600 text-white border-blue-600 hover:bg-blue-700'
              }`}
            >
              <Database className="w-3.5 h-3.5" />
              <span>
                {isSyncingSheets
                  ? '시트 동기화 중...'
                  : analysis.syncedToSheets
                  ? 'Google Sheets 저장됨'
                  : 'Google Sheets에 저장'}
              </span>
            </button>
          ) : (
            <button
              type="button"
              onClick={onOpenSheetsSettings}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium border border-dashed border-slate-300 text-slate-600 hover:border-blue-400 hover:text-blue-600 hover:bg-blue-50/50 transition-colors"
            >
              <Database className="w-3.5 h-3.5" />
              <span>Google Sheets 연동하기</span>
            </button>
          )}

          <button
            type="button"
            onClick={onNewAnalysis}
            className="px-3 py-1.5 rounded-lg bg-slate-900 text-white font-medium hover:bg-slate-800 transition-colors"
          >
            + 새 분석
          </button>
        </div>
      </div>

      {/* Material Header Banner */}
      <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <div className="flex flex-wrap items-center gap-2 mb-2">
              <span
                className={`inline-flex items-center gap-1 text-xs font-bold px-2.5 py-0.5 rounded-full ${
                  isBeginner
                    ? 'bg-blue-100 text-blue-800 border border-blue-200'
                    : 'bg-indigo-100 text-indigo-800 border border-indigo-200'
                }`}
              >
                {isBeginner ? '🌱 입문자 모드 (비유와 쉬운 예시)' : '⚡ 실무자 모드 (스펙과 기술 차이)'}
              </span>
              <span className="text-xs px-2 py-0.5 rounded-full bg-slate-100 text-slate-600 border border-slate-200 font-mono">
                {analysis.analysisId}
              </span>
              <span className="text-xs text-slate-400">
                입력 방식: {analysis.inputType}
              </span>
            </div>

            <h1 className="text-xl sm:text-2xl font-extrabold text-slate-900 tracking-tight">
              {analysis.title}
            </h1>

            <div className="flex items-center gap-4 mt-2 text-xs text-slate-500">
              <span className="flex items-center gap-1">
                <Calendar className="w-3.5 h-3.5 text-slate-400" />
                {new Date(analysis.analyzedAt).toLocaleString('ko-KR', {
                  year: 'numeric',
                  month: 'long',
                  day: 'numeric',
                  hour: '2-digit',
                  minute: '2-digit',
                })}
              </span>
              <span>•</span>
              <span className="text-blue-600 font-semibold">
                전문용어 {analysis.terminology?.length || 0}개 추출
              </span>
              <span>•</span>
              <span className="text-indigo-600 font-semibold">
                핵심 논의 {analysis.summary?.keyDiscussionItems?.length || 0}건
              </span>
            </div>
          </div>
        </div>

        {/* Tab Navigation */}
        <div className="flex border-b border-slate-200 mt-6 -mb-6 overflow-x-auto">
          <button
            type="button"
            onClick={() => setActiveTab('summary')}
            className={`flex items-center gap-2 px-5 py-3 text-sm font-bold border-b-2 transition-colors whitespace-nowrap ${
              activeTab === 'summary'
                ? 'border-blue-600 text-blue-600'
                : 'border-transparent text-slate-500 hover:text-slate-900'
            }`}
          >
            <FileText className="w-4 h-4" />
            <span>1. 회의 핵심 요약</span>
          </button>

          <button
            type="button"
            onClick={() => setActiveTab('background')}
            className={`flex items-center gap-2 px-5 py-3 text-sm font-bold border-b-2 transition-colors whitespace-nowrap ${
              activeTab === 'background'
                ? 'border-blue-600 text-blue-600'
                : 'border-transparent text-slate-500 hover:text-slate-900'
            }`}
          >
            <Workflow className="w-4 h-4" />
            <span>2. 알아야 할 배경지식</span>
          </button>

          <button
            type="button"
            onClick={() => setActiveTab('terminology')}
            className={`flex items-center gap-2 px-5 py-3 text-sm font-bold border-b-2 transition-colors whitespace-nowrap ${
              activeTab === 'terminology'
                ? 'border-blue-600 text-blue-600'
                : 'border-transparent text-slate-500 hover:text-slate-900'
            }`}
          >
            <BookOpen className="w-4 h-4" />
            <span>3. 맞춤형 전문용어 사전</span>
            <span className="px-1.5 py-0.2 rounded-full bg-blue-50 text-blue-700 text-xs font-bold border border-blue-200">
              {analysis.terminology?.length || 0}
            </span>
          </button>

          <button
            type="button"
            onClick={() => setActiveTab('qa')}
            className={`flex items-center gap-2 px-5 py-3 text-sm font-bold border-b-2 transition-colors whitespace-nowrap ${
              activeTab === 'qa'
                ? 'border-blue-600 text-blue-600'
                : 'border-transparent text-slate-500 hover:text-slate-900'
            }`}
          >
            <HelpCircle className="w-4 h-4" />
            <span>4. 실전 예상 질의응답</span>
            <span className="text-[10px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-500 font-medium">
              Should
            </span>
          </button>
        </div>
      </div>

      {/* TAB 1: 회의 핵심 요약 (Must 2) */}
      {activeTab === 'summary' && (
        <div className="space-y-6">
          {/* Topic & Purpose Card */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs font-bold uppercase tracking-wider text-slate-400">
                  회의 주제
                </span>
                <span className="text-[11px] px-2 py-0.5 rounded bg-blue-50 text-blue-700 font-semibold border border-blue-200">
                  Topic
                </span>
              </div>
              <p className="text-base font-bold text-slate-900 leading-snug">
                {analysis.summary.topic}
              </p>
            </div>

            <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs font-bold uppercase tracking-wider text-slate-400">
                  회의 목적
                </span>
                <span
                  className={`text-[11px] px-2 py-0.5 rounded font-semibold border ${
                    analysis.summary.purposeConfirmed
                      ? 'bg-emerald-50 text-emerald-700 border-emerald-200'
                      : 'bg-amber-50 text-amber-700 border-amber-200'
                  }`}
                >
                  {analysis.summary.purposeConfirmed ? '자료 확인 사실' : '자료에서 확인되지 않음'}
                </span>
              </div>
              <p
                className={`text-sm leading-relaxed ${
                  analysis.summary.purposeConfirmed ? 'text-slate-800' : 'text-slate-500 italic'
                }`}
              >
                {analysis.summary.purpose}
              </p>
            </div>
          </div>

          {/* Key Discussion Items (with direct quotes & fact distinction) */}
          <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between">
              <div>
                <h3 className="text-base font-bold text-slate-900">
                  주요 논의 사항 및 의사결정 내용
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  자료 원문에 기반한 사실과 핵심 의사결정 맥락을 정리했습니다.
                </p>
              </div>
              <span className="text-xs font-semibold px-2.5 py-1 rounded-full bg-slate-100 text-slate-700">
                {analysis.summary.keyDiscussionItems?.length || 0}개 안건
              </span>
            </div>

            <div className="space-y-3.5">
              {(analysis.summary.keyDiscussionItems || []).map((item, idx) => (
                <div
                  key={idx}
                  className="p-4 rounded-xl border border-slate-200 bg-slate-50/40 hover:bg-slate-50 transition-colors"
                >
                  <div className="flex items-start justify-between gap-3 mb-1.5">
                    <div className="flex items-center gap-2">
                      <span className="w-5 h-5 rounded-full bg-blue-600 text-white text-xs font-bold flex items-center justify-center shrink-0">
                        {idx + 1}
                      </span>
                      <h4 className="font-bold text-sm text-slate-900">{item.title}</h4>
                    </div>
                    <span className="text-[11px] font-medium px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200 shrink-0 flex items-center gap-1">
                      <ShieldCheck className="w-3 h-3" />
                      자료 확인
                    </span>
                  </div>

                  <p className="text-xs sm:text-sm text-slate-700 leading-relaxed pl-7">
                    {item.description}
                  </p>

                  {item.originalQuote && (
                    <div className="mt-2.5 ml-7 p-2.5 rounded-lg bg-white border border-slate-200 text-xs text-slate-600 flex items-start gap-2">
                      <Quote className="w-3.5 h-3.5 text-blue-500 shrink-0 mt-0.5" />
                      <div className="italic text-slate-700">
                        <span className="font-semibold text-slate-500 not-italic mr-1">관련 원문:</span>
                        "{item.originalQuote}"
                      </div>
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>

          {/* Key Takeaways & Technical Concepts */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {/* Key Points */}
            <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
              <h3 className="text-sm font-bold text-slate-900 mb-3 flex items-center gap-2">
                <Sparkles className="w-4 h-4 text-amber-500" />
                이해해야 할 핵심 포인트
              </h3>
              <div className="space-y-2.5">
                {(analysis.summary.keyPoints || []).map((kp, idx) => (
                  <div key={idx} className="p-3 rounded-xl bg-amber-50/40 border border-amber-200/60 text-xs">
                    <div className="font-bold text-slate-900 mb-1 flex items-center gap-1.5">
                      <span className="text-amber-600">✓</span>
                      <span>{kp.point}</span>
                    </div>
                    <p className="text-slate-600 pl-4">{kp.details}</p>
                  </div>
                ))}
              </div>
            </div>

            {/* Technical Concepts */}
            <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
              <h3 className="text-sm font-bold text-slate-900 mb-3 flex items-center gap-2">
                <Layers className="w-4 h-4 text-blue-600" />
                회의에 등장한 중요 기술·업무 개념
              </h3>
              <div className="space-y-2">
                {(analysis.summary.technicalConcepts || []).map((tc, idx) => (
                  <div key={idx} className="p-2.5 rounded-lg border border-slate-200 bg-white text-xs">
                    <div className="font-bold text-blue-800">{tc.name}</div>
                    <div className="text-slate-600 mt-0.5">{tc.role}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* TAB 2: 알아야 할 배경지식 (Must 2) */}
      {activeTab === 'background' && (
        <div className="space-y-6">
          {/* Basic Concepts Needed */}
          <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
            <h3 className="text-base font-bold text-slate-900 mb-1 flex items-center gap-2">
              <BookOpen className="w-4 h-4 text-blue-600" />
              회의 내용을 이해하기 위해 알아야 하는 기초 개념
            </h3>
            <p className="text-xs text-slate-500 mb-4">
              {isBeginner
                ? '신입사원도 쉽게 파악할 수 있도록 기초 정의와 이번 회의와의 연결고리를 정리했습니다.'
                : '실무자가 회의 안건의 기술적 맥락을 빠르게 조망할 수 있는 핵심 선수 지식입니다.'}
            </p>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5">
              {(analysis.backgroundKnowledge.basicConcepts || []).map((concept, idx) => (
                <div
                  key={idx}
                  className="p-4 rounded-xl border border-slate-200 bg-slate-50/50 hover:bg-white hover:shadow-sm transition-all"
                >
                  <div className="font-bold text-sm text-slate-900 mb-1 flex items-center justify-between">
                    <span>{concept.concept}</span>
                    <span className="text-[10px] px-1.5 py-0.5 rounded bg-blue-100 text-blue-700 font-semibold">
                      기초 지식
                    </span>
                  </div>
                  <p className="text-xs text-slate-700 leading-relaxed mb-2">
                    {concept.explanation}
                  </p>
                  <div className="text-[11px] text-blue-800 bg-blue-50/70 p-2 rounded-lg border border-blue-100">
                    <span className="font-bold mr-1">회의 연관성:</span>
                    {concept.relevance}
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Work Process Stages */}
          <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
            <h3 className="text-base font-bold text-slate-900 mb-1 flex items-center gap-2">
              <Workflow className="w-4 h-4 text-indigo-600" />
              관련 업무 프로세스 및 단계별 흐름
            </h3>
            <p className="text-xs text-slate-500 mb-4">
              이번 회의의 안건이 실제 업무와 개발 라이프사이클의 어느 단계에 위치하는지 보여줍니다.
            </p>

            <div className="space-y-3">
              {(analysis.backgroundKnowledge.workProcesses || []).map((proc, idx) => (
                <div
                  key={idx}
                  className="p-3.5 rounded-xl border border-slate-200 bg-white flex items-start gap-3.5"
                >
                  <div className="px-2.5 py-1 rounded-lg bg-indigo-50 border border-indigo-200 text-indigo-700 font-bold text-xs shrink-0 mt-0.5">
                    {proc.stage}
                  </div>
                  <p className="text-xs sm:text-sm text-slate-700 leading-relaxed">
                    {proc.processDescription}
                  </p>
                </div>
              ))}
            </div>
          </div>

          {/* Tech Relationships & Prep Checklist */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {/* Tech Relationships */}
            <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
              <h3 className="text-sm font-bold text-slate-900 mb-3 flex items-center gap-2">
                <Layers className="w-4 h-4 text-emerald-600" />
                핵심 기술 및 개념 간의 관계도
              </h3>
              <div className="space-y-2.5">
                {(analysis.backgroundKnowledge.techRelationships || []).map((rel, idx) => (
                  <div
                    key={idx}
                    className="p-3 rounded-xl border border-slate-200 bg-slate-50/50 text-xs"
                  >
                    <div className="flex items-center gap-2 font-bold text-slate-800 mb-1">
                      <span className="text-blue-700">{rel.entityA}</span>
                      <ArrowRight className="w-3 h-3 text-slate-400" />
                      <span className="text-indigo-700">{rel.entityB}</span>
                    </div>
                    <p className="text-slate-600 pl-1">{rel.relationship}</p>
                  </div>
                ))}
              </div>
            </div>

            {/* Recommended Prep */}
            <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
              <h3 className="text-sm font-bold text-slate-900 mb-3 flex items-center gap-2">
                <CheckCircle2 className="w-4 h-4 text-blue-600" />
                회의 전 미리 알아두면 좋은 체크리스트
              </h3>
              <div className="space-y-2">
                {(analysis.backgroundKnowledge.recommendedPrep || []).map((tip, idx) => (
                  <div
                    key={idx}
                    className="p-3 rounded-xl bg-blue-50/50 border border-blue-200/60 text-xs text-slate-800 flex items-start gap-2"
                  >
                    <span className="font-bold text-blue-600">•</span>
                    <span>{tip}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* TAB 3: 맞춤형 전문용어 사전 (Must 1 & Should) */}
      {activeTab === 'terminology' && (
        <div className="space-y-6">
          {/* Search & Stats Bar */}
          <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col sm:flex-row items-center justify-between gap-3">
            <div className="relative w-full sm:w-80">
              <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2.5" />
              <input
                type="text"
                value={termSearch}
                onChange={(e) => setTermSearch(e.target.value)}
                placeholder="전문용어, 정의 검색..."
                className="w-full pl-9 pr-3 py-1.5 text-xs rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500"
              />
            </div>
            <div className="flex items-center gap-2 text-xs text-slate-500 w-full sm:w-auto justify-end">
              <span>
                총 <strong className="text-slate-900">{analysis.terminology?.length || 0}</strong>개 용어 추출
              </span>
              <span>•</span>
              <span className="text-blue-600 font-semibold">
                {isBeginner ? '입문자 맞춤형 비유 적용' : '실무자 스펙 및 차이점 분석'}
              </span>
            </div>
          </div>

          {/* Term Cards Grid */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {filteredTerms.map((term: TermItem) => {
              const isExpanded = expandedTermId === term.termId;

              return (
                <div
                  key={term.termId}
                  className={`bg-white rounded-2xl border transition-all ${
                    term.bookmarked
                      ? 'border-amber-300 ring-1 ring-amber-300 shadow-sm'
                      : 'border-slate-200 hover:border-slate-300 shadow-sm'
                  }`}
                >
                  <div className="p-5">
                    {/* Card Header */}
                    <div className="flex items-start justify-between gap-2 mb-2">
                      <div>
                        <div className="flex items-center gap-2">
                          <h4 className="font-extrabold text-base text-slate-900">
                            {term.term}
                          </h4>
                          <span className="text-[10px] font-mono px-1.5 py-0.5 rounded bg-slate-100 text-slate-500">
                            {term.termId}
                          </span>
                        </div>
                        <p className="text-xs text-slate-500 mt-0.5 leading-snug">
                          {term.definition}
                        </p>
                      </div>

                      {/* Bookmark button (Should) */}
                      <button
                        type="button"
                        onClick={() => onToggleBookmark(term.termId, !term.bookmarked)}
                        className={`p-2 rounded-xl transition-all shrink-0 ${
                          term.bookmarked
                            ? 'bg-amber-100 text-amber-600 hover:bg-amber-200'
                            : 'bg-slate-100 text-slate-400 hover:text-amber-500 hover:bg-slate-200'
                        }`}
                        title={term.bookmarked ? '북마크 해제' : '개인 용어 사전에 북마크'}
                      >
                        {term.bookmarked ? (
                          <BookmarkCheck className="w-4 h-4 fill-amber-500 text-amber-600" />
                        ) : (
                          <Bookmark className="w-4 h-4" />
                        )}
                      </button>
                    </div>

                    {/* Custom Explanation (Beginner vs Practitioner) */}
                    <div
                      className={`p-3 rounded-xl text-xs mb-3 leading-relaxed ${
                        isBeginner
                          ? 'bg-blue-50/60 border border-blue-200/60 text-blue-950'
                          : 'bg-indigo-50/60 border border-indigo-200/60 text-indigo-950'
                      }`}
                    >
                      <div className="font-bold flex items-center gap-1 mb-1">
                        <Sparkles className="w-3.5 h-3.5 text-blue-600" />
                        <span>{isBeginner ? '쉬운 맞춤 설명 (입문자)' : '실무 스펙 및 정의 (실무자)'}</span>
                      </div>
                      <p>{term.easyExplanation}</p>
                    </div>

                    {/* Level Note / Analogy */}
                    {term.levelNote && (
                      <div className="text-[11px] text-slate-600 bg-slate-50 p-2.5 rounded-lg border border-slate-200 mb-3 flex items-start gap-1.5">
                        <span className="font-bold text-slate-700 shrink-0">
                          {isBeginner ? '💡 비유:' : '🔍 실무 포인트:'}
                        </span>
                        <span>{term.levelNote}</span>
                      </div>
                    )}

                    {/* Contextual Meaning */}
                    <div className="space-y-1.5 text-xs text-slate-700">
                      <div>
                        <strong className="text-slate-900 block font-semibold mb-0.5">
                          회의 문맥에서의 의미:
                        </strong>
                        <p className="text-slate-600 leading-relaxed">{term.contextualMeaning}</p>
                      </div>

                      {/* Practical Use Cases */}
                      <div className="pt-1.5">
                        <strong className="text-slate-900 block font-semibold mb-0.5">
                          실무 활용 사례:
                        </strong>
                        <p className="text-slate-600 leading-relaxed">{term.practicalUseCases}</p>
                      </div>
                    </div>

                    {/* Source quote accordion toggle */}
                    {term.sourceQuote && (
                      <div className="mt-3 pt-3 border-t border-slate-100">
                        <button
                          type="button"
                          onClick={() => setExpandedTermId(isExpanded ? null : term.termId)}
                          className="flex items-center justify-between w-full text-[11px] text-slate-500 hover:text-slate-800 font-medium"
                        >
                          <span className="flex items-center gap-1">
                            <Quote className="w-3 h-3 text-blue-500" />
                            자료 원문 문맥 보기
                          </span>
                          {isExpanded ? (
                            <ChevronUp className="w-3.5 h-3.5" />
                          ) : (
                            <ChevronDown className="w-3.5 h-3.5" />
                          )}
                        </button>

                        {isExpanded && (
                          <div className="mt-2 p-2.5 rounded-lg bg-slate-50 border border-slate-200 text-xs text-slate-600 italic">
                            "{term.sourceQuote}"
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                </div>
              );
            })}
          </div>

          {filteredTerms.length === 0 && (
            <div className="p-12 text-center bg-white rounded-2xl border border-slate-200 text-slate-400">
              <Search className="w-8 h-8 mx-auto mb-2 opacity-50" />
              <p className="text-sm font-semibold">검색된 전문용어가 없습니다.</p>
            </div>
          )}
        </div>
      )}

      {/* TAB 4: 예상 질의응답 (Should 기능) */}
      {activeTab === 'qa' && (
        <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-4">
          <div className="flex items-center justify-between">
            <div>
              <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                <HelpCircle className="w-4 h-4 text-blue-600" />
                회의 중 나올 수 있는 예상 질문 및 모범 답변
              </h3>
              <p className="text-xs text-slate-500 mt-0.5">
                자료에 명시된 근거와 실무적 관점에서의 추론 답변을 구분하여 회의 발언에 자신감을 더합니다.
              </p>
            </div>
          </div>

          <div className="space-y-3.5">
            {(analysis.expectedQA || []).map((qa, idx) => (
              <div
                key={idx}
                className="p-4 rounded-xl border border-slate-200 bg-slate-50/40 hover:bg-slate-50 transition-colors"
              >
                <div className="flex items-start justify-between gap-3 mb-2">
                  <div className="flex items-center gap-2">
                    <span className="w-6 h-6 rounded-full bg-blue-100 text-blue-700 text-xs font-bold flex items-center justify-center shrink-0">
                      Q{idx + 1}
                    </span>
                    <h4 className="font-bold text-sm text-slate-900">{qa.question}</h4>
                  </div>
                  <span
                    className={`text-[11px] font-semibold px-2 py-0.5 rounded-full border shrink-0 ${
                      qa.answerBasis === 'document'
                        ? 'bg-emerald-50 text-emerald-700 border-emerald-200'
                        : 'bg-amber-50 text-amber-700 border-amber-200'
                    }`}
                  >
                    {qa.answerBasis === 'document' ? '자료 근거 답변' : '실무 추론 답변'}
                  </span>
                </div>

                <div className="pl-8 text-xs sm:text-sm text-slate-700 leading-relaxed">
                  <div className="font-semibold text-slate-900 mb-0.5">예상 답변 / 검토 방향:</div>
                  <p>{qa.suggestedAnswer}</p>
                </div>

                {qa.sourceExcerpt && (
                  <div className="mt-2.5 ml-8 p-2 rounded-lg bg-white border border-slate-200 text-xs text-slate-500">
                    <strong className="text-slate-600 mr-1">근거/배경:</strong>
                    {qa.sourceExcerpt}
                  </div>
                )}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};
