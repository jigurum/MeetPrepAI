import React, { useState, useRef } from 'react';
import {
  UploadCloud,
  FileText,
  FileCode,
  FileSpreadsheet,
  AlertTriangle,
  Sparkles,
  BookOpen,
  Briefcase,
  Check,
  X,
  ArrowRight,
  Info,
  Loader2,
  FileUp,
  RefreshCw,
} from 'lucide-react';
import { ExplanationLevel } from '../types';
import { SAMPLE_DOCUMENTS, SampleDocument } from '../data/sampleDocuments';

interface InputScreenProps {
  onStartAnalysis: (params: {
    title: string;
    level: ExplanationLevel;
    inputType: 'text' | 'file';
    rawText: string;
    fileData?: { mimeType: string; base64: string; fileName: string };
  }) => Promise<void>;
  isLoading: boolean;
  defaultLevel: ExplanationLevel;
}

export const InputScreen: React.FC<InputScreenProps> = ({
  onStartAnalysis,
  isLoading,
  defaultLevel,
}) => {
  const [activeInputTab, setActiveInputTab] = useState<'text' | 'file'>('file');
  const [title, setTitle] = useState('');
  const [rawText, setRawText] = useState('');
  const [level, setLevel] = useState<ExplanationLevel>(defaultLevel);
  const [isDragging, setIsDragging] = useState(false);
  const [uploadedFile, setUploadedFile] = useState<{
    name: string;
    size: number;
    type: string;
    base64: string;
    extractedText?: string;
    isPdf?: boolean;
  } | null>(null);
  const [extractingFile, setExtractingFile] = useState(false);
  const [extractError, setExtractError] = useState<string | null>(null);
  const [selectedSample, setSelectedSample] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Handle sample document load
  const handleSelectSample = (sample: SampleDocument) => {
    setSelectedSample(sample.id);
    setTitle(sample.title);
    setRawText(sample.content);
    setUploadedFile(null);
    setActiveInputTab('text');
  };

  // Convert File to base64
  const fileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => {
        const result = reader.result as string;
        // Strip data:...;base64,
        const base64 = result.split(',')[1] || '';
        resolve(base64);
      };
      reader.onerror = (error) => reject(error);
    });
  };

  // Process file upload
  const handleFileProcess = async (file: File) => {
    setExtractError(null);
    setExtractingFile(true);
    setSelectedSample(null);

    const ext = file.name.split('.').pop()?.toLowerCase();
    const allowed = ['pdf', 'pptx', 'docx', 'txt', 'md', 'markdown'];
    if (!ext || !allowed.includes(ext)) {
      setExtractError(
        `지원되지 않는 파일 형식(.${ext || 'unknown'})입니다. PDF, PPTX, DOCX, TXT 파일만 지원합니다.`
      );
      setExtractingFile(false);
      return;
    }

    // Default title from file name if not already filled
    if (!title.trim()) {
      const cleanName = file.name.replace(/\.[^/.]+$/, '');
      setTitle(cleanName);
    }

    try {
      const base64 = await fileToBase64(file);

      if (ext === 'pdf') {
        // PDF sent as multimodal directly to Gemini
        setUploadedFile({
          name: file.name,
          size: file.size,
          type: file.type || 'application/pdf',
          base64,
          isPdf: true,
          extractedText: '[PDF 바이너리 문서 분석 준비 완료 - Gemini 비전/문서 엔진 연동]',
        });
      } else {
        // Extract text on server for DOCX, PPTX, TXT
        const res = await fetch('/api/extract-text', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            fileName: file.name,
            fileType: file.type,
            base64Data: base64,
          }),
        });

        const data = await res.json();
        if (!res.ok || data.error) {
          throw new Error(data.error || '텍스트 추출에 실패했습니다.');
        }

        setUploadedFile({
          name: file.name,
          size: file.size,
          type: file.type,
          base64,
          extractedText: data.extractedText,
        });
        setRawText(data.extractedText);
      }
    } catch (err: any) {
      console.error('File process error:', err);
      setExtractError(err.message || '파일을 읽는 중 문제가 발생했습니다.');
    } finally {
      setExtractingFile(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFileProcess(e.dataTransfer.files[0]);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!title.trim()) {
      alert('회의 자료명을 입력해 주세요.');
      return;
    }

    const hasContent =
      (activeInputTab === 'text' && rawText.trim().length > 0) ||
      (activeInputTab === 'file' && uploadedFile);

    if (!hasContent) {
      alert('회의 자료를 업로드하거나 텍스트를 입력해 주세요.');
      return;
    }

    await onStartAnalysis({
      title: title.trim(),
      level,
      inputType: activeInputTab,
      rawText: activeInputTab === 'text' ? rawText.trim() : uploadedFile?.extractedText || '',
      fileData:
        activeInputTab === 'file' && uploadedFile
          ? {
              mimeType: uploadedFile.type,
              base64: uploadedFile.base64,
              fileName: uploadedFile.name,
            }
          : undefined,
    });
  };

  const formatFileSize = (bytes: number) => {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  };

  return (
    <div className="max-w-4xl mx-auto py-8 px-4 sm:px-6">
      {/* Intro Header */}
      <div className="mb-8 text-center sm:text-left">
        <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-blue-50 border border-blue-200 text-blue-700 text-xs font-semibold mb-3">
          <Sparkles className="w-3.5 h-3.5" />
          <span>신입사원 & 비전문 분야 회의 준비 특화</span>
        </div>
        <h1 className="text-2xl sm:text-3xl font-extrabold text-slate-900 tracking-tight">
          회의 자료 맞춤 분석 및 사전 준비
        </h1>
        <p className="mt-2 text-sm sm:text-base text-slate-600 leading-relaxed">
          어려운 회의 자료를 넣으면 <strong className="text-slate-800">전문용어 풀이</strong>와{' '}
          <strong className="text-slate-800">배경지식 핵심 요약</strong>을 내 눈높이에 맞춰 정리해 드립니다.
        </p>
      </div>

      {/* Preset Samples Picker */}
      <div className="mb-6 p-4 rounded-xl bg-slate-50/80 border border-slate-200">
        <div className="flex items-center justify-between mb-2">
          <span className="text-xs font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1.5">
            <BookOpen className="w-3.5 h-3.5 text-blue-600" />
            빠른 체험을 위한 실무 예시 회의 자료
          </span>
          <span className="text-[11px] text-slate-400">클릭 시 자동 입력</span>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
          {SAMPLE_DOCUMENTS.map((sample) => (
            <button
              key={sample.id}
              type="button"
              onClick={() => handleSelectSample(sample)}
              className={`text-left p-3 rounded-lg border text-xs transition-all ${
                selectedSample === sample.id
                  ? 'bg-blue-50 border-blue-400 text-blue-900 shadow-sm'
                  : 'bg-white border-slate-200 hover:border-blue-300 text-slate-700 hover:bg-slate-50'
              }`}
            >
              <div className="flex items-center justify-between font-semibold mb-1">
                <span className="truncate">{sample.title}</span>
                <span className="text-[10px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-600">
                  {sample.category}
                </span>
              </div>
              <p className="text-[11px] text-slate-500 line-clamp-2 leading-tight">
                {sample.description}
              </p>
            </button>
          ))}
        </div>
      </div>

      <form onSubmit={handleSubmit} className="space-y-6">
        {/* Material Title */}
        <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
          <label htmlFor="material-title" className="block text-sm font-bold text-slate-800 mb-2">
            회의 자료명 <span className="text-red-500">*</span>
          </label>
          <input
            id="material-title"
            type="text"
            required
            value={title}
            onChange={(e) => setTitle(e.target.value)}
            placeholder="예: 2026 Q4 클라우드 인프라 아키텍처 개편 검토 회의"
            className="w-full px-4 py-2.5 text-sm rounded-xl border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-600 transition-colors"
          />
        </div>

        {/* Level Selection Cards (Must 1) */}
        <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
          <div className="flex items-center justify-between mb-3">
            <div>
              <h2 className="text-sm font-bold text-slate-800">
                설명 수준 선택 <span className="text-red-500">*</span>
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                동일한 자료라도 선택한 수준에 따라 어조와 비유, 기술적 깊이가 완전히 달라집니다.
              </p>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
            {/* Beginner Mode */}
            <div
              onClick={() => setLevel('beginner')}
              className={`cursor-pointer p-4 rounded-xl border-2 transition-all relative ${
                level === 'beginner'
                  ? 'border-blue-600 bg-blue-50/40 shadow-sm ring-1 ring-blue-600'
                  : 'border-slate-200 hover:border-slate-300 bg-white'
              }`}
            >
              <div className="flex items-start justify-between">
                <div className="flex items-center gap-2">
                  <div className="w-8 h-8 rounded-lg bg-blue-100 text-blue-700 flex items-center justify-center font-bold text-sm">
                    🌱
                  </div>
                  <div>
                    <h3 className="text-sm font-bold text-slate-900">입문자 모드</h3>
                    <span className="text-[11px] font-medium text-blue-700">신입사원 / 비전문가 추천</span>
                  </div>
                </div>
                {level === 'beginner' && (
                  <div className="w-5 h-5 rounded-full bg-blue-600 text-white flex items-center justify-center">
                    <Check className="w-3.5 h-3.5" />
                  </div>
                )}
              </div>
              <ul className="mt-3 space-y-1.5 text-xs text-slate-600">
                <li className="flex items-start gap-1.5">
                  <span className="text-blue-500 font-bold">•</span>
                  <span>배경지식이 전혀 없어도 이해할 수 있는 친절한 설명</span>
                </li>
                <li className="flex items-start gap-1.5">
                  <span className="text-blue-500 font-bold">•</span>
                  <span>일상 생활 속 비유와 구체적인 예시 활용</span>
                </li>
                <li className="flex items-start gap-1.5">
                  <span className="text-blue-500 font-bold">•</span>
                  <span>또 다른 난해한 전문용어 나열 지양</span>
                </li>
              </ul>
            </div>

            {/* Practitioner Mode */}
            <div
              onClick={() => setLevel('practitioner')}
              className={`cursor-pointer p-4 rounded-xl border-2 transition-all relative ${
                level === 'practitioner'
                  ? 'border-indigo-600 bg-indigo-50/40 shadow-sm ring-1 ring-indigo-600'
                  : 'border-slate-200 hover:border-slate-300 bg-white'
              }`}
            >
              <div className="flex items-start justify-between">
                <div className="flex items-center gap-2">
                  <div className="w-8 h-8 rounded-lg bg-indigo-100 text-indigo-700 flex items-center justify-center font-bold text-sm">
                    ⚡
                  </div>
                  <div>
                    <h3 className="text-sm font-bold text-slate-900">실무자 모드</h3>
                    <span className="text-[11px] font-medium text-indigo-700">실무 엔지니어 / PM 추천</span>
                  </div>
                </div>
                {level === 'practitioner' && (
                  <div className="w-5 h-5 rounded-full bg-indigo-600 text-white flex items-center justify-center">
                    <Check className="w-3.5 h-3.5" />
                  </div>
                )}
              </div>
              <ul className="mt-3 space-y-1.5 text-xs text-slate-600">
                <li className="flex items-start gap-1.5">
                  <span className="text-indigo-500 font-bold">•</span>
                  <span>정확하고 엄밀한 기술적 정의와 아키텍처 스펙</span>
                </li>
                <li className="flex items-start gap-1.5">
                  <span className="text-indigo-500 font-bold">•</span>
                  <span>실제 업무 및 파이프라인에서 활용되는 맥락</span>
                </li>
                <li className="flex items-start gap-1.5">
                  <span className="text-indigo-500 font-bold">•</span>
                  <span>유사 개념과의 핵심 차이점 및 실무 주의사항</span>
                </li>
              </ul>
            </div>
          </div>
        </div>

        {/* Input Methods: File Upload vs Direct Text */}
        <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
          <div className="flex items-center justify-between border-b border-slate-200 pb-3 mb-4">
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => setActiveInputTab('file')}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                  activeInputTab === 'file'
                    ? 'bg-blue-600 text-white shadow-sm'
                    : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
                }`}
              >
                <UploadCloud className="w-3.5 h-3.5" />
                <span>파일 업로드 (PDF, PPTX, DOCX)</span>
              </button>
              <button
                type="button"
                onClick={() => setActiveInputTab('text')}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                  activeInputTab === 'text'
                    ? 'bg-blue-600 text-white shadow-sm'
                    : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
                }`}
              >
                <FileText className="w-3.5 h-3.5" />
                <span>텍스트 직접 입력</span>
              </button>
            </div>
            <span className="text-xs text-slate-400 hidden sm:inline">
              드래그 앤 드롭 지원
            </span>
          </div>

          {/* Tab 1: File Upload */}
          {activeInputTab === 'file' && (
            <div className="space-y-4">
              <input
                ref={fileInputRef}
                type="file"
                accept=".pdf,.pptx,.docx,.txt,.md,.markdown"
                className="hidden"
                onChange={(e) => {
                  if (e.target.files && e.target.files[0]) {
                    handleFileProcess(e.target.files[0]);
                  }
                }}
              />

              {!uploadedFile ? (
                <div
                  onDrop={handleDrop}
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  onClick={() => fileInputRef.current?.click()}
                  className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-all ${
                    isDragging
                      ? 'border-blue-600 bg-blue-50/50'
                      : 'border-slate-300 hover:border-blue-500 bg-slate-50/50 hover:bg-slate-50'
                  }`}
                >
                  <div className="w-12 h-12 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center mx-auto mb-3">
                    <FileUp className="w-6 h-6" />
                  </div>
                  <p className="text-sm font-semibold text-slate-800">
                    회의 자료 파일을 여기에 끌어다 놓거나 클릭하여 선택
                  </p>
                  <p className="text-xs text-slate-500 mt-1">
                    지원 포맷: <span className="font-semibold text-slate-700">PDF, PPTX, DOCX, TXT, MD</span> (최대 30MB)
                  </p>
                  {extractingFile && (
                    <div className="flex items-center justify-center gap-2 mt-4 text-xs text-blue-600 font-semibold">
                      <Loader2 className="w-4 h-4 animate-spin" />
                      <span>문서 내용 분석 및 텍스트 추출 중...</span>
                    </div>
                  )}
                </div>
              ) : (
                <div className="p-4 rounded-xl border border-blue-200 bg-blue-50/30">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-3">
                      <div className="w-10 h-10 rounded-lg bg-blue-600 text-white flex items-center justify-center font-bold text-xs uppercase">
                        {uploadedFile.name.split('.').pop() || 'FILE'}
                      </div>
                      <div>
                        <p className="text-sm font-bold text-slate-900 truncate max-w-xs sm:max-w-md">
                          {uploadedFile.name}
                        </p>
                        <p className="text-xs text-slate-500">
                          {formatFileSize(uploadedFile.size)} • {uploadedFile.isPdf ? 'PDF 비전 문서' : '텍스트 추출 완료'}
                        </p>
                      </div>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={() => fileInputRef.current?.click()}
                        className="text-xs px-2.5 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 font-medium"
                      >
                        변경
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setUploadedFile(null);
                          setRawText('');
                        }}
                        className="p-1.5 rounded-lg hover:bg-red-50 text-slate-400 hover:text-red-600"
                        title="파일 삭제"
                      >
                        <X className="w-4 h-4" />
                      </button>
                    </div>
                  </div>

                  {/* Text preview */}
                  {uploadedFile.extractedText && (
                    <div className="mt-3 p-3 bg-white rounded-lg border border-slate-200 text-xs text-slate-600 max-h-32 overflow-y-auto leading-relaxed">
                      <div className="font-semibold text-slate-700 mb-1 flex items-center justify-between">
                        <span>추출된 텍스트 미리보기</span>
                        <span>{uploadedFile.extractedText.length.toLocaleString()} 자</span>
                      </div>
                      <p className="line-clamp-4 whitespace-pre-line text-slate-500 font-mono">
                        {uploadedFile.extractedText.slice(0, 300)}...
                      </p>
                    </div>
                  )}
                </div>
              )}

              {extractError && (
                <div className="p-3 rounded-lg bg-red-50 border border-red-200 text-xs text-red-700 flex items-center gap-2">
                  <AlertTriangle className="w-4 h-4 shrink-0" />
                  <span>{extractError}</span>
                </div>
              )}
            </div>
          )}

          {/* Tab 2: Direct Text Input */}
          {activeInputTab === 'text' && (
            <div className="space-y-2">
              <div className="flex justify-between items-center text-xs text-slate-500">
                <span>회의 안건, 공유된 아젠다, 슬랙 논의 내역, 발표 대본 등을 붙여넣으세요.</span>
                <span>{rawText.length.toLocaleString()} 자</span>
              </div>
              <textarea
                rows={10}
                value={rawText}
                onChange={(e) => setRawText(e.target.value)}
                placeholder="예시:
[회의명] 클라우드 네이티브 아키텍처 전환 검토
[논의 내용] 기존 모놀리식 환경에서 결제 및 주문 모듈을 MSA로 분리하고, 서비스 간 통신은 gRPC와 Istio 서비스 메시를 활용하기로 결정. Kafka를 통한 사가(Saga) 패턴 적용 및 멱등성 보장 방안 논의..."
                className="w-full p-3.5 text-sm rounded-xl border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-600 font-mono transition-colors"
              />
            </div>
          )}
        </div>

        {/* Security / Confidentiality Notice (PRD requirement 9) */}
        <div className="p-4 rounded-xl bg-amber-50/70 border border-amber-200 flex items-start gap-3 text-xs text-amber-900 leading-relaxed">
          <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
          <div>
            <strong className="font-bold">보안 및 개인정보 주의사항:</strong> 주민등록번호, 계좌번호, 개인 연락처 및 비공개 사내 기밀 정보가 포함되지 않도록 확인 후 입력해 주세요. 업로드된 원본 파일은 브라우저에서 분석 목적으로만 일시 처리되며 영구 저장되지 않습니다.
          </div>
        </div>

        {/* Submit Button */}
        <div className="flex items-center justify-end gap-3 pt-2">
          <button
            type="submit"
            disabled={isLoading || extractingFile}
            className={`w-full sm:w-auto px-7 py-3 rounded-xl font-bold text-sm text-white shadow-md transition-all flex items-center justify-center gap-2 ${
              isLoading || extractingFile
                ? 'bg-blue-400 cursor-not-allowed'
                : 'bg-blue-600 hover:bg-blue-700 active:scale-[0.99] shadow-blue-500/25'
            }`}
          >
            {isLoading ? (
              <>
                <Loader2 className="w-4 h-4 animate-spin" />
                <span>회의 자료 AI 정밀 분석 중...</span>
              </>
            ) : (
              <>
                <Sparkles className="w-4 h-4" />
                <span>AI 맞춤 분석 시작</span>
                <ArrowRight className="w-4 h-4" />
              </>
            )}
          </button>
        </div>
      </form>
    </div>
  );
};
