export type ExplanationLevel = 'beginner' | 'practitioner';

export interface TermItem {
  termId: string;
  analysisId: string;
  term: string;
  definition: string;
  easyExplanation: string;
  contextualMeaning: string;
  practicalUseCases: string;
  sourceQuote: string;
  levelNote: string;
  bookmarked: boolean;
}

export interface KeyDiscussionItem {
  title: string;
  description: string;
  originalQuote?: string;
  isFromDocument: boolean;
}

export interface TechnicalConcept {
  name: string;
  role: string;
}

export interface KeyPoint {
  point: string;
  details: string;
  isFact: boolean;
}

export interface BasicConcept {
  concept: string;
  explanation: string;
  relevance: string;
}

export interface WorkProcess {
  stage: string;
  processDescription: string;
}

export interface TechRelationship {
  entityA: string;
  relationship: string;
  entityB: string;
}

export interface BackgroundKnowledge {
  basicConcepts: BasicConcept[];
  workProcesses: WorkProcess[];
  techRelationships: TechRelationship[];
  recommendedPrep: string[];
}

export interface ExpectedQAItem {
  question: string;
  suggestedAnswer: string;
  answerBasis: 'document' | 'inference'; // 자료 근거 vs 실무 추론
  sourceExcerpt?: string;
}

export interface AnalysisSummary {
  topic: string;
  purpose: string;
  purposeConfirmed: boolean;
  keyDiscussionItems: KeyDiscussionItem[];
  technicalConcepts: TechnicalConcept[];
  keyPoints: KeyPoint[];
}

export interface AnalysisRecord {
  analysisId: string;
  analyzedAt: string;
  title: string;
  inputType: string;
  level: ExplanationLevel;
  levelLabel: string;
  summary: AnalysisSummary;
  backgroundKnowledge: BackgroundKnowledge;
  terminology: TermItem[];
  expectedQA: ExpectedQAItem[];
  sourcePreview?: string;
  syncedToSheets?: boolean;
}

export interface AppSettings {
  googleAppsScriptUrl: string;
  autoSyncToSheets: boolean;
  defaultLevel: ExplanationLevel;
}
