import { AnalysisRecord, AppSettings, ExplanationLevel } from '../types';
import { saveAnalysisToGoogleSheets, toggleBookmarkInGoogleSheets, deleteAnalysisFromGoogleSheets } from './appsScriptTemplate';

const ANALYSES_STORAGE_KEY = 'meetprep_analyses_v1';
const SETTINGS_STORAGE_KEY = 'meetprep_settings_v1';

export const DEFAULT_SETTINGS: AppSettings = {
  googleAppsScriptUrl: '',
  autoSyncToSheets: true,
  defaultLevel: 'beginner',
};

/**
 * Load settings from localStorage
 */
export function getAppSettings(): AppSettings {
  try {
    const raw = localStorage.getItem(SETTINGS_STORAGE_KEY);
    if (!raw) return DEFAULT_SETTINGS;
    return { ...DEFAULT_SETTINGS, ...JSON.parse(raw) };
  } catch {
    return DEFAULT_SETTINGS;
  }
}

/**
 * Save settings to localStorage
 */
export function saveAppSettings(settings: AppSettings): void {
  try {
    localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(settings));
  } catch (e) {
    console.error('Failed to save settings:', e);
  }
}

/**
 * Load all saved analyses from localStorage
 */
export function getStoredAnalyses(): AnalysisRecord[] {
  try {
    const raw = localStorage.getItem(ANALYSES_STORAGE_KEY);
    if (!raw) return [];
    const list: AnalysisRecord[] = JSON.parse(raw);
    return Array.isArray(list) ? list : [];
  } catch (e) {
    console.error('Failed to parse stored analyses:', e);
    return [];
  }
}

/**
 * Save a new analysis to local storage and optionally sync to Google Sheets
 */
export async function saveAnalysis(
  record: AnalysisRecord,
  syncSheets: boolean = true
): Promise<{ record: AnalysisRecord; sheetSyncResult?: { success: boolean; message?: string; error?: string } }> {
  const list = getStoredAnalyses();
  const existingIdx = list.findIndex((a) => a.analysisId === record.analysisId);

  const updatedRecord = { ...record };

  const settings = getAppSettings();
  let sheetSyncResult: { success: boolean; message?: string; error?: string } | undefined;

  if (syncSheets && settings.googleAppsScriptUrl) {
    try {
      const res = await saveAnalysisToGoogleSheets(settings.googleAppsScriptUrl, record);
      if (res.success) {
        updatedRecord.syncedToSheets = true;
        sheetSyncResult = { success: true, message: 'Google Sheets 동기화 완료' };
      } else {
        sheetSyncResult = { success: false, error: res.error };
      }
    } catch (err: any) {
      sheetSyncResult = { success: false, error: err.message };
    }
  }

  if (existingIdx >= 0) {
    list[existingIdx] = updatedRecord;
  } else {
    list.unshift(updatedRecord);
  }

  try {
    localStorage.setItem(ANALYSES_STORAGE_KEY, JSON.stringify(list));
  } catch (e) {
    console.error('Failed to save to localStorage:', e);
  }

  return { record: updatedRecord, sheetSyncResult };
}

/**
 * Delete analysis by ID
 */
export async function deleteAnalysis(analysisId: string): Promise<void> {
  const list = getStoredAnalyses().filter((a) => a.analysisId !== analysisId);
  try {
    localStorage.setItem(ANALYSES_STORAGE_KEY, JSON.stringify(list));
  } catch (e) {
    console.error('Failed to delete from localStorage:', e);
  }

  const settings = getAppSettings();
  if (settings.googleAppsScriptUrl) {
    deleteAnalysisFromGoogleSheets(settings.googleAppsScriptUrl, analysisId).catch(console.error);
  }
}

/**
 * Toggle bookmark on a term
 */
export async function toggleTermBookmark(
  analysisId: string,
  termId: string,
  newBookmarkedState: boolean
): Promise<AnalysisRecord | null> {
  const list = getStoredAnalyses();
  const targetAnalysis = list.find((a) => a.analysisId === analysisId);
  if (!targetAnalysis) return null;

  const targetTerm = targetAnalysis.terminology.find((t) => t.termId === termId);
  if (targetTerm) {
    targetTerm.bookmarked = newBookmarkedState;
  }

  try {
    localStorage.setItem(ANALYSES_STORAGE_KEY, JSON.stringify(list));
  } catch (e) {
    console.error('Failed to update bookmark in localStorage:', e);
  }

  const settings = getAppSettings();
  if (settings.googleAppsScriptUrl) {
    toggleBookmarkInGoogleSheets(settings.googleAppsScriptUrl, termId, newBookmarkedState).catch(console.error);
  }

  return targetAnalysis;
}

/**
 * Get all bookmarked terms across all analyses
 */
export function getAllBookmarkedTerms(): { term: any; analysisTitle: string; analysisId: string }[] {
  const list = getStoredAnalyses();
  const bookmarks: { term: any; analysisTitle: string; analysisId: string }[] = [];

  list.forEach((analysis) => {
    (analysis.terminology || []).forEach((t) => {
      if (t.bookmarked) {
        bookmarks.push({
          term: t,
          analysisTitle: analysis.title,
          analysisId: analysis.analysisId,
        });
      }
    });
  });

  return bookmarks;
}
