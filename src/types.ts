export type ConversionMode = 'layout' | 'editable';
export type PagePolicy = 'source' | 'administrative';
export type JobStatus = 'queued' | 'processing' | 'completed' | 'failed';

export interface QualityFinding {
  severity: 'critical' | 'major' | 'caution' | 'information';
  code: string;
  message: string;
  details?: Record<string, unknown>;
}

export interface QualityReport {
  mode: ConversionMode;
  page_policy: PagePolicy;
  ocr: {
    mean_confidence: number | null;
    pages: Array<{ page: number; word_count: number; mean_confidence: number | null; orientation_correction_degrees: number }>;
  };
  content_comparison: {
    applicable: boolean;
    source_word_count: number;
    docx_word_count: number;
    recall_percent: number | null;
    precision_percent: number | null;
    order_similarity_percent_first_3000_words: number | null;
  };
  word_structure: {
    paragraph_count: number;
    table_count: number;
    table_shapes: string[];
    image_count: number;
    textboxes: number;
    hard_page_breaks: number;
    line_breaks: number;
    sections: Array<{ width_cm: number; height_cm: number; orientation: 'portrait' | 'landscape' }>;
  };
  findings: QualityFinding[];
  render_comparison: { source_pages: number; rendered_pages: number; page_count_delta: number } | null;
}

export interface ConversionJob {
  id: string;
  status: JobStatus;
  progress: number;
  stage: string;
  download_name?: string;
  pages?: number;
  size?: number;
  warnings?: string[];
  error?: string | null;
  kind?: 'electronic' | 'scanned' | 'mixed';
  page_policy?: PagePolicy;
  quality?: QualityReport | null;
}

export interface PdfSummary {
  pages: number;
  title?: string;
}
