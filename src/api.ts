import type { ConversionJob, ConversionMode, PagePolicy } from './types';

const API_BASE = (import.meta.env.VITE_API_BASE_URL || '').replace(/\/$/, '');
const apiUrl = (path: string) => `${API_BASE}${path}`;

async function readJson<T>(response: Response): Promise<T> {
  const body = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(body.detail || 'Không thể kết nối với máy chủ chuyển đổi.');
  return body as T;
}

export async function startConversion(file: File, mode: ConversionMode, pagePolicy: PagePolicy): Promise<ConversionJob> {
  const form = new FormData();
  form.append('file', file);
  form.append('mode', mode);
  form.append('page_policy', pagePolicy);
  const response = await fetch(apiUrl('/api/conversions'), { method: 'POST', body: form });
  return readJson<ConversionJob>(response);
}

export async function getConversion(id: string): Promise<ConversionJob> {
  const response = await fetch(apiUrl(`/api/conversions/${encodeURIComponent(id)}`), { cache: 'no-store' });
  return readJson<ConversionJob>(response);
}

export async function downloadConversion(id: string): Promise<void> {
  const response = await fetch(apiUrl(`/api/conversions/${encodeURIComponent(id)}/download`));
  if (!response.ok) {
    const body = await response.json().catch(() => ({}));
    throw new Error(body.detail || 'Không thể tải tệp Word.');
  }
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const disposition = response.headers.get('content-disposition') || '';
  const encodedName = disposition.match(/(?:^|;)\s*filename\*\s*=\s*UTF-8''([^;]+)/i)?.[1];
  const plainName = disposition.match(/(?:^|;)\s*filename\s*=\s*"?([^";]+)"?/i)?.[1];
  let fileName = 'tai-lieu.docx';
  if (encodedName) {
    try { fileName = decodeURIComponent(encodedName.trim()); }
    catch { fileName = encodedName.trim(); }
  } else if (plainName) {
    fileName = plainName.trim();
  }
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 30_000);
}
