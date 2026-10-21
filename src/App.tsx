import { useCallback, useEffect, useRef, useState } from 'react';
import type { DragEvent } from 'react';
import type { PDFDocumentProxy } from 'pdfjs-dist';
import {
  ArrowDownToLine, ArrowRight, ArrowUpFromLine, Check, ChevronLeft, ChevronRight,
  CircleAlert, Clock3, FileCheck2, FileText, Files, LayoutTemplate, LockKeyhole,
  RefreshCw, ShieldCheck, Sparkles, Table2, X,
} from 'lucide-react';
import { downloadConversion, getConversion, startConversion } from './api';
import type { ConversionJob, ConversionMode, PagePolicy } from './types';

const MAX_BYTES = 30 * 1024 * 1024;
const FLOW = ['Đọc cấu trúc PDF', 'Phân tích bố cục', 'Nhận diện văn bản và bảng biểu', 'Tạo tài liệu Word', 'Hoàn tất'];

function formatBytes(bytes: number): string {
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function safeError(error: unknown): string {
  return error instanceof Error ? error.message : 'Đã có lỗi xảy ra. Vui lòng thử lại.';
}

function stageIndex(stage: string): number {
  if (stage.includes('OCR') || stage.includes('Nhận diện')) return 2;
  if (stage.includes('Tạo')) return 3;
  if (stage.includes('Hoàn tất')) return 4;
  if (stage.includes('Phân tích')) return 1;
  return 0;
}

export default function App() {
  const inputRef = useRef<HTMLInputElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [file, setFile] = useState<File | null>(null);
  const [pdf, setPdf] = useState<PDFDocumentProxy | null>(null);
  const [pages, setPages] = useState(0);
  const [pageNumber, setPageNumber] = useState(1);
  const [mode, setMode] = useState<ConversionMode>('editable');
  const [pagePolicy, setPagePolicy] = useState<PagePolicy>('source');
  const [dragging, setDragging] = useState(false);
  const [reading, setReading] = useState(false);
  const [job, setJob] = useState<ConversionJob | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [downloading, setDownloading] = useState(false);

  const clearFile = useCallback(() => {
    setFile(null); setPdf(null); setPages(0); setPageNumber(1); setJob(null); setError('');
    if (inputRef.current) inputRef.current.value = '';
  }, []);

  const acceptFile = useCallback(async (nextFile?: File) => {
    if (!nextFile) return;
    setError(''); setJob(null); setPdf(null); setFile(null); setReading(true);
    if (!nextFile.name.toLowerCase().endsWith('.pdf')) {
      setError('Tệp này không phải PDF. Vui lòng chọn tài liệu có đuôi .pdf.'); setReading(false); return;
    }
    if (nextFile.size > MAX_BYTES) {
      setError('PDF vượt quá giới hạn 30 MiB. Vui lòng chọn tệp nhỏ hơn.'); setReading(false); return;
    }
    if (nextFile.size < 8 || new TextDecoder().decode(await nextFile.slice(0, 5).arrayBuffer()) !== '%PDF-') {
      setError('Không thể xác thực cấu trúc PDF của tệp này.'); setReading(false); return;
    }
    try {
      const [{ getDocument, GlobalWorkerOptions }, { default: workerUrl }] = await Promise.all([
        import('pdfjs-dist'),
        import('pdfjs-dist/build/pdf.worker.min.mjs?url'),
      ]);
      GlobalWorkerOptions.workerSrc = workerUrl;
      const document = await getDocument({ data: await nextFile.arrayBuffer() }).promise;
      setPdf(document); setPages(document.numPages); setFile(nextFile); setPageNumber(1);
    } catch {
      setError('Không mở được PDF để xem trước. Tệp có thể bị hỏng hoặc được bảo vệ bằng mật khẩu.');
    } finally {
      setReading(false);
    }
  }, []);

  useEffect(() => {
    let active = true;
    async function renderPage() {
      if (!pdf || !canvasRef.current) return;
      try {
        const page = await pdf.getPage(pageNumber);
        const viewport = page.getViewport({ scale: 1.15 });
        const canvas = canvasRef.current;
        if (!canvas) return;
        const context = canvas.getContext('2d');
        if (!context) return;
        const scale = Math.min(1, 650 / viewport.width);
        const scaled = page.getViewport({ scale: 1.15 * scale });
        canvas.width = scaled.width; canvas.height = scaled.height;
        await page.render({ canvasContext: context, viewport: scaled }).promise;
      } catch {
        if (active) setError('Không thể hiển thị trang PDF này.');
      }
    }
    void renderPage();
    return () => { active = false; };
  }, [pdf, pageNumber]);

  const runConversion = async () => {
    if (!file) return;
    setError(''); setBusy(true); setJob(null);
    try {
      const first = await startConversion(file, mode, pagePolicy);
      setJob(first);
      let current = first;
      while (current.status === 'queued' || current.status === 'processing') {
        await new Promise((resolve) => window.setTimeout(resolve, 850));
        current = await getConversion(first.id);
        setJob(current);
      }
      if (current.status === 'failed') setError(current.error || 'Không thể chuyển đổi tài liệu.');
    } catch (reason) {
      setError(safeError(reason));
    } finally {
      setBusy(false);
    }
  };

  const handleDownload = async () => {
    if (!job) return;
    setDownloading(true);
    try { await downloadConversion(job.id); }
    catch (reason) { setError(safeError(reason)); }
    finally { setDownloading(false); }
  };

  const onDrop = (event: DragEvent<HTMLElement>) => {
    event.preventDefault(); setDragging(false); void acceptFile(event.dataTransfer.files[0]);
  };

  const activeStep = job ? stageIndex(job.stage) : 0;

  return (
    <div className="app-shell">
      <header className="topbar">
        <a className="brand" href="#top" aria-label="PDF to Word Pro - Trang chủ">
          <span className="brand-mark"><FileText size={22} strokeWidth={2.1} /><span className="brand-arrow"><ArrowRight size={10} /></span></span>
          <span className="brand-name">PDF to Word <b>Pro</b></span>
        </a>
        <nav className="nav-links" aria-label="Điều hướng chính">
          <a className="nav-active" href="#top">Chuyển đổi PDF</a>
          <a href="#gioi-thieu">Giới thiệu</a>
        </nav>
        <div className="trust-chip"><ShieldCheck size={15} /> Tài liệu được xử lý riêng tư</div>
      </header>

      <main id="top">
        <section className="hero-section">
          <div className="hero-copy">
            <div className="eyebrow"><span className="eyebrow-dot" /> CÔNG CỤ CHUYỂN ĐỔI TÀI LIỆU</div>
            <h1>Chuyển PDF sang Word <span>chính xác</span></h1>
            <p className="hero-description">Giữ nguyên bố cục, bảng biểu và định dạng. Chuyển đổi tài liệu thành Word có thể chỉnh sửa.</p>
            <div className="hero-points">
              <span><Check size={15} /> Văn bản có thể chỉnh sửa</span>
              <span><Check size={15} /> Nhận diện tiếng Việt</span>
              <span><Check size={15} /> Không lưu file lâu dài</span>
            </div>
          </div>
          <div className="hero-visual" aria-hidden="true">
            <div className="paper paper-back"><span /><span /><span className="short" /><i /></div>
            <div className="convert-orbit"><ArrowRight size={20} /></div>
            <div className="paper paper-front"><div className="word-badge">W</div><span /><span /><span className="short" /><i /></div>
          </div>
        </section>

        <section className={`workspace ${dragging ? 'workspace-dragging' : ''}`} aria-label="Khu vực chuyển đổi PDF"
          onDragOver={(event) => { event.preventDefault(); setDragging(true); }}
          onDragLeave={(event) => { if (!event.currentTarget.contains(event.relatedTarget as Node)) setDragging(false); }}
          onDrop={onDrop}>
          {!file && !busy && !job && <div className="upload-stage">
            <input ref={inputRef} type="file" accept="application/pdf,.pdf" className="visually-hidden" onChange={(event) => void acceptFile(event.target.files?.[0])} />
            <div className={`dropzone ${dragging ? 'dropzone-active' : ''}`}>
              <div className="upload-icon"><ArrowUpFromLine size={24} /></div>
              <h2>{reading ? 'Đang đọc tài liệu…' : 'Thả tệp PDF vào đây'}</h2>
              <p>hoặc chọn tệp từ thiết bị của bạn</p>
              <button className="button button-primary" onClick={() => inputRef.current?.click()} disabled={reading}>
                <Files size={17} /> Chọn tệp PDF
              </button>
              <div className="upload-limit">PDF · Tối đa 30 MiB · Không cần đăng nhập</div>
            </div>
            <div className="privacy-note"><LockKeyhole size={15} /> Tài liệu chỉ được giữ tạm trong lúc xử lý và tự xóa sau khi tải xuống.</div>
          </div>}

          {file && !busy && job?.status !== 'completed' && <div className="file-workspace">
            <div className="workbench-heading">
              <div><div className="section-kicker">TÀI LIỆU CỦA BẠN</div><h2>Sẵn sàng chuyển đổi</h2></div>
              <button className="icon-button" aria-label="Bỏ tài liệu" onClick={clearFile}><X size={18} /></button>
            </div>
            <div className="workbench-grid">
              <div className="preview-card">
                <div className="preview-toolbar"><div className="preview-label"><FileText size={16} /> Xem trước PDF</div>
                  {pages > 1 && <div className="page-controls"><button className="mini-icon" onClick={() => setPageNumber((page) => Math.max(1, page - 1))} disabled={pageNumber <= 1} aria-label="Trang trước"><ChevronLeft size={16} /></button><span>{pageNumber} / {pages}</span><button className="mini-icon" onClick={() => setPageNumber((page) => Math.min(pages, page + 1))} disabled={pageNumber >= pages} aria-label="Trang sau"><ChevronRight size={16} /></button></div>}
                </div>
                <div className="paper-preview"><canvas ref={canvasRef} /></div>
                <div className="preview-foot"><span>PDF đã tải lên an toàn</span><span>{pages} {pages === 1 ? 'trang' : 'trang'}</span></div>
              </div>
              <div className="settings-panel">
                <div className="file-info-card"><div className="file-icon"><FileText size={20} /></div><div className="file-text"><strong title={file.name}>{file.name}</strong><span>{formatBytes(file.size)} <i /> {pages} trang</span></div><button className="mini-icon remove-file" onClick={clearFile} aria-label="Xóa PDF"><X size={16} /></button></div>
                <div className="section-kicker mode-kicker">CHỌN CÁCH CHUYỂN ĐỔI</div>
                <button className={`mode-card ${mode === 'editable' ? 'mode-selected' : ''}`} onClick={() => setMode('editable')}>
                  <span className="mode-radio">{mode === 'editable' && <i />}</span><span className="mode-icon"><FileCheck2 size={19} /></span>
                  <span className="mode-copy"><strong>Ưu tiên chỉnh sửa</strong><small>Word tự xuống dòng; dễ tiếp tục và bổ sung nội dung.</small><em>Phù hợp tài liệu nhiều chữ</em></span>
                </button>
                <button className={`mode-card ${mode === 'layout' ? 'mode-selected' : ''}`} onClick={() => setMode('layout')}>
                  <span className="mode-radio">{mode === 'layout' && <i />}</span><span className="mode-icon"><LayoutTemplate size={19} /></span>
                  <span className="mode-copy"><strong>Ưu tiên bố cục</strong><small>Cố gắng giữ bố cục và các ngắt dòng chính.</small><em>Phù hợp biểu mẫu, công văn</em></span>
                </button>
                <div className="section-kicker mode-kicker">KHỔ GIẤY VÀ THỂ THỨC</div>
                <button className={`mode-card ${pagePolicy === 'source' ? 'mode-selected' : ''}`} onClick={() => setPagePolicy('source')}>
                  <span className="mode-radio">{pagePolicy === 'source' && <i />}</span><span className="mode-icon"><LayoutTemplate size={19} /></span>
                  <span className="mode-copy"><strong>Giữ theo PDF gốc</strong><small>Giữ kích thước và hướng trang nguồn; lề được suy ra có giới hạn an toàn.</small></span>
                </button>
                <button className={`mode-card ${pagePolicy === 'administrative' ? 'mode-selected' : ''}`} onClick={() => setPagePolicy('administrative')}>
                  <span className="mode-radio">{pagePolicy === 'administrative' && <i />}</span><span className="mode-icon"><ShieldCheck size={19} /></span>
                  <span className="mode-copy"><strong>Chuẩn văn bản hành chính</strong><small>A4, lề theo Nghị định 30 và Times New Roman; bảng rộng chuyển ngang khi cần.</small><em>Kiểm tra lại thể thức trước khi ban hành</em></span>
                </button>
                {error && <div className="inline-error"><CircleAlert size={16} /> {error}</div>}
                <button className="button button-primary convert-button" onClick={() => void runConversion()} disabled={reading || busy}>
                  <span>Chuyển sang Word</span><ArrowRight size={17} />
                </button>
                <div className="settings-note"><LockKeyhole size={13} /> File sẽ tự xóa sau khi tải xuống hoặc hết hạn lưu tạm.</div>
              </div>
            </div>
          </div>}

          {busy && job && <div className="processing-view">
            <div className="processing-heading"><span className="processing-icon"><RefreshCw size={21} /></span><div><div className="section-kicker">ĐANG XỬ LÝ TÀI LIỆU</div><h2>Đang chuyển đổi PDF</h2><p>{file?.name}</p></div></div>
            <div className="progress-meta"><span>{job.stage}</span><strong>{job.progress}%</strong></div>
            <div className="progress-track"><span style={{ width: `${Math.max(4, job.progress)}%` }} /></div>
            <div className="step-list">{FLOW.map((step, index) => <div className={`step-row ${index < activeStep ? 'step-done' : index === activeStep ? 'step-active' : ''}`} key={step}><span className="step-indicator">{index < activeStep ? <Check size={13} /> : index + 1}</span><span>{step}</span>{index === activeStep && <span className="step-live">Đang thực hiện</span>}</div>)}</div>
            <div className="processing-foot"><Clock3 size={15} /> Tiến độ cập nhật theo các bước xử lý thực tế của máy chủ.</div>
            {error && <div className="inline-error"><CircleAlert size={16} /> {error}</div>}
          </div>}

          {job?.status === 'completed' && <div className="success-view">
            <button className="icon-button result-close" onClick={clearFile} aria-label="Chuyển tài liệu khác"><X size={18} /></button>
            <div className="success-mark"><Check size={28} /></div>
            <div className="section-kicker">CHUYỂN ĐỔI HOÀN TẤT</div><h2>Tệp Word đã sẵn sàng</h2>
            <p className="success-copy">Tài liệu đã được chuyển đổi thành công. Bạn có thể tải xuống và tiếp tục chỉnh sửa trong Word.</p>
            <div className="result-file"><div className="docx-icon">W</div><div className="result-file-name"><strong>{job.download_name}</strong><span>{formatBytes(job.size || 0)} <i /> DOCX</span></div><span className="ready-pill"><Check size={13} /> Sẵn sàng</span></div>
            {job.warnings?.map((warning) => <div className="warning-note" key={warning}><CircleAlert size={15} /> {warning}</div>)}
            {job.quality && <div className="quality-panel">
              <div className="quality-heading"><ShieldCheck size={15} /><strong>Kiểm tra chất lượng tự động</strong></div>
              <div className="quality-stats">
                <span><b>{job.quality.content_comparison.applicable && job.quality.content_comparison.recall_percent !== null ? `${job.quality.content_comparison.recall_percent}%` : '—'}</b><small>recall nội dung</small></span>
                <span><b>{job.quality.content_comparison.precision_percent !== null ? `${job.quality.content_comparison.precision_percent}%` : '—'}</b><small>precision nội dung</small></span>
                <span><b>{job.quality.content_comparison.order_similarity_percent_first_3000_words !== null ? `${job.quality.content_comparison.order_similarity_percent_first_3000_words}%` : '—'}</b><small>thứ tự từ đầu tài liệu</small></span>
                <span><b>{job.quality.word_structure.paragraph_count}</b><small>đoạn Word</small></span>
                <span><b>{job.quality.word_structure.table_count}</b><small>bảng chỉnh sửa</small></span>
                <span><b>{job.quality.word_structure.image_count}</b><small>hình ảnh</small></span>
                <span><b>{job.quality.word_structure.textboxes}</b><small>TextBox</small></span>
                <span><b>{job.quality.word_structure.hard_page_breaks}</b><small>ngắt trang cứng</small></span>
                <span><b>{job.quality.word_structure.line_breaks}</b><small>ngắt dòng cố định</small></span>
              </div>
              {job.quality.findings.map((finding) => <div className={`quality-finding quality-${finding.severity}`} key={finding.code}><CircleAlert size={14} />{finding.message}</div>)}
              <p className="quality-disclaimer">Đối chiếu tự động dựa trên văn bản PDF trích xuất được; không thay thế việc đọc soát bản Word trước khi sử dụng.</p>
            </div>}
            <button className="button button-primary download-button" onClick={() => void handleDownload()} disabled={downloading}><ArrowDownToLine size={18} /> {downloading ? 'Đang chuẩn bị…' : 'Tải xuống tài liệu Word'}</button>
            <button className="button button-secondary" onClick={clearFile}>Chuyển đổi tài liệu khác</button>
            <div className="settings-note"><LockKeyhole size={13} /> File tạm sẽ bị xóa sau khi tải xuống.</div>
          </div>}

          {!file && !busy && !job && error && <div className="upload-error"><CircleAlert size={17} />{error}<button onClick={() => setError('')} aria-label="Đóng thông báo"><X size={15} /></button></div>}
        </section>

        <section className="feature-strip" aria-label="Đặc điểm">
          <div className="feature-item"><span className="feature-icon"><Sparkles size={18} /></span><div><strong>Giữ nội dung có thể sửa</strong><small>Không biến cả trang PDF thành ảnh</small></div></div>
          <div className="feature-item"><span className="feature-icon"><Table2 size={18} /></span><div><strong>Hỗ trợ bảng biểu</strong><small>Tạo bảng Word có thể hiệu chỉnh</small></div></div>
          <div className="feature-item"><span className="feature-icon"><ShieldCheck size={18} /></span><div><strong>Ưu tiên riêng tư</strong><small>Tệp tạm tự xóa sau khi sử dụng</small></div></div>
        </section>

        <section className="about-section" id="gioi-thieu">
          <div className="about-heading"><div className="section-kicker">THIẾT KẾ CHO TÀI LIỆU THỰC TẾ</div><h2>Ít thao tác hơn.<br /><span>Word dễ chỉnh sửa hơn.</span></h2></div>
          <div className="about-copy"><p>Chọn chế độ phù hợp với tài liệu của bạn. PDF điện tử được trích xuất trực tiếp; trang scan được xử lý bằng OCR tiếng Việt.</p><p>Bảng phức tạp, văn bản nhiều cột hoặc bản scan mờ có thể cần rà soát sau khi chuyển đổi. Chúng tôi không cam kết giữ nguyên tuyệt đối mọi bố cục.</p><a href="#top">Bắt đầu chuyển đổi <ArrowRight size={15} /></a></div>
        </section>
      </main>

      <footer className="footer"><a className="footer-brand" href="#top"><span className="brand-mark small-mark"><FileText size={17} /><span className="brand-arrow"><ArrowRight size={8} /></span></span>PDF to Word Pro</a><span>Chuyển đổi tài liệu, giữ nội dung trong tầm tay.</span><span className="footer-privacy"><LockKeyhole size={13} /> Tệp được lưu tạm trong quá trình xử lý</span></footer>
    </div>
  );
}
