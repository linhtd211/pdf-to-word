# PDF to Word Pro

Ứng dụng web tiếng Việt chuyển PDF thành DOCX có thể chỉnh sửa. Frontend React + TypeScript + Tailwind + PDF.js; backend FastAPI/Python dùng PyMuPDF, pdfplumber, pdf2docx, python-docx và Tesseract OCR `vie`/`eng`.

## Tính năng

- Kéo thả/chọn PDF, xem trước bằng PDF.js, xem tên/kích thước/số trang.
- Hai chế độ có trade-off rõ ràng: **Ưu tiên chỉnh sửa** tạo Paragraph/Run và bảng Word thật, nối các dòng hiển thị trong cùng đoạn để Word tự xuống dòng khi người dùng viết tiếp; **Ưu tiên bố cục** giữ ngắt dòng/vị trí gần PDF hơn (dùng pdf2docx cho PDF điện tử thông thường). PDF scan/mixed ở chế độ bố cục sẽ được cảnh báo nếu cần fallback OCR.
- Ở nhánh dựng Word editable, ghi hướng giấy riêng cho từng section, suy luận căn lề từ cạnh các dòng, ghép transparency mask của ảnh PDF và dùng ảnh neo theo tọa độ/kích thước bbox. Với file mẫu BV/TYT vừa đối chiếu, bốn ảnh chồng giữ bbox gần như trùng nguồn (sai khác khoảng 0,1 pt).
- Giữ đúng kích thước/hướng trang PDF, kể cả trang phụ lục ngang hoặc khổ tùy chỉnh; chỉ áp dụng lề Nghị định 30 cho trang A4 dọc chuẩn. Bảng hành chính lấy độ rộng lưới/cỡ chữ từ PDF thay vì ép đều hoặc cố định một cỡ; các đoạn/list giữ độ rộng toàn khối và khoảng trắng trong text spans.
- Với PDF điện tử nhận diện được cả Quốc hiệu và Tiêu ngữ, cả hai chế độ dựng lại đầu văn bản thành **một bảng Word 2 cột có độ rộng bằng nhau**, căn cơ quan/số/trích yếu đối diện Quốc hiệu–Tiêu ngữ/ngày; dùng lề trên/dưới 20/25 mm, trái/phải 30/15 mm theo lựa chọn trong Phụ lục I Nghị định 30/2020/NĐ-CP. Quốc hiệu giữ 12 pt và scale ngang 90% để không wrap trong ô phải (thỏa hiệp typography cần kiểm tra trong Word thực tế).
- Mẫu BV/TYT: giữ đúng 2 trang A4 (dọc rồi ngang), footer “Nơi nhận”/khối ký, bảng Word 4×10 và 4×11 với tỷ lệ cột theo PDF. So sánh text chuẩn hóa có 449 token ở cả nguồn và bản render, không thiếu/thừa.
- Phân luồng PDF điện tử và trang scan; OCR scan tiếng Việt thành văn bản Word có thể sửa, không đưa nguyên trang scan vào Word như ảnh.
- Tiến trình lấy từ trạng thái các công đoạn backend; tải DOCX; file đầu vào/kết quả nằm trong thư mục tạm, tự xóa sau khi tải hoặc quá thời hạn.
- Phiên bản hiện tại: tối đa 30 MiB và 100 trang; không đăng nhập, không database.

## Chạy trên Linux/macOS

Yêu cầu Node.js 22, pnpm 11, Python 3.12, Tesseract và traineddata `vie` + `eng`.

```bash
pnpm install
python3 -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
pnpm build
uvicorn backend.app:app --host 0.0.0.0 --port 3000
```

Mở `http://localhost:3000`. Trong quá trình phát triển, sau khi chỉnh giao diện hãy chạy lại `pnpm build` để cập nhật static assets FastAPI phục vụ.

## Kiểm thử

```bash
python tests/generate_fixtures.py
pip install -r requirements-dev.txt
python -m pytest -q
```

Bộ fixture gồm PDF điện tử, công văn có hai khối song song, bảng, nhiều trang, hình ảnh và bản scan tiếng Việt. Xem kết quả đo được và giới hạn tại [`docs/bao-cao-kiem-thu.md`](docs/bao-cao-kiem-thu.md), cùng [ghi chú tham chiếu Nghị định 30](docs/tham-chieu-nghi-dinh-30.md).

## Chạy trên Windows

Xem [`docs/huong-dan-windows.md`](docs/huong-dan-windows.md).

## Đóng gói/deploy

Dockerfile đa giai đoạn build giao diện và chạy FastAPI cùng Tesseract tiếng Việt trên cổng `$PORT` (mặc định 3000). Dùng `docker build -t pdf-to-word-pro .` rồi `docker run --rm -p 3000:3000 pdf-to-word-pro`. Xem [`docs/trien-khai.md`](docs/trien-khai.md) để biết lưu ý về Vercel/serverless và dữ liệu tạm.

## Hạn chế cần biết

PDF là định dạng hiển thị, không luôn chứa ngữ nghĩa Word. Header hành chính chỉ được dựng lại khi PDF có lớp text và trích xuất được Quốc hiệu/Tiêu ngữ; scan/OCR sai hoặc header không điển hình chưa được tự nhận diện. Mẫu BV/TYT và Quyết định 3276 đã được so trực tiếp; không bảo đảm mọi văn bản có bố cục khác sẽ khớp tuyệt đối. Bảng scan, ô gộp phức tạp, font nhúng, ngắt dòng thủ công và bản scan mờ có thể bị tái tạo chưa chính xác; cần rà soát DOCX trước khi ban hành. Microsoft Word desktop chưa được thử trực tiếp (đã kiểm tra bằng python-docx/LibreOffice); bảng OCR từ ảnh chưa được tái tạo thành lưới ô.
