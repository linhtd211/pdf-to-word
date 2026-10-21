# Báo cáo kiểm thử PDF to Word Pro

**Ngày cập nhật:** 18/10/2026 · **Phạm vi:** backend, cấu trúc DOCX, render PDF gốc BV/TYT và Quyết định 3276, cùng Preview. Kết quả LibreOffice là bằng chứng tương thích mở/render, không thay thế kiểm tra trên Microsoft Word.

## Kết quả chính

- Frontend `pnpm build` (`tsc -b` + Vite) đạt; backend `pytest -q`: **17 passed** (6 cảnh báo deprecation từ Starlette/SWIG); `pyright --project pyrightconfig.json`: **0 lỗi/0 cảnh báo**; `compileall` thành công.
- DOCX cũ đính kèm được LibreOffice render thành **5 trang**. PDF gốc sau đó do người dùng gửi có **2 trang**: A4 dọc, tiếp theo A4 ngang. Chạy lại converter ở cả hai mode tạo DOCX mở được và render thành đúng **2 trang**, giữ đúng hướng từng trang.
- Luồng Preview HTTPS thực tế: upload trả **202**, job chạy đến **100%**, download trả **200** với DOCX **259,676 byte**; job được dọn khỏi API sau tải (trả 404). Bản tải đã được render lại và đối chiếu các chỉ số dưới đây.
- So sánh text trích từ PDF gốc và PDF render từ DOCX mới sau chuẩn hóa Unicode/whitespace: **449 token nguồn, 449 token đầu ra; không có token thiếu hoặc thừa**.
- Bốn ảnh PDF ở trang đầu có soft-mask được chuyển thành bốn Word drawing anchors. Bbox ảnh trong PDF render mới khớp bbox nguồn trong khoảng **0,1 pt**; dấu đỏ/chữ ký xanh còn chồng nhau, không có nền đen hoặc trang phụ do ảnh phóng lớn.
- Cấu trúc Word đầu ra có bảng header **2×2**, footer **1×2**, và hai bảng phụ lục thật **4×10** và **4×11**. Độ rộng cột riêng được giữ theo PDF (không chia đều); các dòng text trích xuất nằm trong vùng bảng không lặp thành paragraph.
- Trang 1 đã giữ hai khối header, số/ký hiệu, địa danh/ngày, trích yếu hai dòng, body, nơi nhận và khối ký; các field điền rời được ghép vào header. Trang 2 giữ cả hai phụ lục theo thứ tự trên khổ ngang.

## Đối chiếu mẫu BV/TYT

| Hạng mục | DOCX cũ | DOCX mới |
|---|---|---|
| Số trang/hướng giấy | 5 trang; phụ lục bị dàn dọc | 2 trang; A4 dọc rồi A4 ngang |
| Header | Field rời/lặp; trích yếu không nằm trọn trong khối | Bảng Word 2×2; số/ngày ghép đúng; trích yếu hai dòng trong header |
| Dấu/chữ ký | Các ảnh bị đen/phóng lớn, sinh trang tràn | 4 ảnh alpha được neo tại tọa độ PDF, overlap còn giữ |
| Footer | Các khối không cùng hàng | “Nơi nhận” trái, chức danh/chữ ký phải trong bảng Word 1×2 |
| Phụ lục | Hai bảng bị bóp theo chiều dọc | Hai bảng Word chỉnh sửa được 4×10 và 4×11, giữ tỷ lệ cột nguồn |
| Nội dung | Dễ rơi vào đoạn/bảng trùng khi xử lý | So sánh token: 449/449, không thấy thiếu/thừa |

Phần hiển thị được so trực tiếp với hai trang PDF nguồn ở cả hai chế độ. LibreOffice đã mở và render đầu ra. Microsoft Word desktop không có trong môi trường nên chưa được xác minh trực tiếp; line wrapping có thể khác nếu bộ font trên máy người dùng khác, đặc biệt với Times New Roman.

## Hồi quy và fixture

Sáu fixture tổng quát vẫn chạy ở cả hai chế độ: văn bản điện tử, công văn hai khối, bảng số liệu, nhiều trang, ảnh và scan tiếng Việt. Hồi quy tổng hợp mới kiểm tra trang dọc/ngang, header/footer, paragraph canh lề, ảnh trong suốt chồng nhau, bảng 10/11 cột, độ rộng cột và cấu trúc không TextBox. Tệp thật của người dùng chỉ dùng trong phiên để chuyển đổi/đối chiếu, **không đưa vào kho mã**.

| PDF mẫu | Trang | Kết quả |
|---|---:|---|
| `01-van-ban-dien-tu.pdf` | 1 | 356/357 ký tự; DOCX mở được ở cả hai chế độ. |
| `02-cong-van-hanh-chinh.pdf` | 1 | Bảng header native 2×2; giữ các trường chính và render một trang. |
| `03-bang-bieu.pdf` | 1 | Có bảng Word thật và dữ liệu mẫu. |
| `04-nhieu-trang.pdf` | 4 | Bốn section; LibreOffice render 4 trang ở cả hai chế độ. |
| `05-co-hinh-anh.pdf` | 1 | Giữ ảnh thành đối tượng Word; vị trí tổng quát phụ thuộc cấu trúc PDF. |
| `06-scan-tieng-viet.pdf` | 1 | OCR thành paragraph editable; một số dấu/từ OCR vẫn sai, bảng trong ảnh chưa thành lưới ô. |

## Hồi quy PDF Quyết định 3276

PDF do người dùng cung cấp có 7 trang. Bản DOCX lỗi trước đó tràn thành 13 trang do engine ép các trang phụ lục khổ riêng về A4, cố định font bảng lớn hơn PDF nguồn và tính thụt phải theo từng dòng/nhóm nhỏ thay vì toàn khối. Việc lọc bỏ các span chỉ chứa dấu cách còn có thể dính hai từ liền nhau.

Sau sửa, cả hai chế độ **Ưu tiên chỉnh sửa** và **Ưu tiên bố cục** đều render bằng LibreOffice thành đúng **7 trang**. Hai trang đầu là A4 dọc; bốn trang phụ lục giữ kích thước ngang nguồn **905,5 × 612 pt**; trang Phụ lục 2 giữ kích thước dọc riêng **577,3 × 841,7 pt**. DOCX có bảng header thật 2×2 và các bảng phụ lục thật 7×5, 6×5, 7×5, 8×5 và 10×4; dấu/chữ ký văn bản của người ký còn trên trang 2.

Đối chiếu văn bản đã chuẩn hóa Unicode: **2.753 token PDF nguồn, 2.753 token PDF render; không thiếu hoặc thừa token**. Trong DOCX **Ưu tiên chỉnh sửa**, toàn tài liệu không còn `w:br` thủ công: các dòng PDF trong cùng đoạn và trong ô bảng được nối thành văn bản Word tự chảy, đồng thời bỏ right-indent bó hẹp theo mép dòng PDF. Chế độ **Ưu tiên bố cục** vẫn giữ ngắt dòng chính và hình thức gần nguồn; với mẫu này DOCX có 134 hard-break. Cả hai mode vẫn render thành 7 trang. Regression synthetic kiểm tra paragraph, bảng, khổ trang, font, indentation và hai mode; suite hiện **17 passed**. PDF/DOCX thật không được thêm vào kho mã.

## Môi trường và giới hạn

Python 3.12.3; Node 22.13.0; pnpm 11.25.0; Tesseract 5.3.4 (`vie`, `eng`, `osd`); FastAPI 0.115.12; PyMuPDF 1.25.5; pdfplumber 0.11.6; pdf2docx 0.5.8; python-docx 1.1.2; LibreOffice.

Ứng dụng nhận file tối đa 30 MiB/100 trang, không đăng nhập/database; tệp tạm tự xóa khi tải xuống hoặc hết TTL. OCR cho scan mờ/nghiêng và nhận diện bảng trong ảnh còn hạn chế. Nhận diện header hành chính cần có lớp text và các cụm Quốc hiệu/Tiêu ngữ. Xử lý cho form BV/TYT cụ thể đã được kiểm chứng trực tiếp; không suy rộng rằng mọi văn bản hành chính khác sẽ khớp tuyệt đối. Microsoft Word chưa được mở trực tiếp.

## Tham chiếu thể thức

Nguồn: [Nghị định 30/2020/NĐ-CP trên Cổng văn bản Chính phủ](https://vanban.chinhphu.vn/default.aspx?pageid=27160&docid=199378), Phụ lục I (tệp `Số215.doc` người dùng cung cấp). Cấu hình lề và giới hạn nhận diện xem tại [`tham-chieu-nghi-dinh-30.md`](tham-chieu-nghi-dinh-30.md).
