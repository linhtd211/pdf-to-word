# Kế hoạch triển khai — PDF to Word Pro

## Phạm vi và kiến trúc

Sản phẩm là một web app tiếng Việt chạy thử được, có luồng chuyển đổi thực tế ở backend; không đăng nhập, không database, không lưu trữ dài hạn. Chọn React + TypeScript + Vite + Tailwind CSS + Lucide ở frontend; FastAPI/Python ở backend. Frontend được build thành static assets và FastAPI phục vụ cả UI lẫn `/api/*` trên cổng 3000 để preview và triển khai container không cần phụ thuộc serverless. PDF.js (`pdfjs-dist`) dùng để đọc và xem trước file trên client.

Backend nhận file bằng `POST /api/conversions` cùng chế độ chuyển đổi; kiểm tra magic bytes PDF, kích thước và số trang; trả `job_id`. `GET /api/conversions/{job_id}` cung cấp tiến độ/trạng thái được cập nhật bởi từng bước xử lý thật, không dùng bộ đếm thời gian giả; endpoint tải DOCX trả kết quả. Tệp đầu vào và kết quả được giữ trong thư mục tạm riêng cho job, dọn khi tải xuống và tự hết hạn sau 20 phút; không dùng database hay object storage. Giới hạn phiên bản đầu: 30 MiB, 100 trang/job, thông báo rõ nếu vượt giới hạn hoặc PDF mã hóa/hỏng.

## Bộ máy chuyển đổi

- Dùng PyMuPDF để xác thực, đọc metadata/kích thước trang, trích xuất văn bản và xác định PDF điện tử hay scan (hoặc hỗn hợp). OCR chỉ áp dụng cho trang thiếu văn bản; không OCR lại nội dung điện tử đã trích xuất được.
- Với **Ưu tiên bố cục**, dùng `pdf2docx` cho trang điện tử để giữ luồng chữ/hình/bảng có thể chỉnh sửa, thêm bảo toàn khổ giấy và ngắt trang; với trang scan, render từng trang rồi dùng Tesseract OCR `vie` + `eng`, tái tạo các đoạn Word thực sự, không nhúng nguyên trang thành ảnh.
- Với **Ưu tiên chỉnh sửa**, dựng Paragraph/Run/Table/Section từ text span, vị trí, thuộc tính font và bảng phát hiện được bằng PyMuPDF/pdfplumber/python-docx; giữ hình ảnh phù hợp; hạn chế định vị tuyệt đối và TextBox. Trang scan dùng văn bản OCR để sinh các đoạn Word, ưu tiên thứ tự đọc và khả năng sửa.
- Các bước tiến độ phản ánh công việc đã xảy ra: nhận file/kiểm tra cấu trúc, phân tích bố cục và văn bản/bảng/OCR, tạo DOCX, hoàn tất. API báo lỗi có thể hiểu và dọn tệp tạm.
- Tính khả thi thực tế và độ chính xác có giới hạn: cách biểu diễn PDF không lưu ngữ nghĩa Word gốc; bảng scan, bố cục nhiều cột và OCR mờ cần người dùng rà soát. Không tuyên bố chính xác tuyệt đối; báo rõ hạn chế trong ứng dụng và tài liệu.

## Thiết kế trải nghiệm

**Design Movement:** Modern Swiss SaaS / Swiss International Style, tiết chế và giàu khoảng trắng.

**Nguyên tắc:** (1) tải file và chọn cách chuyển đổi là trung tâm; (2) thông tin an toàn, giới hạn và lỗi nói rõ; (3) bản xem trước và kết quả luôn gắn với đúng tài liệu; (4) cấu trúc nhẹ, ưu tiên thao tác nhanh trên mọi cỡ màn hình.

**Màu sắc:** nền trắng và xám lạnh để giảm nhiễu; xanh dương làm màu hành động/tin cậy; màu xanh lá chỉ dùng cho hoàn tất; hổ phách cho cảnh báo OCR/hạn chế. Góc bo nhẹ, viền mảnh, đổ bóng tiết chế.

**Bố cục:** thanh điều hướng gọn; hero lệch nhẹ với headline/mô tả cạnh minh họa tài liệu vector nhỏ; vùng kéo thả lớn là điểm tập trung. Sau khi chọn file, chuyển thành workbench gồm preview PDF và panel thông tin/chế độ; trên mobile xếp dọc. Kết quả có trạng thái hoàn tất, tên/dung lượng và tải xuống.

**Dấu ấn:** khung trang giấy với điểm nhấn xanh, dải tiến trình 5 bước có trạng thái, nhãn nhỏ phân biệt văn bản điện tử/scan.

**Tương tác/hoạt ảnh:** kéo-thả đổi viền và nền trong thời gian hover; trạng thái hover/focus rõ, tuân thủ reduced motion; chuyển trang thái 150–220 ms, không animate chỉ số tiến độ giả. Tiến độ chỉ thay đổi khi backend cập nhật bước đã hoàn tất.

**Typography:** Inter cho tiêu đề/nhãn, hệ thống sans hỗ trợ dấu tiếng Việt làm fallback; headline đậm vừa, nội dung 14–16 px, khoảng cách dòng thoáng.

**Thương hiệu:** “PDF to Word Pro” — chuyển tài liệu văn phòng Việt Nam thành Word chỉnh sửa được; tin cậy, rõ ràng, thực dụng. Giọng văn ngắn gọn, bình tĩnh; ví dụ: “Tải PDF lên để bắt đầu chuyển đổi.” / “Bố cục phức tạp có thể cần kiểm tra lại sau khi tải Word.” Wordmark dùng ký hiệu trang giấy với mũi tên chuyển đổi và chữ W, không dùng logo mặc định; xanh riêng #2563EB.

## Cấu trúc dự kiến

- `src/`: React + TypeScript — app, các thành phần tải file, PDF preview, chọn mode, theo dõi job/kết quả; API client; style Tailwind.
- `backend/app.py`: FastAPI — upload, trạng thái, tải xuống, static UI và xử lý giới hạn/xóa file.
- `backend/converter.py`: định dạng/đọc PDF, trích xuất/OCR, tạo DOCX theo 2 chế độ.
- `backend/jobs.py`: registry job trong bộ nhớ, tiến trình từng bước, TTL và dọn thư mục tạm.
- `public/manus-routes.json`: manifest route trang web; `/api/*` là API và không đưa vào manifest.
- `tests/` và `tests/fixtures/`: bộ PDF kiểm thử tổng hợp và kiểm thử backend.
- `Dockerfile`, `requirements.txt`, `package.json`, `pnpm-workspace.yaml`: cài/build/run frontend + Python + Tesseract OCR `vie`.
- `docs/`: hướng dẫn chạy Windows/triển khai, báo cáo test và hạn chế.
- `app.config.ts`: logoUrl URL HTTPS bền vững của logo trước checkpoint.

## Cài đặt và vận hành

Cài dependencies frontend bằng pnpm đã pin trong `packageManager`, chính sách script lifecycle được khai báo trước khi cài. Backend dùng Python 3.12, FastAPI/Uvicorn, PyMuPDF, pdfplumber, pdf2docx, python-docx, pytesseract/Pillow; Tesseract và dữ liệu ngôn ngữ `vie`/`eng` là gói hệ thống. Khởi chạy backend bind `0.0.0.0:3000`; Vite build frontend vào thư mục static backend. Dockerfile đa giai đoạn build giao diện và cài backend/OCR để chạy container. Preview không đồng nghĩa với xuất bản; chưa cấu hình hoặc bật auto-publish.

## Bộ kiểm thử/deliverable

Sinh sáu PDF kiểm thử đại diện: văn bản điện tử, công văn hai khối, bảng biểu, nhiều trang, văn bản có hình ảnh, scan tiếng Việt. Kiểm tra API upload/trạng thái/tải, cấu trúc DOCX/paragraph/table, text/dấu tiếng Việt, xóa dữ liệu tạm và mở DOCX bằng thư viện Word. Ghi lại kết quả thực tế và các hạn chế. Bàn giao Web app Preview, mã nguồn, PDF kiểm thử, hướng dẫn Windows/triển khai, báo cáo hoàn tất và hạn chế.


## Điều chỉnh theo yêu cầu bổ sung: Nghị định 30/2020/NĐ-CP

- Dùng Phụ lục I Nghị định 30/2020/NĐ-CP (người dùng cung cấp; đối chiếu Cổng văn bản Chính phủ) làm chuẩn cho PDF điện tử có Quốc hiệu và Tiêu ngữ nhận diện được.
- Thay nhánh `pdf2docx` cho trường hợp này bằng tái dựng editable có bảng Word native 2 cột bằng nhau: cơ quan ban hành đối diện Quốc hiệu–Tiêu ngữ; dòng dưới căn số/ký hiệu với địa danh/ngày tháng; trích yếu công văn theo dưới số. Không giữ lại header gốc dưới dạng paragraph để tránh trùng nội dung.
- Áp dụng khổ A4 (xoay ngang nếu nguồn ngang), lề 20/25/30/15 mm (trên/dưới/trái/phải), Times New Roman màu đen 13 pt cho nội dung, đoạn cách 6 pt và dòng đơn; Quốc hiệu 12 pt đậm, Tiêu ngữ 13 pt đậm căn giữa có đường kẻ dưới, ngày tháng nghiêng. Đây là lựa chọn nằm trong các khoảng Phụ lục I.
- Chỉ tự động áp dụng khi hai dòng Quốc hiệu–Tiêu ngữ được trích xuất từ PDF; scan/OCR sai hoặc header không điển hình tiếp tục theo luồng cũ và không tự suy đoán. Giữ cảnh báo rà soát số liệu/định dạng hành chính.
- Tài liệu căn cứ và giới hạn: `docs/tham-chieu-nghi-dinh-30.md`; hồi quy fixture `02-cong-van-hanh-chinh.pdf` kiểm tra cả hai chế độ.


## Bổ sung theo DOCX người dùng báo sai form (16/10/2026)

Tệp DOCX mẫu cho thấy đầu ra hiện tại làm sai nhiều lớp bố cục: thân thư bị căn giữa/giãn sai; ảnh chữ ký và dấu có nền đen, bị phóng lớn; trang mẫu nằm ngang nhưng cờ hướng giấy vẫn Portrait; nội dung tràn thành trang trắng; bảng phụ lục 10/11 cột bị ép vào trang dọc. Chỉ nhận DOCX đầu ra, chưa có PDF gốc đối ứng, do đó chưa thể khẳng định vị trí chính xác trên từng trang nguồn.

Hướng sửa: (1) ghi rõ `w:orient` đồng bộ với khổ trang tại từng section, kiểm chứng A4 ngang; (2) phân tích alignment từ các dòng trong block thay vì tâm bounding box toàn block; (3) ghép soft-mask của ảnh PDF vào alpha khi tạo PNG Word, dùng kích thước/điểm đặt tương đối từ bbox thay vì cố định 5.6 inch; (4) giữ ngắt section theo trang và các bảng thật. Tạo regression synthetic có text block canh trái, ảnh alpha chữ ký/dấu và trang landscape; sau khi nhận PDF gốc thì rerun end-to-end và tinh chỉnh theo đối chiếu thật.


### Kết quả vòng sửa form

Đã triển khai cờ orientation theo từng section; phân biệt alignment từ bbox từng dòng; ghép ảnh mask vào alpha, neo ảnh ở bbox nguồn; lọc text trùng lưới bảng và giữ width từng cột; ghép field số/ngày/trích yếu, dựng footer hai cột và giữ khoảng cách dọc theo tọa độ Y. Test synthetic bao phủ header/footer, ảnh alpha chồng, A4 dọc/ngang và bảng 10/11 cột. Sau khi nhận PDF gốc BV/TYT, cả hai mode được chạy lại: đúng 2 trang A4 theo hướng nguồn, 4 ảnh với bbox sai khác khoảng 0,1 pt, 2 bảng 4×10/4×11 giữ width không đều; text sau chuẩn hóa là 449 token ở cả nguồn và đầu ra. `pytest`: 11 passed; Pyright 0 lỗi/0 cảnh báo.

Bổ sung cho header Nghị định 30: Quốc hiệu giữ cỡ Times New Roman 12 pt nhưng scale ngang 90% để vừa một dòng trong cột phải của bảng cân bằng. Scale này được thử trên LibreOffice synthetic render; đã ghi rõ đây là thỏa hiệp của ứng dụng, không phải quy định riêng của Nghị định.


## Đối chiếu PDF gốc BV/TYT (nhận 16/10/2026)

PDF gốc có 2 trang: trang đầu A4 dọc, trang phụ lục A4 ngang. Trang đầu có Quốc hiệu/cơ quan song song, số/ngày điền thành các span nằm trên dòng khác, trích yếu 2 dòng, thân văn bản, footer song song và 4 image blocks trong suốt chồng ở vùng chữ ký/dấu. Trang 2 có 2 bảng (4×10 và 4×11) với độ rộng cột không đều. DOCX cũ 5 trang và hiển thị sai. Luồng đã ghép field/trích yếu, neo ảnh, lọc text trong bảng, giữ column widths và dựng footer hai cột. Cả hai mode xuất 2 trang; bốn bbox ảnh khớp khoảng 0,1 pt và word-token multiset là 449/449 không thiếu/thừa. PDF/DOCX nguồn người dùng chỉ dùng trong phiên, không đưa vào kho mã.

## Kế hoạch mở rộng theo yêu cầu hoàn thiện 18/10/2026

Tiếp tục trên đúng project, giữ giao diện hiện có và không đổi kiến trúc lớn. Dùng `/home/ubuntu/upload/3276.QĐ-BYT.pdf` cùng DOCX lỗi `/home/ubuntu/upload/pasted_file_2DSuOE_3276.QĐ-BYT.docx` làm cặp đối chiếu; tệp người dùng đã sửa không được dùng làm nguồn cho DOCX mới.

- **Tái tạo đoạn theo ngữ nghĩa:** gom line spans liên tiếp và text blocks liền kề dựa trên khoảng cách dòng/đoạn, điểm kết câu, thụt lề, alignment, font/style, danh sách và khoảng cột. Nối wrap thành một paragraph Word trong editable mode; ranh giới Điều/tiêu đề/danh sách/đoạn có khoảng cách rõ vẫn tách. Giữ Run format (đậm/nghiêng/gạch chân, font/cỡ/màu), không chuyển dòng PDF thành paragraph hoặc `w:br` vô điều kiện. Văn bản trong mỗi ô bảng được tái dựng độc lập.
- **Page policy tách khỏi paragraph mode:** thêm lựa chọn `source` (giữ kích thước trang, hướng và lề nguồn khi có bằng chứng đáng tin) và `administrative` (chuẩn hóa trang văn bản theo Nghị định 30). Giữ cả hai hành vi cấu trúc editable/layout hiện tại, không phá request client cũ. Page geometry không được suy lề dưới/phải cực lớn chỉ từ mép bbox nội dung; không thêm khoảng trắng để giả vị trí.
- **Sections và pagination:** chỉ tạo section khi kích thước giấy/hướng hoặc cấu trúc trang thực sự đổi; không tạo section cho từng trang đồng dạng. Trong chế độ editable, bỏ ngắt trang cứng tùy tiện giữa đoạn để nội dung có thể chảy; ngắt chỉ tại ranh giới layout cần thiết. Giữ page break theo nguồn chỉ khi policy/structure yêu cầu. Thử LibreOffice kiểm tra không có trang trắng.
- **Bảng và hành chính:** tiếp tục Table Word native, mở rộng lưới/cạnh/ô gộp, căn chỉnh, shading và màu chữ có thể nhận diện; nối wrap trong ô. Giữ header, trường số/ngày, trích yếu, căn cứ, Điều, footer, người ký, con dấu/ảnh theo vị trí tương đối. Page policy hành chính dùng giấy A4 21 × 29,7 cm, lề trên/dưới 2–2,5 cm, trái 3–3,5 cm, phải 1,5–2 cm, Times New Roman Unicode màu đen; bảng rộng được A4 landscape với lề hợp lý. Chỉ tuyên bố đúng thể thức khi cả thành phần, không chỉ lề, đã được kiểm tra.
- **PDF/OCR:** phân loại điện tử/scan/mixed theo từng trang; ưu tiên text gốc; OCR tiếng Việt chỉ các vùng scan cần thiết; khảo sát xoay/deskew, nâng chất lượng ảnh độ phân giải thấp và giữ confidence thấp trong cảnh báo kiểm tra. Không sửa/suy đoán nội dung pháp lý; khu vực không chắc cần thể hiện là chưa chắc, không âm thầm thay thế text.
- **Kiểm tra chất lượng:** thêm module/CLI có thể chạy sau chuyển đổi, đối chiếu text PDF–DOCX (mất/lặp/thứ tự), số bảng/ảnh, native-table structure, paragraph count/breaks, page size/margins/sections, render DOCX sang PDF bằng LibreOffice và sinh báo cáo có các sai lệch/cảnh báo/mức độ nghiêm trọng. Không tuyên bố hoàn hảo; phân biệt chuẩn hóa chủ động với lỗi. Thêm regression synthetic và bài kiểm thử tài liệu Quyết định 3276 từ PDF nguồn.
- **Backend và bảo vệ tài nguyên:** giữ FastAPI/ThreadPool/container; bổ sung giới hạn hàng đợi/concurrency, timeout có thể quan sát, recovery rõ ràng cho job bị gián đoạn, giới hạn upload/page hiện tại, dọn tệp/orphan đúng hạn và lỗi an toàn. Không cam kết phục hồi binary đang chạy sau khi tiến trình/container bị thay thế nếu không có durable storage; trong trường hợp đó, trả trạng thái failed/retry rõ thay vì mất job im lặng. Không thêm database lớn nếu không cần.

**Project structure mở rộng:** `backend/paragraphs.py` xử lý line/block grouping; `backend/page_policy.py` quyết định section/geometry; `backend/quality.py` kiểm tra DOCX/text/render và tạo báo cáo; `backend/jobs.py`/`backend/app.py` giữ API/job guard; `backend/converter.py` điều phối các module hiện có. `tests/fixtures/` giữ fixture synthetic; tệp thật người dùng chỉ lưu ngoài Git. `docs/bao-cao-kiem-thu.md` ghi bằng chứng và phần chưa hoàn tất.

Giữ nguyên nhận diện và giao diện Swiss SaaS hiện tại. Chỉ thêm lựa chọn page policy/label và báo cáo chất lượng nếu cần để người dùng hiểu khác biệt giữa giữ PDF và chuẩn hành chính; không đại tu giao diện. Các thay đổi production cần build, regression, render cả hai mode, Git checkpoint; publish chỉ sau thao tác confirmation riêng.
