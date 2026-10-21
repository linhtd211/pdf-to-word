# TODO — PDF to Word Pro

**Phạm vi nguồn:** yêu cầu trong `Pasted_content.txt`; bổ sung chuẩn văn bản hành chính theo Nghị định 30/2020/NĐ-CP và đối chiếu mẫu PDF BV/TYT người dùng cung cấp. Trạng thái `[x]` chỉ áp dụng cho hành vi có bằng chứng; `[ ]` là phần còn hạn chế/chưa đủ kiểm chứng.

## 1. Giao diện và trải nghiệm tiếng Việt
- [x] Giao diện hoàn toàn bằng tiếng Việt, phong cách SaaS hiện đại, tối giản, chuyên nghiệp; màu xanh dương/trắng/xám nhạt, góc bo tinh tế, Lucide Icons, chuyển động nhẹ, responsive cho desktop/tablet/mobile; không thêm chức năng thừa, ưu tiên thao tác đơn giản và nhanh.
- [x] Thanh điều hướng gồm logo **PDF to Word Pro**, “Chuyển đổi PDF”, “Giới thiệu”. Tiêu đề chính là “Chuyển PDF sang Word chính xác”; mô tả là “Giữ nguyên bố cục, bảng biểu và định dạng. Chuyển đổi tài liệu thành Word có thể chỉnh sửa.”
- [x] Màn hình tải có vùng kéo-thả/chọn file nổi bật; sau khi tải hiện preview PDF, thông tin tài liệu và lựa chọn mode; màn hình kết quả có thông báo thành công, tên/kích thước DOCX, nút tải và chuyển tài liệu khác.
- [ ] Preview/so sánh Word cạnh PDF trong ứng dụng nếu khả thi; hiện app có preview PDF, chưa có preview DOCX side-by-side.

## 2. Tải file, xem trước và xử lý lỗi
- [x] Hỗ trợ kéo-thả và chọn PDF; hiển thị tên, dung lượng, tổng số trang; xem trước bằng PDF.js; kiểm tra phần mở rộng, PDF magic bytes, khả năng mở và giới hạn trang.
- [x] Báo lỗi rõ khi file hỏng, sai định dạng, vượt giới hạn hoặc không hỗ trợ; giữ giới hạn hiện tại 30 MiB và 100 trang/job.

## 3. Chuyển đổi PDF điện tử và scan
- [x] Phân tích PDF trước khi xử lý, phân biệt text điện tử với trang scan/mixed; không chạy OCR toàn tài liệu nếu có thể trích xuất text điện tử.
- [ ] PDF điện tử phải giữ dấu tiếng Việt, nội dung, font/cỡ chữ, đậm/nghiêng/gạch chân, căn trái/phải/giữa/đều, khoảng cách dòng/đoạn, thụt đầu dòng, lề, ngắt trang, tiêu đề/đoạn, bảng, ảnh và vị trí tương đối. **Đã kiểm tra text, bảng, trang, ảnh và alignment trên fixture/nguồn BV-TYT; định dạng font/spacing và bố cục PDF tổng quát vẫn có thể thay đổi.**
- [ ] PDF scan phải OCR tiếng Việt có dấu, phân tích vùng văn bản, bảng, ảnh và tái tạo Word chỉnh sửa được; không đưa nguyên trang scan thành ảnh. **Tesseract `vie`/`eng` tạo paragraph editable trên scan rõ; chưa có bộ nhận dạng lưới bảng trong ảnh và chất lượng scan mờ/nghiêng chưa được xác minh.**

## 4. Cấu trúc Word và văn bản hành chính
- [x] Ưu tiên Paragraph, Run, Table, Cell, Section, Header/Footer chuẩn Word; không tạo TextBox cho từng dòng/toàn bộ văn bản và không dùng quá nhiều đối tượng nổi. **Các fixture/đầu ra nguồn được kiểm tra không chứa `w:txbxContent`; ảnh dùng floating anchor để giữ tọa độ PDF.**
- [ ] Tối ưu công văn, quyết định, kế hoạch, báo cáo, thông báo, tờ trình, biểu mẫu và văn bản y tế; nhận diện đúng các phần Quốc hiệu–Tiêu ngữ, cơ quan, số/ký hiệu, địa danh/ngày, trích yếu, căn cứ, nội dung, nơi nhận, chữ ký/dấu, bảng/phụ lục; không đảo thứ tự hoặc tự ý sửa nội dung. **Đã dựng header/footer cho mẫu BV/TYT và so khớp toàn bộ từ; các cấu trúc hành chính chưa có fixture riêng vẫn cần rà soát.**
- [x] Với PDF điện tử nhận diện được Quốc hiệu và Tiêu ngữ, tạo bảng header Word thật 2 cột cân bằng, tránh lặp header trong body; số/ngày/trích yếu được gom vào đúng khối. Áp dụng font/căn chỉnh theo tham chiếu Phụ lục I Nghị định 30/2020/NĐ-CP; Quốc hiệu Times New Roman 12 pt được scale ngang 90% để tránh wrap (thỏa hiệp của ứng dụng, không phải thông số Nghị định quy định riêng).

## 5. Bảng biểu và hình ảnh
- [ ] Phát hiện bảng, số hàng/cột, ô gộp, viền, căn chỉnh, kích thước tương đối; đầu ra phải là bảng Word chỉnh sửa được. **Bảng vector/text được dựng thành Table Word và giữ width từng cột; merge-cell detection tổng quát chưa đầy đủ.**
- [x] Trong PDF BV/TYT gốc: trang phụ lục có đúng hai bảng Word thật lần lượt 4×10 và 4×11 cột, giữ tỷ lệ độ rộng cột không đều, đúng thứ tự, trên trang A4 ngang thứ hai; text nằm trong vùng bảng không bị thêm lại thành paragraph.
- [x] Giữ ảnh/signature/stamp trong suốt, kích thước và vị trí theo bbox PDF; hỗ trợ ảnh chồng lấp bằng drawing anchors. **Bốn ảnh của PDF BV/TYT giữ được soft-mask và tọa độ gần như chính xác (sai khác đo được khoảng 0,1 pt trong LibreOffice), không tạo trang phụ do ảnh phóng lớn.**

## 6. Khổ giấy, bố cục và hai chế độ
- [x] Giữ trang A4/A5/kích thước khác, dọc/ngang, lề, vị trí và ngắt trang theo section PDF; khi xung đột, ưu tiên Word sửa tự nhiên nhưng giảm tối đa thay đổi hình thức.
- [x] Hai chế độ phải khác nhau thật sự: **Ưu tiên chỉnh sửa** gộp các dòng hiển thị thuộc cùng đoạn thành đoạn Word tự nhiên, không chèn ngắt dòng theo từng dòng PDF và không thu hẹp chiều rộng đoạn theo dòng PDF, để Word tự xuống dòng khi người dùng sửa/bổ sung; **Ưu tiên bố cục** giữ ngắt dòng/vị trí gần PDF nhất có thể. Kiểm tra riêng cả hai chế độ. **Đã kiểm tra Quyết định 3276: editable có 0 hard-break trong DOCX, layout giữ 134 hard-break chính; cả hai render 7 trang và đủ 2.753 token.**
- [x] Với PDF BV/TYT gốc, cả hai mode xuất thành đúng hai trang: trang đầu A4 dọc, trang phụ lục A4 ngang; LibreOffice mở được và render đúng hai trang.

## 7. Tiến trình và kết quả chuyển đổi
- [x] Tiến trình phản ánh bước backend thực tế: đọc cấu trúc PDF, phân tích bố cục, nhận diện văn bản/bảng biểu, tạo Word, hoàn tất; không tăng phần trăm bằng timer giả và không treo UI trong lúc job chạy.
- [x] Màn hình kết quả hiển thị trạng thái, tên/kích thước Word, tải xuống và chuyển file khác; khi lỗi có thông báo rõ.
- [ ] Preview trực tiếp DOCX hoặc so sánh trực quan hai tài liệu ngay trên app là tùy chọn “nếu có thể”, chưa triển khai.

## 8. Bảo mật, hiệu năng và kiến trúc
- [x] Phiên bản đầu không yêu cầu đăng nhập/không dùng database; file nằm trong thư mục tạm, tự xóa khi tải xuống/lỗi và có TTL job; giới hạn 30 MiB/100 trang; xác thực PDF; lỗi rõ ràng.
- [x] Frontend React + TypeScript + Tailwind + Lucide + PDF.js; backend Python/FastAPI; xử lý PyMuPDF, pdfplumber, pdf2docx, python-docx và Tesseract `vie`/`eng`; frontend static được FastAPI phục vụ trên cổng 3000.
- [x] Không yêu cầu Vercel; hướng dẫn deploy dùng container vì xử lý PDF/OCR nặng. Nếu triển khai serverless phải nêu rõ giới hạn/thay backend.

## 9. Kiểm thử, tài liệu và bàn giao
- [x] Có PDF kiểm thử đại diện: văn bản điện tử, công văn, bảng, nhiều trang, ảnh, scan tiếng Việt; kiểm tra nội dung/dấu, bảng editable, TextBox, thứ tự đoạn, DOCX mở/render và thay đổi bố cục.
- [x] Có regression synthetic cho header/body/footer, ảnh alpha chồng, lề/căn chỉnh, section A4 dọc/ngang, bảng 10/11 cột, khổ ngang tùy chỉnh, cỡ font bảng theo nguồn, thụt đoạn theo toàn khối, whitespace spans và natural flow; lần chạy gần nhất `pytest` **17 passed**, Pyright **0 lỗi/0 cảnh báo**, Python `compileall`, TypeScript typecheck/build.
- [ ] Microsoft Word desktop chưa có trong môi trường; tính tương thích được xác minh bằng `python-docx` và LibreOffice, không tuyên bố đã kiểm tra trên Microsoft Word.
- [x] Bàn giao Web App Preview, source frontend/backend, hướng dẫn Windows, hướng dẫn triển khai, fixtures và báo cáo chức năng/hạn chế; không đưa PDF/DOCX thực của người dùng vào kho mã.

## 10. Đối chiếu cụ thể PDF BV/TYT người dùng gửi
- [x] Nguồn: PDF 2 trang; trang 1 A4 dọc có header hai khối, field số/ngày nằm rời, trích yếu 2 dòng, body, nơi nhận, chữ ký/dấu; trang 2 A4 ngang có hai phụ lục.
- [x] Header sau chuyển đổi: ghép số/ký hiệu và ngày/tháng/năm vào đúng ô; giữ cả hai dòng trích yếu trong khối header; không lặp các field trong body.
- [x] Footer sau chuyển đổi: “Nơi nhận” bên trái, chức danh/chữ ký/tên người ký bên phải; giữ lớp dấu đỏ/chữ ký xanh chồng nhau.
- [x] Text kiểm tra theo từ đã chuẩn hóa Unicode: PDF nguồn và bản render mới đều có **449 token**, không thiếu hoặc thừa token; bbox bốn ảnh trên trang 1 khớp nguồn trong khoảng 0,1 pt.
- [x] Cấu trúc/renders: DOCX có bảng header 2×2, footer 1×2 và bảng phụ lục 4×10/4×11 với độ rộng cột theo nguồn; LibreOffice render đúng 2 trang và đúng hướng.
- [ ] Khác biệt còn có thể có do font Word cài đặt trên máy người dùng, line-wrap và cách Word/LibreOffice render; cần rà trực tiếp trong Microsoft Word trước khi ban hành văn bản chính thức.

## 11. Nguyên tắc giữ phạm vi
- [x] Ưu tiên chất lượng chuyển đổi; không thêm chỉnh sửa/gộp/tách PDF hoặc định dạng ngoài DOCX; không mô phỏng chuyển đổi; không chèn ảnh nguyên trang PDF để thay cho tài liệu chỉnh sửa được; không dùng TextBox cho toàn bộ nội dung.

## 12. Hồi quy PDF Quyết định 3276
- [x] Giữ chính xác kích thước trang PDF không chuẩn cho các trang phụ lục ngang/dọc; cả hai chế độ chuyển đổi từ PDF 7 trang đều render thành 7 trang bằng LibreOffice, không còn khoảng trắng/overflow tạo thêm trang.
- [x] Bảo toàn nội dung text: so sánh PDF nguồn với PDF render DOCX sau chuẩn hóa Unicode cho kết quả **2.753 token/2.753 token**, không thiếu/thừa; người ký vẫn ở trang 2.
- [x] Bảng Word editable của header và phụ lục: header 2×2; các bảng 7×5, 6×5, 7×5, 8×5 và 10×4. Cỡ chữ bảng, chiều rộng lưới, thụt đoạn và khoảng trắng text spans được suy ra từ nguồn; PDF/DOCX thật không được thêm vào kho mã.

## 13. Yêu cầu hoàn thiện bổ sung (18/10/2026)

### 13.1 Semantic paragraph reconstruction — ưu tiên cao nhất
- [ ] Các dòng trong PDF thuộc cùng một đoạn không được chuyển thành những paragraph riêng. Tự động ghép các dòng PDF thuộc cùng một đoạn thành một paragraph Word; phân biệt xuống dòng do chiều rộng trang với kết thúc đoạn thực sự; khi chỉnh sửa Điều 1, Điều 2, Điều 3, các dòng phải tự động dồn và giãn như văn bản Word thông thường.
- [ ] Không sử dụng tọa độ tuyệt đối để cố định từng dòng văn bản thông thường. Dùng paragraph, căn lề, thụt đầu dòng, giãn dòng và khoảng cách đoạn của Word; giữ định dạng chữ đậm, nghiêng, gạch chân và định dạng đặc biệt.
- [ ] Không tạo TextBox cho dòng/đoạn văn thông thường; không ghép nhầm các đoạn độc lập, tiêu đề hoặc nội dung giữa các ô bảng. Nghiệm thu: thêm/xóa chữ tự động reflow, không ép dòng theo tọa độ PDF và có thể căn đều hai lề đúng cách.

### 13.2 Page setup, lề, section và pagination theo lựa chọn người dùng
- [ ] Có chế độ **Giữ nguyên bố cục gốc**: giữ kích thước giấy, hướng trang và lề thực tế của PDF khi xác định được; không tự ý tạo lề bất thường để dịch chuyển chữ; không tạo section không cần thiết; vẫn ưu tiên cấu trúc Word tự nhiên.
- [ ] Có chế độ **Chuẩn thể thức văn bản hành chính Việt Nam**: dùng khổ A4 21 × 29,7 cm; lề trên 2–2,5 cm, lề dưới 2–2,5 cm, lề trái 3–3,5 cm, lề phải 1,5–2 cm; Times New Roman Unicode màu đen cho phần văn bản theo quy định; định dạng thành phần thể thức hành chính phù hợp. Không coi đúng lề là đã đạt toàn bộ thể thức.
- [ ] Trang có bảng rộng được phép dùng A4 ngang khi cần nhưng không được cắt nội dung; hướng và lề phải hợp lý cho bảng rộng.
- [ ] Chỉ dùng section khi kích thước giấy, hướng giấy hoặc cấu trúc trang thực sự cần thiết; không tạo section mới chỉ để giữ tọa độ dòng chữ; không dùng lề trên/dưới quá lớn để tạo khoảng trắng; không ngắt trang cứng tùy tiện giữa đoạn.
- [ ] Nội dung phải tự động chảy qua trang tiếp theo khi người dùng chỉnh sửa; không sinh trang trắng ngoài ý muốn. Kiểm tra trang 7 của PDF 3276 để không lặp lại lề dưới bất thường khoảng 17,33 cm hoặc section không đồng nhất.

### 13.3 Bảng Word thực sự và nội dung từng ô
- [ ] Tạo bảng Word thật, không chuyển toàn bộ bảng thành ảnh; giữ hàng, cột, ô gộp, đường viền, màu nền, màu chữ, chữ đậm và căn chỉnh.
- [ ] Nội dung mỗi ô tự động xuống dòng theo chiều rộng ô; không tạo paragraph riêng cho mỗi dòng PDF bên trong một ô; không làm mất/chuyển nội dung sang ô khác.
- [ ] Bảng phải chỉnh sửa được và không vỡ khi kéo dài qua nhiều trang. Ưu tiên cấu trúc chỉnh sửa hơn định vị tuyệt đối.

### 13.4 Nhận diện cấu trúc hành chính và ảnh ký/dấu
- [ ] Nhận diện đúng: tên cơ quan ban hành; Quốc hiệu, Tiêu ngữ; số/ký hiệu; địa danh/ngày tháng năm; tên loại/trích yếu; nội dung/điều khoản/căn cứ pháp lý; nơi nhận; chức vụ/họ tên người ký; hình ảnh chữ ký và con dấu.
- [ ] Không làm mất chữ ký/con dấu; giữ nguyên chúng dưới dạng ảnh, không tự tạo lại hoặc sửa nội dung ảnh.

### 13.5 PDF scan, PDF hỗn hợp và chất lượng OCR
- [ ] Phân biệt PDF điện tử, PDF scan và PDF hỗn hợp. Với PDF điện tử ưu tiên trích xuất text gốc; với PDF scan dùng OCR tiếng Việt; nhận diện vùng văn bản, bảng, hình ảnh và tiêu đề.
- [ ] Hỗ trợ xử lý ảnh nghiêng/độ phân giải thấp ở mức có thể; không OCR lại vùng đã có text chính xác nếu không cần; cảnh báo vùng OCR có confidence thấp.
- [ ] Không tự suy đoán hoặc sửa nội dung pháp lý khi không đủ căn cứ; giữ nguyên text được nhận diện và nêu cảnh báo khi không chắc.

### 13.6 Bộ kiểm tra chất lượng và báo cáo có bằng chứng
- [ ] Kiểm tra nội dung: so sánh PDF với DOCX; phát hiện chữ mất, lặp hoặc sai thứ tự; kiểm tra số lượng bảng, hình ảnh và vùng văn bản.
- [ ] Kiểm tra cấu trúc Word: paragraph bị chia không cần thiết; hard-break bất thường; lề/kích thước giấy/section bất thường; bảng có thể chỉnh sửa.
- [ ] Kiểm tra bố cục: render DOCX thành PDF; so sánh trực quan với PDF nguồn; phát hiện sai khác vị trí, font, bảng, hình ảnh và số trang; phân biệt sai khác do chuẩn hóa thể thức với lỗi chuyển đổi.
- [ ] Báo cáo kết quả kiểm thử, lỗi còn tồn tại và mức độ nghiêm trọng; không tuyên bố hoàn hảo nếu chưa có bằng chứng. Tạo regression tự động cho các lỗi đã phát hiện và chuyển đổi lại từ chính PDF nguồn Quyết định 3276.

### 13.7 Độ bền backend, an toàn và hoàn tất bàn giao
- [ ] Rà soát tác vụ treo và bổ sung timeout xử lý; bảo đảm có giới hạn tài nguyên CPU/RAM và có thông báo lỗi/thử lại rõ ràng.
- [ ] Rà soát xử lý đồng thời nhiều yêu cầu; tránh nhận hàng đợi vô hạn; dọn tệp tạm/orphan và xử lý PDF lỗi an toàn.
- [ ] Rà soát trạng thái job khi máy chủ khởi động lại; không âm thầm làm mất trạng thái: khôi phục trạng thái bền nếu khả năng lưu trữ cho phép, nếu không thì đánh dấu job bị gián đoạn/thất bại rõ ràng. Không thay kiến trúc lớn nếu có thể khắc phục trên hệ thống hiện tại.
- [ ] Sửa trực tiếp thuật toán trong mã nguồn để các tài liệu PDF tương tự được xử lý đúng; không chỉ chỉnh sửa file Word mẫu và không chỉ cập nhật README/TODO/báo cáo.
- [ ] Giữ giao diện và chức năng hiện có ổn định, không xây lại từ đầu hoặc thay đổi giao diện nếu không cần thiết.
- [ ] Bàn giao DOCX mới tạo trực tiếp từ PDF gốc, báo cáo lỗi đã sửa, kết quả kiểm thử và hạn chế còn lại. Tệp Word đã được người dùng chỉnh sửa không được dùng làm nguồn nếu chưa có chỉ dẫn rõ về nội dung chỉnh sửa cần giữ.
