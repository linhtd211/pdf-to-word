# Tham chiếu thể thức văn bản hành chính — Nghị định 30/2020/NĐ-CP

**Nguồn chính thức:** [Nghị định số 30/2020/NĐ-CP của Chính phủ — Cổng văn bản Chính phủ](https://vanban.chinhphu.vn/default.aspx?pageid=27160&docid=199378), ban hành ngày 05/03/2020. Trang chính thức dẫn tới bản ký số PDF: [30.signed.pdf](https://datafiles.chinhphu.vn/cpp/files/vbpq/2020/03/30.signed.pdf).

Tài liệu người dùng gửi `Số215.doc` là Phụ lục I của Nghị định này. Các thông số dưới đây được đối chiếu với Phần I, các mục I, II.1 và IV–V trong phụ lục đính kèm.

## Quy tắc áp dụng trong converter

- Với PDF điện tử mà trang đầu nhận diện được đủ hai dòng **“CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM”** và **“Độc lập - Tự do - Hạnh phúc”**, tạo **một bảng Word thật, 2 cột cân bằng, không viền**. Cột trái giữ tên cơ quan ban hành; cột phải đặt Quốc hiệu và Tiêu ngữ. Hàng kế tiếp căn số/ký hiệu bên trái với địa danh, ngày tháng bên phải; trích yếu kiểu công văn nằm dưới số/ký hiệu. Nội dung header gốc được loại khỏi luồng paragraph để không bị lặp.
- Quốc hiệu dùng Times New Roman 12 pt, chữ hoa, đậm; converter scale ngang dòng này còn 90% để vừa một dòng trong bảng 2 cột cân bằng, không thay đổi cỡ chữ point. Đây là quyết định typography của sản phẩm để chống wrap, không phải thông số được Nghị định quy định riêng; cần kiểm tra trong Microsoft Word. Tiêu ngữ Times New Roman 13 pt, đậm, căn giữa, có đường kẻ dưới. Tên cơ quan đặt ở cột trái, tên cơ quan ban hành đậm và có đường kẻ dưới. Ngày tháng ở cột phải dùng chữ nghiêng.
- Khi nhận diện văn bản hành chính, thiết lập trang A4 dọc (A4 ngang nếu PDF nguồn là ngang), lề trên 20 mm, dưới 25 mm, trái 30 mm, phải 15 mm; font văn bản Times New Roman 13 pt, màu đen, khoảng cách đoạn 6 pt và giãn dòng đơn. Đây là các giá trị được chọn bên trong khoảng quy định tại Phụ lục I.
- Áp dụng quy tắc này cho cả lựa chọn **Ưu tiên bố cục** và **Ưu tiên chỉnh sửa** nếu phát hiện header hành chính; nhánh bố cục tổng quát `pdf2docx` vẫn dùng cho PDF khác. API trả cảnh báo để người dùng rà lại tên cơ quan, số/ký hiệu và ngày tháng được trích.

## Phạm vi và giới hạn

Nhận diện tự động dựa trên text PDF có thể trích xuất được; không xác định header cho trang scan không có lớp text, header bị OCR sai, PDF có bố cục lạ hoặc nội dung thiếu một trong hai cụm từ. Trong các trường hợp đó, converter giữ nhánh xử lý hiện có và không tự chèn dữ liệu Quốc hiệu/Tiêu ngữ. Trên PDF BV/TYT được cung cấp, converter cũng dựng footer “Nơi nhận”/khối ký và hai bảng phụ lục có số cột, width cột theo nguồn; đây là quy tắc đã kiểm chứng cho form cụ thể, không chứng minh hỗ trợ đầy đủ ô 5a/5b, căn cứ, header/footer/số trang ở mọi loại công văn. Tệp DOCX đầu ra cần được rà soát, nhất là trước khi ban hành văn bản chính thức.
