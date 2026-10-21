# Báo cáo sửa lỗi bước 1

- Loại bỏ hard line breaks theo từng dòng PDF trong luồng chuyển đổi native, kể cả chế độ layout; các dòng thuộc một đoạn được nối bằng khoảng trắng.
- Nội dung trong ô bảng cũng được chuẩn hóa dòng để Word tự dàn lại.
- Cho phép chạy luồng native khi thư viện pdf2docx chưa cài; khi pdf2docx có sẵn, đường chuyển đổi layout/source không phải hành chính vẫn dùng thư viện này.
- Bổ sung hai regression tests về ngắt dòng đoạn và ô bảng.
- Kiểm thử trong môi trường hiện tại: 25/25 passed.
- Chuyển đổi thử 3276.QĐ-BYT.pdf ở chế độ editable/administrative thành công.

## Chưa hoàn thiện
- Chưa nghiệm thu trực quan từng trang trên Microsoft Word; chưa xác nhận bảng/phụ lục giữ nguyên toàn bộ bố cục.
- Thuật toán ghép đoạn vẫn là heuristic; PDF nhiều cột, bảng scan và bố cục phức tạp cần thử thêm.
- Đường chuyển đổi pdf2docx dành cho layout/source có thể vẫn giữ ngắt dòng theo thiết kế của thư viện.
