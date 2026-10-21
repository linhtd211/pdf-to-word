# Hướng dẫn triển khai

## Container (đề xuất)

PDF processing và OCR dùng thư viện Python/native binary, bộ nhớ và thư mục tạm; vì vậy chọn một runtime container có backend thường trực thay vì đặt toàn bộ xử lý trong một hàm serverless ngắn hạn.

```bash
docker build -t pdf-to-word-pro .
docker run --rm -p 3000:3000 -e PORT=3000 pdf-to-word-pro
```

Ứng dụng bind `0.0.0.0:$PORT` và có health endpoint `GET /api/health`. Khi deploy, cấu hình health path `/api/health`; cài cả Tesseract và language data `vie`, `eng`. Dockerfile hiện đã thực hiện các bước đó. Người vận hành cần chọn giới hạn CPU/RAM đủ cho PDF/OCR và giới hạn thời gian phù hợp; không mở public endpoint upload mà không giữ giới hạn file/trang.

## Vercel/serverless

Không nên chạy PDF/OCR nặng trong Vercel Functions: giới hạn thời gian thực thi, RAM, dung lượng tệp/thư mục tạm, native dependency và cold starts có thể làm job dài thất bại. Nếu muốn dùng Vercel, chỉ triển khai frontend tĩnh ở đó; đặt API chuyển đổi trên một backend container riêng, bật CORS chỉ cho origin đã chọn, và dùng storage tạm/queue thích hợp nếu cần chạy nhiều instance. API hiện lưu trạng thái job trong RAM và file trong filesystem cục bộ, nên không phù hợp scale ngang hay restart giữa job. Không bật lưu trữ lâu dài trừ khi có yêu cầu và chính sách retention rõ ràng.

## An toàn dữ liệu và giới hạn bản đầu

File tối đa 30 MiB, 100 trang; không có đăng nhập; không dùng DB; tạo thư mục tạm ngẫu nhiên trên server; kết quả bị xóa sau download hoặc hết hạn 20 phút. Registry job chỉ tồn tại trong RAM. Trên container, hệ thống tệp là ephemeral; đây là chủ ý của bản đầu, nhưng không phải cam kết bảo mật tuyệt đối/tuân thủ pháp lý. Có thể rà soát TLS, quyền ghi, quota và chính sách xóa của nhà cung cấp trước khi dùng với tài liệu nhạy cảm.
