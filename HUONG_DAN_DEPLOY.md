# Thử nghiệm triển khai GitHub + Render + Vercel

## 1. GitHub
Đưa **nội dung bên trong thư mục dự án** lên một repository GitHub mới. Không upload ZIP nguyên khối.

## 2. Render (backend)
Render → New → Blueprint → chọn repo GitHub → xác nhận `render.yaml`. Dùng gói **Starter hoặc cao hơn** phù hợp tài nguyên. Chờ build Docker thành công. Kiểm tra `https://<ten-dich-vu>.onrender.com/api/health` trả về JSON có `ok: true`.

## 3. Vercel (frontend)
Vercel → Add New → Project → Import cùng repository. Root Directory để trống; Framework Preset Vite; Build Command `pnpm build`; Output Directory `dist`. Trong Environment Variables thêm `VITE_API_BASE_URL` = **URL Render thực tế**, ví dụ `https://pdfword-pro-api.onrender.com` (không có dấu / cuối). Deploy.

## 4. Cho phép CORS
Quay lại Render → Environment → `FRONTEND_ORIGINS` = **URL Vercel thực tế**, ví dụ `https://pdfword-pro.vercel.app` (không có / cuối). Save/Deploy lại backend. Nếu có domain riêng, phân cách nhiều origin bằng dấu phẩy.

## 5. Kiểm thử
- Truy cập `/api/health` của Render, xác nhận trả về JSON.
- Mở web Vercel, upload PDF nhỏ có text, đợi chuyển đổi, tải DOCX.
- Thử PDF bảng biểu và PDF scan tiếng Việt; kiểm tra kết quả Word thủ công.
- Nếu upload bị CORS, so sánh URL Vercel thực tế với `FRONTEND_ORIGINS` trên Render.
- Nếu gặp 503 hoặc quá hạn: kiểm tra logs Render, tài nguyên RAM, giới hạn 30 MiB/100 trang, timeout 5 phút.

## Giới hạn của bản thử nghiệm
- Chưa có cơ sở dữ liệu/hàng đợi bền vững. Chạy **một instance Render**; khi redeploy hoặc restart, job đang xử lý có thể mất.
- Không phù hợp tài liệu nhạy cảm khi chưa bổ sung xác thực, chính sách lưu trữ và kiểm soát truy cập.
- Chưa kiểm chứng build/deploy thực tế trên tài khoản Render/Vercel của người dùng.
- Vercel chỉ host giao diện; việc chuyển PDF → Word chạy tại Render.
