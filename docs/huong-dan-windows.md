# Hướng dẫn chạy trên Windows

## Yêu cầu

- Windows 10/11 64-bit.
- Node.js 22 LTS và pnpm 11. `corepack enable` có thể bật pnpm đi cùng Node; dự án pin pnpm `11.25.0` trong `package.json`.
- Python 3.12 64-bit.
- Tesseract OCR cài bản Windows và language data `vie` + `eng`; thư mục cài đặt Tesseract cần có trong `PATH` (thường `C:\Program Files\Tesseract-OCR`). Có thể đặt `TESSDATA_PREFIX` nếu traineddata không nằm trong thư mục mặc định.

## Cài đặt và khởi chạy

Mở PowerShell tại thư mục dự án:

```powershell
corepack enable
pnpm install
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
pnpm build
python -m uvicorn backend.app:app --host 0.0.0.0 --port 3000
```

Nếu chính sách PowerShell không cho kích hoạt virtualenv, dùng `Set-ExecutionPolicy -Scope Process Bypass` trong cửa sổ PowerShell hiện tại hoặc gọi trực tiếp `.venv\Scripts\python.exe -m uvicorn ...`.

Mở `http://localhost:3000`. Tắt máy chủ bằng `Ctrl+C`. Nếu thay đổi mã frontend, chạy lại `pnpm build`; mã Python được nạp lại khi khởi động lại Uvicorn.

## Tạo tài liệu thử nghiệm và chạy test

Cài thêm test tools: `pip install -r requirements-dev.txt`, sau đó chạy `python tests/generate_fixtures.py` và `python -m pytest -q`. Generator cần font có hỗ trợ tiếng Việt (Arial hoặc font tương đương); test kiểm OCR cần Tesseract `vie`/`eng`.

## Khắc phục sự cố

- Lỗi “chưa cài dữ liệu OCR tiếng Việt”: kiểm tra `tesseract --list-langs` có `vie` và `eng`; đặt đúng `TESSDATA_PREFIX`, rồi khởi động lại app.
- Cổng 3000 đang được dùng: dừng tiến trình đó trước; khi chạy trên môi trường Webdev, giữ cổng cấu hình do Preview yêu cầu.
- PDF bị từ chối: giới hạn phiên bản đầu là 30 MiB và 100 trang; PDF có mật khẩu phải mở khóa trước.
