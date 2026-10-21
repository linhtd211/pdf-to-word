from __future__ import annotations

from pathlib import Path
from xml.sax.saxutils import escape

from PIL import Image, ImageDraw, ImageFont
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

ROOT = Path(__file__).parent / "fixtures"
ROOT.mkdir(parents=True, exist_ok=True)
FONT_CANDIDATES = [Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"), Path("C:/Windows/Fonts/arial.ttf"), Path("C:/Windows/Fonts/calibri.ttf")]
FONT_PATH = next((path for path in FONT_CANDIDATES if path.exists()), None)
if FONT_PATH is None:
    raise RuntimeError("Cần font hỗ trợ tiếng Việt (DejaVu Sans hoặc Arial) để tạo bộ PDF mẫu.")
pdfmetrics.registerFont(TTFont("Vietnamese", str(FONT_PATH)))
styles = getSampleStyleSheet()
body = ParagraphStyle("VietnameseBody", parent=styles["BodyText"], fontName="Vietnamese", fontSize=11, leading=17, spaceAfter=8)
title = ParagraphStyle("VietnameseTitle", parent=styles["Title"], fontName="Vietnamese", fontSize=16, leading=22, alignment=TA_CENTER, textColor=colors.HexColor("#1e3a8a"), spaceAfter=14)
small = ParagraphStyle("VietnameseSmall", parent=body, fontSize=9, leading=13)


def build(name: str, flowables: list, margins=(48, 48, 50, 50)) -> None:
    left, right, top, bottom = margins
    doc = SimpleDocTemplate(str(ROOT / name), pagesize=A4, leftMargin=left, rightMargin=right, topMargin=top, bottomMargin=bottom)
    doc.build(flowables)


# 1) PDF văn bản điện tử
build("01-van-ban-dien-tu.pdf", [
    Paragraph("HƯỚNG DẪN SỬ DỤNG HỆ THỐNG", title),
    Paragraph("Tài liệu này được tạo bằng văn bản điện tử có thể trích xuất trực tiếp.", body),
    Paragraph("Mục tiêu", ParagraphStyle("H", parent=body, fontSize=13, leading=18, textColor=colors.HexColor("#1d4ed8"))),
    Paragraph("Chuyển nội dung PDF sang tài liệu Word có thể chỉnh sửa, giữ các đoạn văn và dấu tiếng Việt.", body),
    Paragraph("Nội dung thử nghiệm có chữ đậm, chữ nghiêng và các câu tiếng Việt: Cộng hòa xã hội chủ nghĩa Việt Nam; Độc lập - Tự do - Hạnh phúc.", body),
    Paragraph("Kết thúc tài liệu mẫu.", small),
])

# 2) Công văn hành chính có khối song song
left_block = Paragraph("<b>SỞ NỘI VỤ</b><br/>Số: 15/CV-SNV<br/>V/v hướng dẫn lập hồ sơ", small)
right_block = Paragraph("<b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</b><br/><b>Độc lập - Tự do - Hạnh phúc</b><br/><br/>Hà Nội, ngày 15 tháng 10 năm 2026", ParagraphStyle("Center", parent=small, alignment=TA_CENTER))
letter_header = Table([[left_block, right_block]], colWidths=[235, 270])
letter_header.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP"), ("ALIGN", (1,0), (1,0), "CENTER"), ("LEFTPADDING", (0,0), (-1,-1), 2), ("RIGHTPADDING", (0,0), (-1,-1), 2), ("LINEBELOW", (0,0), (-1,-1), .3, colors.HexColor("#e2e8f0"))]))
build("02-cong-van-hanh-chinh.pdf", [letter_header, Spacer(1, 19), Paragraph("Kính gửi: Các đơn vị trực thuộc", body), Paragraph("Căn cứ quy định về công tác văn thư, Sở Nội vụ đề nghị các đơn vị thực hiện việc lập hồ sơ công việc đầy đủ, bảo đảm đúng trình tự và thời hạn.", body), Paragraph("Nơi nhận: Các đơn vị; Lưu: VT, HCTH.", small), Spacer(1, 15), Paragraph("GIÁM ĐỐC<br/><br/><br/>Nguyễn Văn An", ParagraphStyle("Sign", parent=small, alignment=TA_CENTER))])

# 3) Bảng có header và số liệu
rows = [["STT", "Tên đơn vị", "Số hồ sơ", "Ghi chú"], ["1", "Phòng Hành chính", "12", "Đã hoàn thành"], ["2", "Phòng Tổ chức", "8", "Đang rà soát"], ["3", "Trung tâm Lưu trữ", "16", "Bổ sung phụ lục"]]
table = Table([[Paragraph(escape(cell), small) for cell in row] for row in rows], colWidths=[42, 172, 72, 180])
table.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor("#dbeafe")), ("GRID", (0,0), (-1,-1), .6, colors.HexColor("#91a4bd")), ("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("LEFTPADDING", (0,0), (-1,-1), 7), ("TOPPADDING", (0,0), (-1,-1), 8), ("BOTTOMPADDING", (0,0), (-1,-1), 8)]))
build("03-bang-bieu.pdf", [Paragraph("BẢNG TỔNG HỢP HỒ SƠ", title), Paragraph("Số liệu kiểm thử có cấu trúc bảng để xác minh bảng Word có thể chỉnh sửa.", body), table])

# 4) Tài liệu nhiều trang
multi: list = [Paragraph("BÁO CÁO TỔNG HỢP", title)]
for page in range(1, 5):
    multi.extend([Paragraph(f"Phần {page}: Nội dung báo cáo", ParagraphStyle(f"Part{page}", parent=body, fontSize=13, textColor=colors.HexColor("#1d4ed8"))), Paragraph((f"Đây là đoạn văn bản thử nghiệm của trang {page}. " * 18), body), Paragraph("Các chỉ tiêu được ghi nhận theo biểu mẫu và cần được lưu trữ đầy đủ.", body)])
    if page < 4:
        multi.append(PageBreak())
build("04-nhieu-trang.pdf", multi)

# 5) PDF có hình ảnh bitmap cùng văn bản
img = Image.new("RGB", (680, 250), "white")
draw = ImageDraw.Draw(img)
font = ImageFont.truetype(str(FONT_PATH), 38)
draw.rounded_rectangle((10, 10, 670, 240), radius=24, fill="#eaf2ff", outline="#4f83dc", width=4)
draw.text((45, 55), "SƠ ĐỒ QUY TRÌNH", font=font, fill="#1e3a8a")
draw.rounded_rectangle((50, 145, 250, 205), radius=12, fill="#ffffff", outline="#3b82f6", width=3)
draw.text((79, 158), "Tiếp nhận", font=ImageFont.truetype(str(FONT_PATH), 23), fill="#334155")
draw.line((260, 175, 405, 175), fill="#2563eb", width=6)
draw.polygon([(405, 175), (383, 160), (383, 190)], fill="#2563eb")
draw.rounded_rectangle((420, 145, 630, 205), radius=12, fill="#ffffff", outline="#3b82f6", width=3)
draw.text((456, 158), "Lưu trữ", font=ImageFont.truetype(str(FONT_PATH), 23), fill="#334155")
image_path = ROOT / "workflow.png"
img.save(image_path)
flow_image = Image.open(image_path)
image_canvas = canvas.Canvas(str(ROOT / "05-co-hinh-anh.pdf"), pagesize=A4)
image_canvas.setFont("Vietnamese", 15)
image_canvas.drawString(52, 785, "Mô tả quy trình xử lý hồ sơ")
image_canvas.drawImage(str(image_path), 55, 490, width=500, height=184, preserveAspectRatio=True, mask="auto")
image_canvas.setFont("Vietnamese", 11)
image_canvas.drawString(52, 455, "Hình minh họa cần được giữ lại như một đối tượng ảnh riêng trong Word.")
image_canvas.save()

# 6) Scan ảnh tiếng Việt được đóng gói thành một trang PDF
scan = Image.new("RGB", (1700, 2300), "white")
scan_draw = ImageDraw.Draw(scan)
font_title = ImageFont.truetype(str(FONT_PATH), 53)
font_body = ImageFont.truetype(str(FONT_PATH), 40)
font_small = ImageFont.truetype(str(FONT_PATH), 34)
scan_draw.text((150, 140), "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", font=font_title, fill="#111111")
scan_draw.text((470, 225), "Độc lập - Tự do - Hạnh phúc", font=font_body, fill="#111111")
scan_draw.line((570, 290, 1120, 290), fill="#111111", width=3)
scan_draw.text((550, 390), "THÔNG BÁO", font=font_title, fill="#111111")
lines = [
    "Kính gửi: Các đơn vị trực thuộc,",
    "Căn cứ kế hoạch công tác năm 2026, đề nghị các đơn vị",
    "rà soát hồ sơ và gửi báo cáo trước ngày 30 tháng 10.",
    "Nội dung cần bảo đảm đầy đủ, chính xác và đúng thời hạn.",
    "Nơi nhận: Như trên; Lưu: Văn thư.",
]
y = 560
for line in lines:
    scan_draw.text((170, y), line, font=font_body, fill="#111111")
    y += 145
scan_draw.text((920, 1540), "THỦ TRƯỞNG ĐƠN VỊ", font=font_body, fill="#111111")
scan_draw.text((1010, 1920), "Nguyễn Văn Bình", font=font_small, fill="#111111")
scan_path = ROOT / "scan-page.png"
scan.save(scan_path, dpi=(220, 220))
scan_canvas = canvas.Canvas(str(ROOT / "06-scan-tieng-viet.pdf"), pagesize=A4)
scan_canvas.drawImage(str(scan_path), 0, 0, width=A4[0], height=A4[1], preserveAspectRatio=False)
scan_canvas.save()
print(f"Đã tạo 6 PDF kiểm thử tại {ROOT}")
