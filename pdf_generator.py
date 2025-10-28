from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from config import TABLE_CONFIG
import os

def generate_pdf(data, file_path, font_dir):
    font_path = os.path.join(font_dir, 'DejaVuSans.ttf')
    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont('DejaVu', font_path))
        font_name = 'DejaVu'
    else:
        font_name = 'Helvetica'

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='TitleCenter', fontName=font_name, fontSize=16, alignment=TA_CENTER, spaceAfter=20))
    styles.add(ParagraphStyle(name='Header', fontName=font_name, fontSize=12, spaceAfter=8))
    styles.add(ParagraphStyle(name='Cell', fontName=font_name, fontSize=9, leading=10))

    doc = SimpleDocTemplate(file_path, pagesize=A4, topMargin=20*mm, bottomMargin=20*mm, leftMargin=15*mm, rightMargin=15*mm)
    story = []

    logo_path = os.path.join(os.path.dirname(__file__), 'logo.png')
    if os.path.exists(logo_path):
        story.append(Image(logo_path, width=50*mm, height=20*mm, hAlign='CENTER'))
        story.append(Spacer(1, 5*mm))

    story.append(Paragraph("ТЕХНОЛОГИЧЕСКИЙ ПРОЦЕСС", styles['TitleCenter']))
    story.append(Paragraph(f"Модель: <b>{data['model']}</b>", styles['TitleCenter']))
    story.append(Spacer(1, 10*mm))

    if data['document_details']:
        d = data['document_details'][0]
        table_data = [["Параметр", "Значение"]] + [[k, v or "—"] for k, v in zip(["Организация", "Обозначение изделия", "Обозначение документа", "Разработал", "Проверил"], d)]
        t = Table(table_data, colWidths=[50*mm, 120*mm])
        t.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2E7D32')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.grey)]))
        story.append(Paragraph("Реквизиты", styles['Header']))
        story.append(t)
        story.append(Spacer(1, 8*mm))

    for key, cfg in TABLE_CONFIG.items():
        items = data.get(key, [])
        if not items: continue
        table_data = [cfg["headers"]]
        for i, item in enumerate(items, 1):
            row = []
            for j, field in enumerate(cfg["fields"]):
                val = str(item.get(field, "—"))
                if j in cfg.get("wrap", []) and len(val) > 30:
                    val = Paragraph(val, styles['Cell'])
                row.append(val if field != "number" else (val or str(i)))
            table_data.append(row)
        col_widths = [w*mm for w in cfg["col_widths"]]
        t = Table(table_data, colWidths=col_widths, rowHeights=cfg["row_height"]*mm)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor(cfg["color"])),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('FONTNAME', (0,0), (-1,-1), font_name),
            ('FONTSIZE', (0,1), (-1,-1), 9),
        ]))
        story.append(Paragraph(cfg["title"], styles['Header']))
        story.append(t)
        story.append(Spacer(1, 8*mm))

    story.append(Paragraph(f"Дата: {data['timestamp']}", ParagraphStyle(name='Footer', fontName=font_name, fontSize=9, alignment=TA_RIGHT)))
    doc.build(story, onFirstPage=lambda c, d: c.drawString(180*mm, 10*mm, f"Страница {c.getPageNumber()}"))
