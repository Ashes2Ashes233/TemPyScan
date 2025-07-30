# report_generator.py

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4


def generate_pdf_report(path, report_data, plot_data_list):
    doc = SimpleDocTemplate(path, pagesize=A4,
                            rightMargin=inch / 2, leftMargin=inch / 2,
                            topMargin=inch / 2, bottomMargin=inch / 2)

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Center', alignment=TA_CENTER))
    styles.add(ParagraphStyle(name='Left', alignment=TA_LEFT))
    styles.add(ParagraphStyle(name='Bold', fontName='Helvetica-Bold'))

    story = []

    # 1. 标题
    story.append(Paragraph("Microwave Frequency Analysis Report", styles['Heading1']))
    story.append(Spacer(1, 0.3 * inch))

    # 2. 基本信息
    info_data = [
        [Paragraph('Test Name:', styles['Bold']), Paragraph(report_data.get('Test name', ''), styles['Normal'])],
        [Paragraph('Test Type:', styles['Bold']), Paragraph(report_data.get('Test type', ''), styles['Normal'])],
        [Paragraph('Sample No.:', styles['Bold']), Paragraph(report_data.get('Sample number', ''), styles['Normal'])],
        [Paragraph('Model No.:', styles['Bold']), Paragraph(report_data.get('Model number', ''), styles['Normal'])],
        [Paragraph('Lab Request No.:', styles['Bold']),
         Paragraph(report_data.get('Lab request', ''), styles['Normal'])],
        [Paragraph('Tester:', styles['Bold']), Paragraph(report_data.get('Tester', ''), styles['Normal'])],
        [Paragraph('Equipment:', styles['Bold']), Paragraph(report_data.get('Equipment', ''), styles['Normal'])],
        [Paragraph('Start Time:', styles['Bold']), Paragraph(report_data.get('Start time', ''), styles['Normal'])],
    ]

    info_table = Table(info_data, colWidths=[1.8 * inch, 5.5 * inch])
    info_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
    ]))

    story.append(info_table)
    story.append(Spacer(1, 0.3 * inch))

    # 3. 观察结果
    story.append(Paragraph("Observations:", styles['Heading2']))
    story.append(Paragraph(report_data.get('Observations', '').replace('\n', '<br/>'), styles['BodyText']))
    story.append(Spacer(1, 0.3 * inch))

    # 4. 频率数据表格
    story.append(Paragraph("Top Frequencies:", styles['Heading2']))
    story.append(Spacer(1, 0.1 * inch))

    table_data = report_data.get('frequency_data', [])
    if not table_data:
        table_data = [["Rank", "Frequency (MHz)", "Count"]]

    table = Table(table_data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    story.append(table)
    story.append(PageBreak())

    # 5. 频率分布图
    story.append(Paragraph("Frequency Distribution", styles['Heading2']))
    story.append(Spacer(1, 0.2 * inch))

    if plot_data_list:
        for plot_info in plot_data_list:
            plot_title = plot_info.get('title', 'Frequency Distribution')
            plot_path = plot_info.get('path')

            if plot_path:
                try:
                    img = Image(plot_path, width=6.5 * inch, height=5 * inch)
                    story.append(img)
                    story.append(Spacer(1, 0.2 * inch))
                except Exception as e:
                    story.append(Paragraph(f"Error loading image: {e}", styles['BodyText']))

    # 生成PDF
    try:
        doc.build(story)
        return True
    except Exception as e:
        print(f"Error building PDF: {e}")
        return False