# report_generator.py

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4

# 接收一个 plot_data_list
def generate_pdf_report(path, report_data, plot_data_list):
    doc = SimpleDocTemplate(path, pagesize=A4,
                            rightMargin=inch / 2, leftMargin=inch / 2,
                            topMargin=inch / 2, bottomMargin=inch / 2)

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Center', alignment=TA_CENTER))
    styles.add(ParagraphStyle(name='Left', alignment=TA_LEFT))

    story = []

    # (1. 标题, 2. 基础信息, 3. 现象/结果/备注, 4. 测试数据表格 )
    story.append(Paragraph("Heating Test Report", styles['Title']))
    story.append(Spacer(1, 0.2 * inch))
    story.append(Paragraph("Heating Test Info:", styles['h2']))
    info_data = [
        [Paragraph('Test Name:', styles['Normal']), Paragraph(report_data.get('Test name', ''), styles['Normal'])],
        [Paragraph('Test Type:', styles['Normal']), Paragraph(report_data.get('Test type', ''), styles['Normal'])],
        [Paragraph('Sample No.:', styles['Normal']), Paragraph(report_data.get('Sample number', ''), styles['Normal'])],
        [Paragraph('Model No.:', styles['Normal']), Paragraph(report_data.get('Model number', ''), styles['Normal'])],
        [Paragraph('Rating voltage/freq.:', styles['Normal']),
         Paragraph(f"{report_data.get('Rating Voltage', '')} / {report_data.get('Rating Frequency', '')}",
                   styles['Normal'])],
        [Paragraph('Lab request no.:', styles['Normal']),
         Paragraph(report_data.get('Lab request', ''), styles['Normal'])],
        [Paragraph('Tester:', styles['Normal']), Paragraph(report_data.get('Tester', ''), styles['Normal'])],
        [Paragraph('Equipment:', styles['Normal']), Paragraph(report_data.get('Equipment', ''), styles['Normal'])],
        # 从settings中获取
        [Paragraph('Operating Voltage:', styles['Normal']),
         Paragraph(report_data.get('Operating Voltage', ''), styles['Normal'])],
        [Paragraph('Operating Frequency:', styles['Normal']),
         Paragraph(report_data.get('Operating Frequency', ''), styles['Normal'])],
        [Paragraph('Operating Duration:', styles['Normal']),
         Paragraph(report_data.get('Operating Duration', ''), styles['Normal'])],
        [Paragraph('Start time:', styles['Normal']), Paragraph(report_data.get('Start time', ''), styles['Normal'])],
        [Paragraph('Stop time:', styles['Normal']), Paragraph(report_data.get('Stop time', ''), styles['Normal'])],
        [Paragraph('Ambient Channel:', styles['Normal']),
         Paragraph(report_data.get('Ambient Channel', ''), styles['Normal'])],
        [Paragraph('Ambient Temperature:', styles['Normal']), Paragraph(
            f"Start(T1):{report_data.get('Ambient Temp Start', 'N/A')}  Stop(T2):{report_data.get('Ambient Temp Stop', 'N/A')}",
            styles['Normal'])],
    ]
    info_table = Table(info_data, colWidths=[1.8 * inch, 5.5 * inch])
    info_table.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'LEFT'), ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                                    ('BOTTOMPADDING', (0, 0), (-1, -1), 2), ('TOPPADDING', (0, 0), (-1, -1), 2), ]))
    story.append(info_table)
    story.append(Spacer(1, 0.2 * inch))
    story.append(Paragraph("Phenomena And Result:", styles['h2']))
    story.append(Paragraph(report_data.get('Phenomena And Result', '').replace('\n', '<br/>'), styles['BodyText']))
    story.append(Spacer(1, 0.2 * inch))
    #story.append(Paragraph("Check result:", styles['h2']));
    #story.append(Paragraph(report_data.get('Check result', '').replace('\n', '<br/>'), styles['BodyText']));
    #story.append(Spacer(1, 0.2 * inch))
    story.append(Paragraph("Notes:", styles['h2']))
    story.append(Paragraph(report_data.get('Notes', '').replace('\n', '<br/>'), styles['BodyText']))
    story.append(Spacer(1, 0.2 * inch))
    story.append(PageBreak())

    story.append(Paragraph("Test Data", styles['h2']))
    story.append(Spacer(1, 0.1 * inch))
    data_table_content = [["Item", "Test Location", "Channel No.", "Max Temp.(°C)", "Limit(°C)", "P/F"]]
    test_data = report_data.get('test_data', [])
    for row in test_data: data_table_content.append(row)
    data_table = Table(data_table_content,
                       colWidths=[0.5 * inch, 1.5 * inch, 1 * inch, 1.2 * inch, 1 * inch, 0.8 * inch])
    data_table.setStyle(TableStyle(
        [('BACKGROUND', (0, 0), (-1, 0), colors.grey), ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
         ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
         ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
         ('BACKGROUND', (0, 1), (-1, -1), colors.beige), ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
    story.append(data_table)
    story.append(PageBreak())

    # 循环添加多个图像
    #story.append(Spacer(1, 0.2 * inch))
    story.append(Paragraph("Test Graph", styles['h2']))

    if plot_data_list:
        for plot_info in plot_data_list:
            #plot_title = plot_info.get('title', 'Test Graph')
            plot_path = plot_info.get('path')

            # 添加每个图的标题 (已废弃）
            #story.append(Spacer(1, 0.2 * inch))
            #story.append(Paragraph(plot_title, styles['h3']))
            #story.append(Spacer(1, 0.1 * inch))

            if plot_path:
                try:
                    img = Image(plot_path, width=7 * inch, height=5.25 * inch, kind='proportional')
                    story.append(img)
                except Exception as e:
                    story.append(Paragraph(f"Error loading image: {e}", styles['BodyText']))

    try:
        doc.build(story)
        return True
    except Exception as e:
        print(f"Error building PDF: {e}")
        return False