"""
Generate Executive Summary PDF from dashboard data
"""
import json
import os
from datetime import datetime
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend
import matplotlib.pyplot as plt
import numpy as np

def generate_executive_summary_pdf(data, output_path='ECN_Executive_Summary.pdf'):
    """
    Generate a PDF executive summary from dashboard data
    """

    # Create the PDF document
    doc = SimpleDocTemplate(output_path, pagesize=letter,
                           topMargin=0.5*inch, bottomMargin=0.5*inch,
                           leftMargin=0.5*inch, rightMargin=0.5*inch)

    # Container for the 'Flowable' objects
    elements = []

    # Define styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=colors.HexColor('#1f4788'),
        spaceAfter=30,
        alignment=TA_CENTER
    )

    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=16,
        textColor=colors.HexColor('#1f4788'),
        spaceAfter=12,
        spaceBefore=12
    )

    normal_style = styles['Normal']

    # Add logo if available
    logo_file = 'ADI-Logo-RGB-FullColor.png'
    if os.path.exists(logo_file):
        img = Image(logo_file, width=2*inch, height=0.5*inch)
        img.hAlign = 'CENTER'
        elements.append(img)
        elements.append(Spacer(1, 0.2*inch))

    # Title
    title = Paragraph("Backend Foundry ECN Cycle Time", title_style)
    elements.append(title)

    subtitle = Paragraph(f"Executive Summary - {datetime.now().strftime('%B %d, %Y')}", normal_style)
    subtitle.alignment = TA_CENTER
    elements.append(subtitle)
    elements.append(Spacer(1, 0.3*inch))

    # Overview Statistics
    elements.append(Paragraph("Overview Statistics", heading_style))

    stats = data.get('overall_stats', {})
    overview_data = [
        ['Metric', 'Value'],
        ['Total ECN Requests', f"{stats.get('total_requests', 0):,}"],
        ['Closed ECNs', f"{stats.get('total_closed', 0):,}"],
        ['Void ECNs', f"{stats.get('total_void', 0):,}"],
        ['Avg Processing CT', f"{stats.get('avg_proc_ct', 0):.1f} days"],
        ['Avg Total CT', f"{stats.get('avg_total_ct', 0):.1f} days"],
        ['Median Processing CT', f"{stats.get('median_proc_ct', 0):.1f} days"],
        ['90th Percentile (Proc CT)', f"{stats.get('percentile_90_proc_ct', 0):.1f} days"],
    ]

    overview_table = Table(overview_data, colWidths=[3*inch, 2*inch])
    overview_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
    ]))

    elements.append(overview_table)
    elements.append(Spacer(1, 0.3*inch))

    # ECN Volume by Topic
    elements.append(Paragraph("ECN Volume by Topic", heading_style))

    topic_data = data.get('topic_comparison', [])[:10]  # Top 10
    if topic_data:
        topic_table_data = [['ECN Topic', 'Count', 'Avg Proc CT (days)']]
        for item in topic_data:
            topic_table_data.append([
                item.get('ECN Topic', 'N/A'),
                f"{item.get('RequestNum', 0):,}",
                f"{item.get('ProcCT(days)', 0):.1f}"
            ])

        topic_table = Table(topic_table_data, colWidths=[3.5*inch, 1*inch, 1.5*inch])
        topic_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
        ]))

        elements.append(topic_table)

    elements.append(PageBreak())

    # Top 10 Coordinators
    elements.append(Paragraph("Top 10 ECN Coordinators by Volume", heading_style))

    coordinator_data = data.get('coordinator_comparison', [])[:10]
    if coordinator_data:
        coord_table_data = [['Coordinator', 'Region', 'Count', 'Avg Proc CT (days)']]
        for item in coordinator_data:
            coord_table_data.append([
                item.get('ECNCoordinator', 'N/A'),
                item.get('Region', 'N/A'),
                f"{item.get('RequestNum', 0):,}",
                f"{item.get('ProcCT(days)', 0):.1f}"
            ])

        coord_table = Table(coord_table_data, colWidths=[2.5*inch, 1.5*inch, 1*inch, 1.5*inch])
        coord_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
        ]))

        elements.append(coord_table)

    elements.append(Spacer(1, 0.3*inch))

    # Process Outlier Analysis
    elements.append(Paragraph("Process Optimization - Outlier Analysis", heading_style))

    outlier_data = data.get('outlier_analysis', {})
    if outlier_data:
        outlier_summary_data = [
            ['Metric', 'Value'],
            ['Total Outlier ECNs', f"{outlier_data.get('total_outliers', 0):,}"],
            ['Excess Days', f"{outlier_data.get('total_excess_days', 0):,.0f}"],
            ['Avg Excess per Outlier', f"{outlier_data.get('avg_excess_per_outlier', 0):.1f} days"],
        ]

        outlier_summary_table = Table(outlier_summary_data, colWidths=[3*inch, 2*inch])
        outlier_summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#dc2626')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
        ]))

        elements.append(outlier_summary_table)
        elements.append(Spacer(1, 0.2*inch))

        # High variability ECN types
        high_var_types = outlier_data.get('high_variability_types', [])[:5]
        if high_var_types:
            var_text = Paragraph("<b>ECN Types with Highest Variability (Need Process Standardization):</b>", normal_style)
            elements.append(var_text)
            elements.append(Spacer(1, 0.1*inch))

            var_table_data = [['ECN Topic', 'Count', 'CV', 'Range (days)']]
            for item in high_var_types:
                var_table_data.append([
                    item.get('ECN Topic', 'N/A'),
                    f"{item.get('Count', 0):,}",
                    f"{item.get('CV', 0):.2f}",
                    f"{item.get('Min', 0):.0f} - {item.get('Max', 0):.0f}"
                ])

            var_table = Table(var_table_data, colWidths=[3*inch, 1*inch, 0.8*inch, 1.7*inch])
            var_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#dc2626')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 9),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
            ]))

            elements.append(var_table)
            elements.append(Spacer(1, 0.2*inch))

        # Link to detailed analysis
        link_text = Paragraph(
            "<b>For detailed outlier analysis, see: ECN_Outlier_Analysis.xlsx</b><br/>"
            "<i>(Includes breakdowns by coordinator, site, and ECN type)</i>",
            normal_style
        )
        elements.append(link_text)

    elements.append(Spacer(1, 0.3*inch))

    # 90th Percentile Analysis
    elements.append(Paragraph("90th Percentile ECNs by Topic (Slowest Processing)", heading_style))

    percentile_data = stats.get('top_categories_90th', [])[:10]
    if percentile_data:
        perc_table_data = [['ECN Topic', 'Count', 'Avg Proc CT (days)']]
        for item in percentile_data:
            perc_table_data.append([
                item.get('ECN Topic', 'N/A'),
                f"{item.get('RequestNum', 0):,}",
                f"{item.get('ProcCT(days)', 0):.1f}"
            ])

        perc_table = Table(perc_table_data, colWidths=[4*inch, 1*inch, 1.5*inch])
        perc_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#dc2626')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
        ]))

        elements.append(perc_table)

    # Footer
    elements.append(Spacer(1, 0.5*inch))
    footer_text = Paragraph(
        f"<i>Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</i>",
        normal_style
    )
    footer_text.alignment = TA_CENTER
    elements.append(footer_text)

    # Build PDF
    doc.build(elements)
    print(f"[OK] Executive Summary PDF generated: {output_path}")

    return output_path


if __name__ == '__main__':
    # Read data
    with open('data.json', 'r') as f:
        data = json.load(f)

    # Generate PDF
    generate_executive_summary_pdf(data)
