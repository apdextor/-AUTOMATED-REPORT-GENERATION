import pandas as pd
import numpy as np
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, HRFlowable, Frame, PageTemplate
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.charts.lineplots import LinePlot
from reportlab.graphics.shapes import Drawing, String
import os

# Generate sample data if needed
def generate_sample_data(filename, rows=100):
    categories = [
        'Electronics', 'Clothing', 'Groceries', 'Furniture', 'Books',
        'Toys', 'Sports', 'Beauty', 'Automotive', 'Garden'
    ]
    countries = [
        'US', 'UK', 'CA', 'AU', 'DE', 'FR', 'IN', 'CN', 'JP', 'BR'
    ]
    products = [
        f'Product {chr(65+i)}' for i in range(20)
    ]
    
    # Use 2024 for the data period
    dates = pd.date_range(start='2024-01-01', end='2024-12-31', periods=rows)
    data = {
        'Date': np.random.choice(dates, rows),
        'Product': np.random.choice(products, rows),
        'Category': np.random.choice(categories, rows),
        'Amount': np.random.uniform(10, 5000, rows).round(2),
        'Country': np.random.choice(countries, rows)
    }
    df = pd.DataFrame(data)
    df.to_csv(filename, index=False)
    return df

# Analyze data
def analyze_data(df):
    results = {}
    df['Date'] = pd.to_datetime(df['Date'])
    
    # Basic metrics
    results['total_sales'] = df['Amount'].sum()
    results['avg_sale'] = df['Amount'].mean()
    results['transactions'] = len(df)
    
    # Grouped metrics
    results['by_category'] = df.groupby('Category')['Amount'].sum().sort_values(ascending=False)
    results['by_country'] = df.groupby('Country')['Amount'].sum().sort_values(ascending=False)
    results['by_month'] = df.groupby(df['Date'].dt.strftime('%Y-%m'))['Amount'].sum()
    results['top_products'] = df.groupby('Product')['Amount'].sum().nlargest(5)
    
    return results

def add_page_number(canvas, doc):
    page_num = canvas.getPageNumber()
    text = f"Page {page_num}"
    canvas.saveState()
    canvas.setFont('Helvetica', 9)
    canvas.setFillColor(colors.grey)
    canvas.drawRightString(200*mm, 10*mm, text)
    canvas.restoreState()

# Generate PDF report
def create_pdf_report(filename, data, analysis):
    styles = getSampleStyleSheet()
    # Custom styles
    title_style = ParagraphStyle(
        'Title',
        parent=styles['Heading1'],
        fontSize=28,
        alignment=1,
        spaceAfter=20,
        fontName='Helvetica-Bold'
    )
    heading_style = ParagraphStyle(
        'Heading2',
        parent=styles['Heading2'],
        fontSize=18,
        spaceAfter=12,
        textColor=colors.darkblue,
        fontName='Helvetica-Bold'
    )
    subheading_style = ParagraphStyle(
        'Heading3',
        parent=styles['Heading3'],
        fontSize=14,
        spaceAfter=8,
        textColor=colors.darkgreen,
        fontName='Helvetica-Bold'
    )
    body_style = ParagraphStyle(
        'BodyText',
        parent=styles['BodyText'],
        fontSize=11,
        leading=16,
        spaceAfter=8,
        fontName='Helvetica'
    )
    small_style = ParagraphStyle(
        'Small',
        parent=styles['BodyText'],
        fontSize=9,
        leading=12,
        textColor=colors.grey
    )

    doc = SimpleDocTemplate(
        filename,
        pagesize=letter,
        rightMargin=36, leftMargin=36,
        topMargin=36, bottomMargin=36
    )
    story = []

    # Cover Page
    story.append(Paragraph(f"Annual Sales Report", title_style))
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph(f"<b>Generated on:</b> {datetime.now().strftime('%B %d, %Y %H:%M')}", body_style))
    story.append(Paragraph(
        f"<b>Data Period:</b> {data['Date'].min().date()} to {data['Date'].max().date()}",
        body_style))
    story.append(Spacer(1, 0.5*inch))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.darkblue))
    story.append(Spacer(1, 0.3*inch))
    story.append(Paragraph(
        f"<b>Total Sales:</b> <font size=16 color='darkgreen'>${analysis['total_sales']:,.2f}</font>", heading_style))
    story.append(Spacer(1, 0.2*inch))
    # story.append(Paragraph(
    #     "Prepared by: <b>Sales Analytics Team</b><br/>Confidential - For internal use only.",
    #     small_style))
    # story.append(PageBreak())

    # Executive Summary
    story.append(Paragraph("Executive Summary", heading_style))
    story.append(Spacer(1, 0.1*inch))
    year = data['Date'].min().year if not data.empty else 2024
    story.append(Paragraph(
        f"In {year}, our company processed <b>{analysis['transactions']}</b> sales transactions across "
        f"<b>{len(data['Category'].unique())}</b> product categories and <b>{len(data['Country'].unique())}</b> countries. "
        f"Total sales reached <b>${analysis['total_sales']:,.2f}</b>, with an average transaction value of "
        f"<b>${analysis['avg_sale']:,.2f}</b>. The <b>{analysis['by_category'].index[0]}</b> category led all segments, "
        f"contributing <b>${analysis['by_category'].iloc[0]:,.2f}</b> to the annual revenue.",
        body_style))
    story.append(Paragraph(
        "This report provides a comprehensive overview of sales performance, highlights top products, and identifies "
        "key trends and opportunities for growth. The following sections offer detailed breakdowns by category, region, and month.",
        body_style))
    story.append(Spacer(1, 0.2*inch))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.grey))
    story.append(Spacer(1, 0.2*inch))

    # Key Metrics Table
    metrics_data = [
        ['Total Sales', f"${analysis['total_sales']:,.2f}"],
        ['Transactions', f"{analysis['transactions']}"],
        ['Average Sale', f"${analysis['avg_sale']:,.2f}"],
        ['Top Category', f"{analysis['by_category'].index[0]} (${analysis['by_category'].iloc[0]:,.2f})"]
    ]
    metrics_table = Table(metrics_data, colWidths=[2.5*inch, 2.5*inch])
    metrics_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.lightblue),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 12),
        ('BOTTOMPADDING', (0,0), (-1,0), 12),
        ('BACKGROUND', (0,1), (-1,-1), colors.whitesmoke),
        ('GRID', (0,0), (-1,-1), 1, colors.grey)
    ]))
    story.append(metrics_table)
    story.append(Spacer(1, 0.3*inch))
    story.append(Paragraph(
        "The table above summarizes the most important sales metrics for the year. "
        "The top category contributed significantly to the overall revenue.",
        body_style))
    story.append(PageBreak())

    # Category Performance
    story.append(Paragraph("Sales by Category", heading_style))
    story.append(Spacer(1, 0.1*inch))
    story.append(Paragraph(
        "The charts below illustrate the distribution of sales across different product categories. "
        "These visualizations help identify which categories are driving the most revenue and where to focus future efforts.",
        body_style))

    # Bar Chart for Category
    drawing_bar = Drawing(450, 250)
    bc = VerticalBarChart()
    bc.x = 70
    bc.y = 50
    bc.height = 180
    bc.width = 350
    bc.data = [list(analysis['by_category'].values)]
    bc.categoryAxis.categoryNames = list(analysis['by_category'].index)
    bc.valueAxis.valueMin = 0
    bc.valueAxis.valueMax = float(max(analysis['by_category'].values)) * 1.15
    bc.valueAxis.valueStep = float(max(analysis['by_category'].values)) / 5 if max(analysis['by_category'].values) else 1
    bc.bars[0].fillColor = colors.HexColor("#3b5998")
    bc.bars[0].strokeColor = colors.HexColor("#1a237e")
    bc.bars[0].strokeWidth = 1.5
    bc.categoryAxis.labels.angle = 30
    bc.categoryAxis.labels.dx = 8
    bc.categoryAxis.labels.dy = -2
    bc.categoryAxis.labels.fontSize = 9
    bc.valueAxis.labels.fontSize = 9
    bc.categoryAxis.visibleGrid = True
    bc.valueAxis.visibleGrid = True
    bc.valueAxis.gridStrokeColor = colors.lightgrey
    bc.categoryAxis.gridStrokeColor = colors.lightgrey
    drawing_bar.add(bc)
    drawing_bar.add(String(220, 20, "Category", fontSize=13, fontName='Helvetica-Bold'))
    drawing_bar.add(String(4, 125, "Sales", fontSize=13, fontName='Helvetica-Bold', angle=90))
    story.append(Paragraph("Bar Chart: Sales by Category", subheading_style))
    story.append(drawing_bar)
    story.append(Spacer(1, 0.2*inch))

    # Pie Chart for Category
    drawing_pie = Drawing(350, 220)
    pie = Pie()
    pie.x = 100
    pie.y = 35
    pie.width = 150
    pie.height = 150
    pie.data = list(analysis['by_category'].values)
    pie.labels = list(analysis['by_category'].index)
    pie.slices.strokeWidth = 1
    pie.slices.strokeColor = colors.white
    pie.slices[0].fillColor = colors.HexColor("#3b5998")
    pie.slices[1].fillColor = colors.HexColor("#8b9dc3")
    pie.slices[2].fillColor = colors.HexColor("#00a79d")
    pie.slices[3].fillColor = colors.HexColor("#f39c12")
    pie.slices[4].fillColor = colors.HexColor("#e74c3c")
    # Add more colors if needed
    pie.sideLabels = True
    pie.slices.fontSize = 9
    pie.slices.fontName = 'Helvetica-Bold'
    drawing_pie.add(pie)
    drawing_pie.add(String(150, 10, "Category Share", fontSize=13, fontName='Helvetica-Bold'))
    story.append(Paragraph("Pie Chart: Category Market Share", subheading_style))
    story.append(drawing_pie)
    story.append(Spacer(1, 0.2*inch))
    story.append(PageBreak())

    # Line Chart for Category (showing monthly trend per top 3 categories)
    top3_cats = analysis['by_category'].index[:3]
    monthly = data.copy()
    monthly['Month'] = monthly['Date'].dt.strftime('%Y-%m')
    cat_month = monthly.groupby(['Month', 'Category'])['Amount'].sum().unstack().fillna(0)
    months = list(cat_month.index)
    drawing_line = Drawing(450, 250)
    lp = LinePlot()
    lp.x = 70
    lp.y = 50
    lp.height = 180
    lp.width = 350
    lp.data = [
        [(i, cat_month.iloc[i][cat]) for i in range(len(months))]
        for cat in top3_cats
    ]
    max_sales = cat_month[top3_cats].values.max()
    lp.yValueAxis.valueMin = 0
    lp.yValueAxis.valueMax = float(max_sales) * 1.15
    lp.yValueAxis.valueStep = float(max_sales) / 5 if max_sales else 1
    lp.xValueAxis.valueMin = 0
    lp.xValueAxis.valueMax = len(months) - 1
    lp.xValueAxis.valueSteps = list(range(len(months)))
    lp.xValueAxis.labelTextFormat = lambda i: months[int(i)] if int(i) < len(months) else ""
    lp.xValueAxis.visibleGrid = True
    lp.yValueAxis.visibleGrid = True
    lp.xValueAxis.labels.fontSize = 8
    lp.yValueAxis.labels.fontSize = 8
    lp.lines[0].strokeColor = colors.HexColor("#3b5998")
    lp.lines[1].strokeColor = colors.HexColor("#f39c12")
    lp.lines[2].strokeColor = colors.HexColor("#e74c3c")
    drawing_line.add(lp)
    drawing_line.add(String(20, 130, "Sales", fontSize=12, angle=90))
    drawing_line.add(String(220, 15, "Month", fontSize=12))
    legend_y = 220
    for i, cat in enumerate(top3_cats):
        drawing_line.add(String(300, legend_y, cat, fontSize=8))
        legend_y -= 15
    story.append(Paragraph("Line Chart: Monthly Sales Trend (Top 3 Categories)", subheading_style))
    story.append(drawing_line)
    story.append(Spacer(1, 0.3*inch))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.grey))
    story.append(Spacer(1, 0.2*inch))

    # Top Products Table
    story.append(Paragraph("Top Performing Products", heading_style))
    story.append(Spacer(1, 0.1*inch))
    story.append(Paragraph(
        "The following charts and table list the top five products by total sales. "
        "These products have shown outstanding performance during the year and represent key revenue drivers.",
        body_style))

    # Bar Chart for Top Products
    drawing_prod_bar = Drawing(450, 250)
    bc_prod = VerticalBarChart()
    bc_prod.x = 70
    bc_prod.y = 50
    bc_prod.height = 180
    bc_prod.width = 350
    bc_prod.data = [list(analysis['top_products'].values)]
    bc_prod.categoryAxis.categoryNames = list(analysis['top_products'].index)
    bc_prod.valueAxis.valueMin = 0
    bc_prod.valueAxis.valueMax = float(max(analysis['top_products'].values)) * 1.15
    bc_prod.valueAxis.valueStep = float(max(analysis['top_products'].values)) / 5 if max(analysis['top_products'].values) else 1
    bc_prod.bars[0].fillColor = colors.HexColor("#00a79d")
    bc_prod.bars[0].strokeColor = colors.HexColor("#00695c")
    bc_prod.bars[0].strokeWidth = 1.5
    bc_prod.categoryAxis.labels.angle = 30
    bc_prod.categoryAxis.labels.dx = 8
    bc_prod.categoryAxis.labels.dy = -2
    bc_prod.categoryAxis.labels.fontSize = 9
    bc_prod.valueAxis.labels.fontSize = 9
    bc_prod.categoryAxis.visibleGrid = True
    bc_prod.valueAxis.visibleGrid = True
    bc_prod.valueAxis.gridStrokeColor = colors.lightgrey
    bc_prod.categoryAxis.gridStrokeColor = colors.lightgrey
    drawing_prod_bar.add(bc_prod)
    drawing_prod_bar.add(String(220, 20, "Product", fontSize=13, fontName='Helvetica-Bold'))
    drawing_prod_bar.add(String(4, 125, "Sales", fontSize=13, fontName='Helvetica-Bold', angle=90))
    story.append(Paragraph("Bar Chart: Top 5 Products by Sales", subheading_style))
    story.append(drawing_prod_bar)
    story.append(Spacer(1, 0.2*inch))
    story.append(PageBreak())

    # Pie Chart for Top Products
    drawing_prod_pie = Drawing(350, 220)
    pie_prod = Pie()
    pie_prod.x = 100
    pie_prod.y = 35
    pie_prod.width = 150
    pie_prod.height = 150
    pie_prod.data = list(analysis['top_products'].values)
    pie_prod.labels = list(analysis['top_products'].index)
    pie_prod.slices.strokeWidth = 1
    pie_prod.slices.strokeColor = colors.white
    pie_prod.slices[0].fillColor = colors.HexColor("#00a79d")
    pie_prod.slices[1].fillColor = colors.HexColor("#f39c12")
    pie_prod.slices[2].fillColor = colors.HexColor("#e74c3c")
    pie_prod.slices[3].fillColor = colors.HexColor("#3b5998")
    pie_prod.slices[4].fillColor = colors.HexColor("#8b9dc3")
    pie_prod.sideLabels = True
    pie_prod.slices.fontSize = 9
    pie_prod.slices.fontName = 'Helvetica-Bold'
    drawing_prod_pie.add(pie_prod)
    drawing_prod_pie.add(String(150, 15, "Product Share", fontSize=13, fontName='Helvetica-Bold'))
    story.append(Paragraph("Pie Chart: Top 5 Products Market Share", subheading_style))
    story.append(drawing_prod_pie)
    story.append(Spacer(1, 0.2*inch))
    
    # Line Chart for Top Products (monthly trend)
    top_products = analysis['top_products'].index
    prod_month = monthly.groupby(['Month', 'Product'])['Amount'].sum().unstack().fillna(0)
    drawing_prod_line = Drawing(450, 250)
    lp_prod = LinePlot()
    lp_prod.x = 70
    lp_prod.y = 50
    lp_prod.height = 180
    lp_prod.width = 350
    lp_prod.data = [
        [(i, prod_month.iloc[i][prod]) for i in range(len(months))]
        for prod in top_products
    ]
    max_prod_sales = prod_month[top_products].values.max()
    lp_prod.yValueAxis.valueMin = 0
    lp_prod.yValueAxis.valueMax = float(max_prod_sales) * 1.15
    lp_prod.yValueAxis.valueStep = float(max_prod_sales) / 5 if max_prod_sales else 1
    lp_prod.xValueAxis.valueMin = 0
    lp_prod.xValueAxis.valueMax = len(months) - 1
    lp_prod.xValueAxis.valueSteps = list(range(len(months)))
    lp_prod.xValueAxis.labelTextFormat = lambda i: months[int(i)] if int(i) < len(months) else ""
    lp_prod.xValueAxis.visibleGrid = True
    lp_prod.yValueAxis.visibleGrid = True
    lp_prod.xValueAxis.labels.fontSize = 8
    lp_prod.yValueAxis.labels.fontSize = 8
    prod_colors = [
        colors.HexColor("#00a79d"),
        colors.HexColor("#f39c12"),
        colors.HexColor("#e74c3c"),
        colors.HexColor("#3b5998"),
        colors.HexColor("#8b9dc3"),
    ]
    for i, color in enumerate(prod_colors):
        lp_prod.lines[i].strokeColor = color
    drawing_prod_line.add(lp_prod)
    drawing_prod_line.add(String(20, 130, "Sales", fontSize=12, angle=90))
    drawing_prod_line.add(String(220, 15, "Month", fontSize=12))
    legend_y = 220
    for i, prod in enumerate(top_products):
        drawing_prod_line.add(String(300, legend_y, prod, fontSize=8))
        legend_y -= 15
    story.append(Paragraph("Line Chart: Monthly Sales Trend (Top 5 Products)", subheading_style))
    story.append(drawing_prod_line)
    story.append(Spacer(1, 0.2*inch))

    # Table for Top Products
    top_prod_data = [['Product', 'Total Sales']]
    for product, sales in analysis['top_products'].items():
        top_prod_data.append([product, f"${sales:,.2f}"])
    prod_table = Table(top_prod_data, colWidths=[3*inch, 2*inch])
    prod_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#00a79d")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#e0f7fa")),
        ('GRID', (0,0), (-1,-1), 1, colors.grey)
    ]))
    story.append(prod_table)
    story.append(Spacer(1, 0.2*inch))
    story.append(PageBreak())

    # Regional Performance
    story.append(Paragraph("Regional Performance", heading_style))
    story.append(Spacer(1, 0.1*inch))
    story.append(Paragraph(
        "This section provides a breakdown of sales by country, including each region's market share. "
        "Understanding regional performance is crucial for strategic planning and resource allocation.",
        body_style))

    # Bar Chart for Country
    drawing_country_bar = Drawing(450, 250)
    bc_country = VerticalBarChart()
    bc_country.x = 70
    bc_country.y = 50
    bc_country.height = 180
    bc_country.width = 350
    bc_country.data = [list(analysis['by_country'].values)]
    bc_country.categoryAxis.categoryNames = list(analysis['by_country'].index)
    bc_country.valueAxis.valueMin = 0
    bc_country.valueAxis.valueMax = float(max(analysis['by_country'].values)) * 1.15
    bc_country.valueAxis.valueStep = float(max(analysis['by_country'].values)) / 5 if max(analysis['by_country'].values) else 1
    bc_country.bars[0].fillColor = colors.HexColor("#f39c12")
    bc_country.bars[0].strokeColor = colors.HexColor("#b35400")
    bc_country.bars[0].strokeWidth = 1.5
    bc_country.categoryAxis.labels.angle = 30
    bc_country.categoryAxis.labels.dx = 8
    bc_country.categoryAxis.labels.dy = -2
    bc_country.categoryAxis.labels.fontSize = 9
    bc_country.valueAxis.labels.fontSize = 9
    bc_country.categoryAxis.visibleGrid = True
    bc_country.valueAxis.visibleGrid = True
    bc_country.valueAxis.gridStrokeColor = colors.lightgrey
    bc_country.categoryAxis.gridStrokeColor = colors.lightgrey
    drawing_country_bar.add(bc_country)
    drawing_country_bar.add(String(220, 20, "Country", fontSize=13, fontName='Helvetica-Bold'))
    drawing_country_bar.add(String(4, 125, "Sales", fontSize=13, fontName='Helvetica-Bold', angle=90))
    story.append(Paragraph("Bar Chart: Sales by Country", subheading_style))
    story.append(drawing_country_bar)
    story.append(Spacer(1, 0.2*inch))

    # Pie Chart for Country
    drawing_country_pie = Drawing(350, 220)
    pie_country = Pie()
    pie_country.x = 100
    pie_country.y = 35
    pie_country.width = 150
    pie_country.height = 150
    pie_country.data = list(analysis['by_country'].values)
    pie_country.labels = list(analysis['by_country'].index)
    pie_country.slices.strokeWidth = 1
    pie_country.slices.strokeColor = colors.white
    country_colors = [
        colors.HexColor("#f39c12"), colors.HexColor("#e74c3c"), colors.HexColor("#3b5998"),
        colors.HexColor("#00a79d"), colors.HexColor("#8b9dc3"), colors.HexColor("#27ae60"),
        colors.HexColor("#9b59b6"), colors.HexColor("#34495e"), colors.HexColor("#16a085"),
        colors.HexColor("#d35400")
    ]
    for i, color in enumerate(country_colors):
        if i < len(pie_country.slices):
            pie_country.slices[i].fillColor = color
    pie_country.sideLabels = True
    pie_country.slices.fontSize = 9
    pie_country.slices.fontName = 'Helvetica-Bold'
    drawing_country_pie.add(pie_country)
    drawing_country_pie.add(String(175, 15, "Country Share", fontSize=13, fontName='Helvetica-Bold'))
    story.append(Paragraph("Pie Chart: Country Market Share", subheading_style))
    story.append(drawing_country_pie)
    story.append(Spacer(1, 0.2*inch))
    story.append(PageBreak())

    # Line Chart for Country (monthly trend for top 3 countries)
    top3_countries = analysis['by_country'].index[:3]
    country_month = monthly.groupby(['Month', 'Country'])['Amount'].sum().unstack().fillna(0)
    drawing_country_line = Drawing(450, 250)
    lp_country = LinePlot()
    lp_country.x = 70
    lp_country.y = 50
    lp_country.height = 180
    lp_country.width = 350
    lp_country.data = [
        [(i, country_month.iloc[i][country]) for i in range(len(months))]
        for country in top3_countries
    ]
    max_country_sales = country_month[top3_countries].values.max()
    lp_country.yValueAxis.valueMin = 0
    lp_country.yValueAxis.valueMax = float(max_country_sales) * 1.15
    lp_country.yValueAxis.valueStep = float(max_country_sales) / 5 if max_country_sales else 1
    lp_country.xValueAxis.valueMin = 0
    lp_country.xValueAxis.valueMax = len(months) - 1
    lp_country.xValueAxis.valueSteps = list(range(len(months)))
    lp_country.xValueAxis.labelTextFormat = lambda i: months[int(i)] if int(i) < len(months) else ""
    lp_country.xValueAxis.visibleGrid = True
    lp_country.yValueAxis.visibleGrid = True
    lp_country.xValueAxis.labels.fontSize = 8
    lp_country.yValueAxis.labels.fontSize = 8
    country_colors = [
        colors.HexColor("#f39c12"),
        colors.HexColor("#e74c3c"),
        colors.HexColor("#3b5998")
    ]
    for i, color in enumerate(country_colors):
        lp_country.lines[i].strokeColor = color
    drawing_country_line.add(lp_country)
    drawing_country_line.add(String(20, 130, "Sales", fontSize=12, angle=90))
    drawing_country_line.add(String(220, 15, "Month", fontSize=12))
    legend_y = 220
    for i, country in enumerate(top3_countries):
        drawing_country_line.add(String(300, legend_y, country, fontSize=8))
        legend_y -= 15
    story.append(Paragraph("Line Chart: Monthly Sales Trend (Top 3 Countries)", subheading_style))
    story.append(drawing_country_line)
    story.append(Spacer(1, 0.2*inch))

    # Table for Country
    country_data = [['Country', 'Total Sales', 'Market Share']]
    for country, sales in analysis['by_country'].items():
        share = (sales / analysis['total_sales']) * 100
        country_data.append([
            country, 
            f"${sales:,.2f}", 
            f"{share:.1f}%"
        ])
    country_table = Table(country_data, colWidths=[1.5*inch, 1.5*inch, 1.5*inch])
    country_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#27ae60")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#eafaf1")),
        ('GRID', (0,0), (-1,-1), 1, colors.grey)
    ]))
    story.append(country_table)
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph(
        "The above table highlights the contribution of each country to the total sales. "
        "Regions with higher market share may present opportunities for further growth.",
        body_style))
    story.append(PageBreak())

    # Monthly Sales Trend
    story.append(Paragraph("Monthly Sales Trend", heading_style))
    story.append(Spacer(1, 0.1*inch))
    story.append(Paragraph(
        f"The following charts and table show the total sales for each month in {year}. "
        "This helps in identifying seasonal trends and planning inventory accordingly.",
        body_style))

    # Bar Chart for Monthly Sales
    drawing_month_bar = Drawing(450, 250)
    bc_month = VerticalBarChart()
    bc_month.x = 70
    bc_month.y = 50
    bc_month.height = 180
    bc_month.width = 350
    bc_month.data = [list(analysis['by_month'].values)]
    bc_month.categoryAxis.categoryNames = list(analysis['by_month'].index)
    bc_month.valueAxis.valueMin = 0
    bc_month.valueAxis.valueMax = float(max(analysis['by_month'].values)) * 1.15
    bc_month.valueAxis.valueStep = float(max(analysis['by_month'].values)) / 5 if max(analysis['by_month'].values) else 1
    bc_month.bars[0].fillColor = colors.HexColor("#9b59b6")
    bc_month.bars[0].strokeColor = colors.HexColor("#512da8")
    bc_month.bars[0].strokeWidth = 1.5
    bc_month.categoryAxis.labels.angle = 30
    bc_month.categoryAxis.labels.dx = 8
    bc_month.categoryAxis.labels.dy = -2
    bc_month.categoryAxis.labels.fontSize = 9
    bc_month.valueAxis.labels.fontSize = 9
    bc_month.categoryAxis.visibleGrid = True
    bc_month.valueAxis.visibleGrid = True
    bc_month.valueAxis.gridStrokeColor = colors.lightgrey
    bc_month.categoryAxis.gridStrokeColor = colors.lightgrey
    drawing_month_bar.add(bc_month)
    drawing_month_bar.add(String(220, 20, "Month", fontSize=13, fontName='Helvetica-Bold'))
    drawing_month_bar.add(String(4, 125, "Sales", fontSize=13, fontName='Helvetica-Bold', angle=90))
    story.append(Paragraph("Bar Chart: Sales by Month", subheading_style))
    story.append(drawing_month_bar)
    story.append(Spacer(1, 0.2*inch))

    # Pie Chart for Monthly Sales
    drawing_month_pie = Drawing(350, 220)
    pie_month = Pie()
    pie_month.x = 100
    pie_month.y = 35
    pie_month.width = 150
    pie_month.height = 150
    pie_month.data = list(analysis['by_month'].values)
    pie_month.labels = list(analysis['by_month'].index)
    pie_month.slices.strokeWidth = 1
    pie_month.slices.strokeColor = colors.white
    month_colors = [
        colors.HexColor("#9b59b6"), colors.HexColor("#e74c3c"), colors.HexColor("#3b5998"),
        colors.HexColor("#00a79d"), colors.HexColor("#8b9dc3"), colors.HexColor("#27ae60"),
        colors.HexColor("#f39c12"), colors.HexColor("#34495e"), colors.HexColor("#16a085"),
        colors.HexColor("#d35400"), colors.HexColor("#e67e22"), colors.HexColor("#2980b9")
    ]
    for i, color in enumerate(month_colors):
        if i < len(pie_month.slices):
            pie_month.slices[i].fillColor = color
    pie_month.sideLabels = True
    pie_month.slices.fontSize = 9
    pie_month.slices.fontName = 'Helvetica-Bold'
    drawing_month_pie.add(pie_month)
    drawing_month_pie.add(String(175, 15, "Month Share", fontSize=13, fontName='Helvetica-Bold'))
    story.append(Paragraph("Pie Chart: Monthly Sales Share", subheading_style))
    story.append(drawing_month_pie)
    story.append(Spacer(1, 0.2*inch))
    story.append(PageBreak())

    # Line Chart for Monthly Sales
    drawing_month_line = Drawing(450, 250)
    lp_month = LinePlot()
    lp_month.x = 70
    lp_month.y = 50
    lp_month.height = 180
    lp_month.width = 350
    lp_month.data = [[(i, v) for i, v in enumerate(list(analysis['by_month'].values))]]
    lp_month.lines[0].strokeColor = colors.HexColor("#9b59b6")
    max_month_sales = max(analysis['by_month'].values) if len(analysis['by_month']) > 0 else 1
    lp_month.yValueAxis.valueMin = 0
    lp_month.yValueAxis.valueMax = float(max_month_sales) * 1.15
    lp_month.yValueAxis.valueStep = float(max_month_sales) / 5 if max_month_sales else 1
    lp_month.xValueAxis.valueMin = 0
    lp_month.xValueAxis.valueMax = len(analysis['by_month']) - 1
    lp_month.xValueAxis.valueSteps = list(range(len(analysis['by_month'])))
    months_list = list(analysis['by_month'].index)
    lp_month.xValueAxis.labelTextFormat = lambda i: months_list[int(i)] if int(i) < len(months_list) else ""
    lp_month.xValueAxis.visibleGrid = True
    lp_month.yValueAxis.visibleGrid = True
    lp_month.xValueAxis.labels.fontSize = 8
    lp_month.yValueAxis.labels.fontSize = 8
    drawing_month_line.add(lp_month)
    drawing_month_line.add(String(20, 130, "Sales", fontSize=12, angle=90))
    drawing_month_line.add(String(220, 15, "Month", fontSize=12))
    story.append(Paragraph("Line Chart: Monthly Sales Trend", subheading_style))
    story.append(drawing_month_line)
    story.append(Spacer(1, 0.2*inch))

    # Table for Monthly Sales
    month_data = [['Month', 'Total Sales']]
    for month, sales in analysis['by_month'].items():
        month_data.append([month, f"${sales:,.2f}"])
    month_table = Table(month_data, colWidths=[2*inch, 2*inch])
    month_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#9b59b6")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#f5eaf7")),
        ('GRID', (0,0), (-1,-1), 1, colors.grey)
    ]))
    story.append(month_table)
    story.append(Spacer(1, 0.2*inch))
    # story.append(Paragraph(
    #     "Monitoring monthly sales enables the business to adapt to changing market conditions and optimize sales strategies.",
    #     body_style))
    story.append(PageBreak())

    # Conclusion
    story.append(Paragraph("Conclusion & Recommendations", heading_style))
    story.append(Spacer(1, 0.1*inch))
    story.append(Paragraph(
        f"The analysis of {year} sales data reveals strong performance in several categories and regions. "
        "To capitalize on these trends, we recommend focusing marketing efforts on top-performing products and expanding in high-growth regions. "
        "Continuous monitoring of monthly trends will help anticipate demand fluctuations and optimize inventory management.",
        body_style))
    story.append(Spacer(1, 0.2*inch))
    # story.append(Paragraph(
    #     "For further details or custom analysis, please contact the Sales Analytics Team.",
    #     small_style))

    # Build with page numbers   
    def on_page(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 9)
        canvas.setFillColor(colors.grey)
        page_num = canvas.getPageNumber()
        canvas.drawRightString(7.5*inch, 0.5*inch, f"Page {page_num}")
        canvas.restoreState()

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)

if __name__ == "__main__":
    # Configuration
    input_file = "sales_data.csv"
    # Update output PDF filename to 2024
    output_pdf = "Sales_Report_2024.pdf"

    print(f"Current working directory: {os.getcwd()}")

    # Always generate new sample data for 100 rows
    df = generate_sample_data(input_file, rows=100)
    print(f"Sample data generated and saved to {input_file}")

    # Analyze data and generate report
    analysis = analyze_data(df)
    create_pdf_report(output_pdf, df, analysis)
    print(f"Report generated: {output_pdf}")