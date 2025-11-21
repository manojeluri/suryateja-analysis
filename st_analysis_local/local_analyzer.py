"""
Enhanced Local Sales Data Analyzer
Reads from local .xls file and generates comprehensive analytics
"""

import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime
import os
import warnings
warnings.filterwarnings('ignore')
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import Rectangle, FancyBboxPatch

# Set style for better visualizations
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (12, 6)
plt.rcParams['font.size'] = 10

# ============================================
# Company Product Mappings - Loaded from CSV files
# ============================================
def load_company_products_from_csv(folder_path="Company Wise Products"):
    """
    Load company-product mappings from CSV files in the specified folder
    Each CSV file should be named: CompanyName_Products.csv
    """
    company_products = {}

    if not os.path.exists(folder_path):
        print(f"⚠️  Warning: '{folder_path}' folder not found. Using empty mappings.")
        return company_products

    # Get all CSV files
    csv_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.csv')]

    print(f"\n📁 Loading company products from '{folder_path}/' folder...")
    print(f"   Found {len(csv_files)} CSV files")

    for csv_file in csv_files:
        # Extract company name from filename
        # Example: "Gharda_Products.csv" -> "Gharda"
        company_name = csv_file.replace('_Products.csv', '').replace('_Product_Names.csv', '')

        # Clean up company name (handle special cases)
        company_name = company_name.replace('_', ' ')
        if company_name == 'BestAgrolife':
            company_name = 'Best Agrolife'
        elif company_name == 'NovaAgriScience':
            company_name = 'Nova Agri Science'
        elif company_name == 'Nova Agri Tech':
            pass  # Already correct
        elif company_name == 'T Stanes':
            pass  # Already correct

        # Read CSV file
        file_path = os.path.join(folder_path, csv_file)
        try:
            df = pd.read_csv(file_path)

            # Get product names (assuming first column is "Product Name")
            if 'Product Name' in df.columns:
                products = df['Product Name'].dropna().str.strip().tolist()
            else:
                # If no header, assume first column contains products
                products = df.iloc[:, 0].dropna().str.strip().tolist()

            # Remove empty strings
            products = [p for p in products if p]

            company_products[company_name] = products
            print(f"   ✅ {company_name}: {len(products)} products")

        except Exception as e:
            print(f"   ⚠️  Error reading {csv_file}: {str(e)}")

    print(f"\n✅ Loaded {len(company_products)} companies with {sum(len(p) for p in company_products.values())} total products")

    return company_products

# Load company products from CSV files
COMPANY_PRODUCTS = load_company_products_from_csv()

# Create reverse mapping: product -> company
PRODUCT_TO_COMPANY = {}
for company, products in COMPANY_PRODUCTS.items():
    for product in products:
        PRODUCT_TO_COMPANY[product] = company

# ============================================
# Enhanced Analysis Functions
# ============================================

def load_data(filename):
    """Load data from XLS file"""
    print(f"\n📂 Loading data from: {filename}")

    if not os.path.exists(filename):
        raise FileNotFoundError(f"File not found: {filename}")

    # Try different engines for reading XLS
    try:
        df = pd.read_excel(filename, engine='xlrd')
    except:
        try:
            df = pd.read_excel(filename, engine='openpyxl')
        except Exception as e:
            raise Exception(f"Failed to read Excel file: {str(e)}")

    # Map column names
    column_mapping = {
        'NAMT': 'TAXBLEAMT',
        'PER': 'GST'
    }
    df.rename(columns=column_mapping, inplace=True)

    print(f"✅ Loaded {len(df)} rows and {len(df.columns)} columns")
    print(f"📋 Columns: {', '.join(df.columns.tolist())}")

    return df

def categorize_by_company(df):
    """Add company categorization to dataframe"""
    df['COMPANY'] = df['ITNAME'].map(PRODUCT_TO_COMPANY).fillna('Other')
    return df

def calculate_advanced_metrics(df):
    """Calculate advanced business metrics"""
    metrics = {}

    # Price per unit
    df['PRICE_PER_UNIT'] = df['TAXBLEAMT'] / df['QTY']

    # GST Amount
    df['GST_AMOUNT'] = df['TAXBLEAMT'] * df['GST'] / 100

    # Total amount including GST
    df['TOTAL_WITH_GST'] = df['TAXBLEAMT'] + df['GST_AMOUNT']

    # Calculate percentile rankings
    df['REVENUE_PERCENTILE'] = df['TAXBLEAMT'].rank(pct=True) * 100
    df['QTY_PERCENTILE'] = df['QTY'].rank(pct=True) * 100

    # Performance score (normalized combination of quantity and revenue)
    revenue_norm = (df['TAXBLEAMT'] - df['TAXBLEAMT'].min()) / (df['TAXBLEAMT'].max() - df['TAXBLEAMT'].min())
    qty_norm = (df['QTY'] - df['QTY'].min()) / (df['QTY'].max() - df['QTY'].min())
    df['PERFORMANCE_SCORE'] = (revenue_norm * 0.6 + qty_norm * 0.4) * 100

    return df

def generate_company_analysis(df):
    """Generate comprehensive company-wise analysis"""
    company_summary = df.groupby('COMPANY').agg({
        'ITNAME': 'count',
        'QTY': ['sum', 'mean', 'std'],
        'TAXBLEAMT': ['sum', 'mean', 'std'],
        'GST_AMOUNT': 'sum',
        'TOTAL_WITH_GST': 'sum',
        'PRICE_PER_UNIT': 'mean'
    }).round(2)

    # Flatten column names
    company_summary.columns = ['_'.join(col).strip() for col in company_summary.columns.values]
    company_summary.rename(columns={
        'ITNAME_count': 'Product_Count',
        'QTY_sum': 'Total_Quantity',
        'QTY_mean': 'Avg_Quantity',
        'QTY_std': 'Quantity_StdDev',
        'TAXBLEAMT_sum': 'Total_Revenue',
        'TAXBLEAMT_mean': 'Avg_Revenue',
        'TAXBLEAMT_std': 'Revenue_StdDev',
        'GST_AMOUNT_sum': 'Total_GST',
        'TOTAL_WITH_GST_sum': 'Total_With_GST',
        'PRICE_PER_UNIT_mean': 'Avg_Price_Per_Unit'
    }, inplace=True)

    # Calculate market share
    total_revenue = company_summary['Total_Revenue'].sum()
    company_summary['Market_Share_%'] = (company_summary['Total_Revenue'] / total_revenue * 100).round(2)

    # Sort by revenue
    company_summary = company_summary.sort_values('Total_Revenue', ascending=False)

    return company_summary

def generate_product_analysis(df):
    """Generate detailed product-level analysis"""
    product_summary = df.groupby(['ITNAME', 'COMPANY']).agg({
        'QTY': 'sum',
        'TAXBLEAMT': 'sum',
        'GST_AMOUNT': 'sum',
        'TOTAL_WITH_GST': 'sum',
        'PRICE_PER_UNIT': 'mean',
        'PERFORMANCE_SCORE': 'first'
    }).round(2)

    product_summary.rename(columns={
        'QTY': 'Total_Quantity',
        'TAXBLEAMT': 'Total_Revenue',
        'GST_AMOUNT': 'Total_GST',
        'TOTAL_WITH_GST': 'Total_With_GST',
        'PRICE_PER_UNIT': 'Avg_Price',
        'PERFORMANCE_SCORE': 'Performance_Score'
    }, inplace=True)

    product_summary = product_summary.sort_values('Total_Revenue', ascending=False)

    return product_summary

def generate_gst_analysis(df):
    """Analyze GST distribution and impact"""
    gst_summary = df.groupby('GST').agg({
        'ITNAME': 'count',
        'TAXBLEAMT': 'sum',
        'GST_AMOUNT': 'sum',
        'QTY': 'sum'
    }).round(2)

    gst_summary.rename(columns={
        'ITNAME': 'Product_Count',
        'TAXBLEAMT': 'Taxable_Amount',
        'GST_AMOUNT': 'GST_Collected',
        'QTY': 'Total_Quantity'
    }, inplace=True)

    total_taxable = gst_summary['Taxable_Amount'].sum()
    gst_summary['Contribution_%'] = (gst_summary['Taxable_Amount'] / total_taxable * 100).round(2)

    return gst_summary

def identify_star_performers(df, top_n=10):
    """Identify top and bottom performers"""
    performers = {
        'top_by_revenue': df.nlargest(top_n, 'TAXBLEAMT')[['ITNAME', 'COMPANY', 'TAXBLEAMT', 'QTY', 'PERFORMANCE_SCORE']],
        'top_by_quantity': df.nlargest(top_n, 'QTY')[['ITNAME', 'COMPANY', 'QTY', 'TAXBLEAMT', 'PERFORMANCE_SCORE']],
        'top_by_performance': df.nlargest(top_n, 'PERFORMANCE_SCORE')[['ITNAME', 'COMPANY', 'PERFORMANCE_SCORE', 'TAXBLEAMT', 'QTY']],
        'bottom_performers': df.nsmallest(top_n, 'PERFORMANCE_SCORE')[['ITNAME', 'COMPANY', 'PERFORMANCE_SCORE', 'TAXBLEAMT', 'QTY']],
        'highest_price_per_unit': df.nlargest(top_n, 'PRICE_PER_UNIT')[['ITNAME', 'COMPANY', 'PRICE_PER_UNIT', 'QTY']],
        'lowest_price_per_unit': df.nsmallest(top_n, 'PRICE_PER_UNIT')[['ITNAME', 'COMPANY', 'PRICE_PER_UNIT', 'QTY']]
    }
    return performers

# ============================================
# Enhanced Visualization Functions
# ============================================

def create_all_visualizations(df, company_summary, gst_summary, output_dir='outputs', report_type=''):
    """Create comprehensive set of visualizations"""

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    title_prefix = f"{report_type} - " if report_type else ""

    viz_files = []

    # 1. Top 15 Products by Revenue
    plt.figure(figsize=(14, 8))
    top_products = df.nlargest(15, 'TAXBLEAMT')
    colors = plt.cm.Spectral(np.linspace(0, 1, len(top_products)))
    bars = plt.barh(range(len(top_products)), top_products['TAXBLEAMT'], color=colors)
    plt.yticks(range(len(top_products)), top_products['ITNAME'])
    plt.xlabel('Revenue (Rs.)', fontsize=12, fontweight='bold')
    plt.title(f'{title_prefix}Top 15 Products by Revenue', fontsize=14, fontweight='bold')
    plt.gca().invert_yaxis()
    plt.tight_layout()
    filename = f'{output_dir}/01_top_products_revenue.png'
    plt.savefig(filename, dpi=300, bbox_inches='tight')
    viz_files.append(filename)
    plt.close()

    # 2. Top 15 Products by Quantity
    plt.figure(figsize=(14, 8))
    top_qty = df.nlargest(15, 'QTY')
    colors = plt.cm.viridis(np.linspace(0, 1, len(top_qty)))
    plt.barh(range(len(top_qty)), top_qty['QTY'], color=colors)
    plt.yticks(range(len(top_qty)), top_qty['ITNAME'])
    plt.xlabel('Quantity Sold', fontsize=12, fontweight='bold')
    plt.title(f'{title_prefix}Top 15 Products by Quantity', fontsize=14, fontweight='bold')
    plt.gca().invert_yaxis()
    plt.tight_layout()
    filename = f'{output_dir}/02_top_products_quantity.png'
    plt.savefig(filename, dpi=300, bbox_inches='tight')
    viz_files.append(filename)
    plt.close()

    # 3. Company Revenue Distribution
    plt.figure(figsize=(14, 8))
    top_companies = company_summary.nlargest(15, 'Total_Revenue')
    colors = plt.cm.Set3(np.linspace(0, 1, len(top_companies)))
    plt.barh(range(len(top_companies)), top_companies['Total_Revenue'], color=colors)
    plt.yticks(range(len(top_companies)), top_companies.index)
    plt.xlabel('Total Revenue (Rs.)', fontsize=12, fontweight='bold')
    plt.title(f'{title_prefix}Top 15 Companies by Revenue', fontsize=14, fontweight='bold')
    plt.gca().invert_yaxis()

    # Add value labels
    for i, v in enumerate(top_companies['Total_Revenue']):
        plt.text(v, i, f' Rs.{v:,.0f}', va='center', fontsize=10, fontweight='bold')

    plt.tight_layout()
    filename = f'{output_dir}/03_company_revenue.png'
    plt.savefig(filename, dpi=300, bbox_inches='tight')
    viz_files.append(filename)
    plt.close()

    print(f"\n✅ Generated {len(viz_files)} visualizations in '{output_dir}/' directory")

    return viz_files

def generate_html_report(df, company_summary, product_summary, gst_summary, performers, viz_files):
    """Generate comprehensive HTML report"""

    html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sales Analysis Report</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: #333;
        }}
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            padding: 40px;
            border-radius: 15px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
        }}
        h1 {{
            color: #667eea;
            text-align: center;
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
        }}
        .subtitle {{
            text-align: center;
            color: #666;
            font-size: 1.2em;
            margin-bottom: 40px;
        }}
        h2 {{
            color: #764ba2;
            border-bottom: 3px solid #667eea;
            padding-bottom: 10px;
            margin-top: 40px;
        }}
        h3 {{
            color: #667eea;
            margin-top: 30px;
        }}
        .metrics-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin: 30px 0;
        }}
        .metric-card {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 25px;
            border-radius: 10px;
            color: white;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
            transition: transform 0.3s;
        }}
        .metric-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 8px 20px rgba(0,0,0,0.2);
        }}
        .metric-label {{
            font-size: 0.9em;
            opacity: 0.9;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        .metric-value {{
            font-size: 2em;
            font-weight: bold;
            margin-top: 10px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }}
        th {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: bold;
        }}
        td {{
            padding: 12px 15px;
            border-bottom: 1px solid #ddd;
        }}
        tr:hover {{
            background-color: #f5f5f5;
        }}
        tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        .viz-container {{
            margin: 30px 0;
            text-align: center;
        }}
        .viz-container img {{
            max-width: 100%;
            height: auto;
            border-radius: 10px;
            box-shadow: 0 5px 20px rgba(0,0,0,0.15);
            margin: 20px 0;
        }}
        .insight-box {{
            background: #f0f4ff;
            border-left: 5px solid #667eea;
            padding: 20px;
            margin: 20px 0;
            border-radius: 5px;
        }}
        .footer {{
            text-align: center;
            margin-top: 50px;
            padding-top: 20px;
            border-top: 2px solid #eee;
            color: #666;
        }}
        .number-positive {{
            color: #28a745;
            font-weight: bold;
        }}
        .number-negative {{
            color: #dc3545;
            font-weight: bold;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Sales Analysis Report</h1>
        <div class="subtitle">Generated on {datetime.now().strftime('%B %d, %Y at %H:%M:%S')}</div>

        <h2>📈 Executive Summary</h2>
        <div class="metrics-grid">
            <div class="metric-card">
                <div class="metric-label">Total Products</div>
                <div class="metric-value">{len(df):,}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Quantity Sold</div>
                <div class="metric-value">{int(df['QTY'].sum()):,}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Revenue</div>
                <div class="metric-value">Rs.{df['TAXBLEAMT'].sum():,.0f}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total GST</div>
                <div class="metric-value">Rs.{df['GST_AMOUNT'].sum():,.0f}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Grand Total</div>
                <div class="metric-value">Rs.{df['TOTAL_WITH_GST'].sum():,.0f}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Companies</div>
                <div class="metric-value">{df['COMPANY'].nunique()}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Avg Price/Unit</div>
                <div class="metric-value">Rs.{df['PRICE_PER_UNIT'].mean():,.2f}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Avg Order Value</div>
                <div class="metric-value">Rs.{df['TAXBLEAMT'].mean():,.0f}</div>
            </div>
        </div>

        <h2>🏆 Top Performers</h2>

        <h3>Top 10 Products by Revenue</h3>
        <table>
            <tr>
                <th>Rank</th>
                <th>Product</th>
                <th>Company</th>
                <th>Revenue</th>
                <th>Quantity</th>
                <th>Performance Score</th>
            </tr>
"""

    for idx, row in enumerate(performers['top_by_revenue'].head(10).itertuples(), 1):
        html_content += f"""
            <tr>
                <td>{idx}</td>
                <td>{row.ITNAME}</td>
                <td>{row.COMPANY}</td>
                <td class="number-positive">Rs.{row.TAXBLEAMT:,.2f}</td>
                <td>{int(row.QTY):,}</td>
                <td>{row.PERFORMANCE_SCORE:.2f}</td>
            </tr>
"""

    html_content += """
        </table>

        <h3>Top 10 Products by Quantity</h3>
        <table>
            <tr>
                <th>Rank</th>
                <th>Product</th>
                <th>Company</th>
                <th>Quantity</th>
                <th>Revenue</th>
                <th>Performance Score</th>
            </tr>
"""

    for idx, row in enumerate(performers['top_by_quantity'].head(10).itertuples(), 1):
        html_content += f"""
            <tr>
                <td>{idx}</td>
                <td>{row.ITNAME}</td>
                <td>{row.COMPANY}</td>
                <td class="number-positive">{int(row.QTY):,}</td>
                <td>Rs.{row.TAXBLEAMT:,.2f}</td>
                <td>{row.PERFORMANCE_SCORE:.2f}</td>
            </tr>
"""

    html_content += f"""
        </table>

        <h2>🏢 Company Analysis</h2>
        <table>
            <tr>
                <th>Company</th>
                <th>Products</th>
                <th>Total Revenue</th>
                <th>Market Share</th>
                <th>Avg Price/Unit</th>
                <th>Total GST</th>
            </tr>
"""

    for company, row in company_summary.head(15).iterrows():
        html_content += f"""
            <tr>
                <td><strong>{company}</strong></td>
                <td>{int(row['Product_Count'])}</td>
                <td class="number-positive">Rs.{row['Total_Revenue']:,.2f}</td>
                <td>{row['Market_Share_%']:.2f}%</td>
                <td>Rs.{row['Avg_Price_Per_Unit']:,.2f}</td>
                <td>Rs.{row['Total_GST']:,.2f}</td>
            </tr>
"""

    html_content += f"""
        </table>

        <h2>💰 GST Analysis</h2>
        <table>
            <tr>
                <th>GST Rate (%)</th>
                <th>Product Count</th>
                <th>Taxable Amount</th>
                <th>GST Collected</th>
                <th>Contribution</th>
            </tr>
"""

    for gst_rate, row in gst_summary.iterrows():
        html_content += f"""
            <tr>
                <td><strong>{gst_rate}%</strong></td>
                <td>{int(row['Product_Count'])}</td>
                <td>Rs.{row['Taxable_Amount']:,.2f}</td>
                <td class="number-positive">Rs.{row['GST_Collected']:,.2f}</td>
                <td>{row['Contribution_%']:.2f}%</td>
            </tr>
"""

    html_content += """
        </table>

        <h2>📊 Visualizations</h2>
"""

    # Add all visualizations
    for viz_file in viz_files:
        viz_name = os.path.basename(viz_file).replace('.png', '').replace('_', ' ').title()
        html_content += f"""
        <div class="viz-container">
            <h3>{viz_name}</h3>
            <img src="{viz_file}" alt="{viz_name}">
        </div>
"""

    # Add Key Insights
    html_content += f"""
        <h2>💡 Key Insights</h2>

        <div class="insight-box">
            <h3>Business Performance</h3>
            <ul>
                <li>Total revenue generated: <strong>Rs.{df['TAXBLEAMT'].sum():,.0f}</strong></li>
                <li>Grand total (with GST): <strong>Rs.{df['TOTAL_WITH_GST'].sum():,.0f}</strong></li>
                <li>Total GST collected: <strong>Rs.{df['GST_AMOUNT'].sum():,.0f}</strong></li>
                <li>Average transaction value: <strong>Rs.{df['TAXBLEAMT'].mean():,.2f}</strong></li>
                <li>Total units sold: <strong>{int(df['QTY'].sum()):,}</strong></li>
            </ul>
        </div>

        <div class="insight-box">
            <h3>Top Performer</h3>
            <ul>
                <li>Best-selling product by revenue: <strong>{performers['top_by_revenue'].iloc[0]['ITNAME']}</strong> (Rs.{performers['top_by_revenue'].iloc[0]['TAXBLEAMT']:,.2f})</li>
                <li>Best-selling product by quantity: <strong>{performers['top_by_quantity'].iloc[0]['ITNAME']}</strong> ({int(performers['top_by_quantity'].iloc[0]['QTY']):,} units)</li>
                <li>Highest performance score: <strong>{performers['top_by_performance'].iloc[0]['ITNAME']}</strong> (Score: {performers['top_by_performance'].iloc[0]['PERFORMANCE_SCORE']:.2f})</li>
            </ul>
        </div>

        <div class="insight-box">
            <h3>Company Insights</h3>
            <ul>
                <li>Top company by revenue: <strong>{company_summary.index[0]}</strong> (Rs.{company_summary.iloc[0]['Total_Revenue']:,.2f})</li>
                <li>Market leader's share: <strong>{company_summary.iloc[0]['Market_Share_%']:.2f}%</strong></li>
                <li>Company with most products: <strong>{company_summary.nlargest(1, 'Product_Count').index[0]}</strong> ({int(company_summary.nlargest(1, 'Product_Count').iloc[0]['Product_Count'])} products)</li>
                <li>Total companies: <strong>{df['COMPANY'].nunique()}</strong></li>
            </ul>
        </div>

        <div class="insight-box">
            <h3>Pricing Analysis</h3>
            <ul>
                <li>Average price per unit: <strong>Rs.{df['PRICE_PER_UNIT'].mean():,.2f}</strong></li>
                <li>Highest priced product: <strong>{performers['highest_price_per_unit'].iloc[0]['ITNAME']}</strong> (Rs.{performers['highest_price_per_unit'].iloc[0]['PRICE_PER_UNIT']:,.2f}/unit)</li>
                <li>Lowest priced product: <strong>{performers['lowest_price_per_unit'].iloc[0]['ITNAME']}</strong> (Rs.{performers['lowest_price_per_unit'].iloc[0]['PRICE_PER_UNIT']:,.2f}/unit)</li>
                <li>Price range: Rs.{df['PRICE_PER_UNIT'].min():,.2f} - Rs.{df['PRICE_PER_UNIT'].max():,.2f}</li>
            </ul>
        </div>

        <div class="footer">
            <p>Report generated by Enhanced Sales Analyzer</p>
            <p>Generated on {datetime.now().strftime('%B %d, %Y at %H:%M:%S')}</p>
        </div>
    </div>
</body>
</html>
"""

    report_filename = 'outputs/Sales_Analysis_Report.html'
    with open(report_filename, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"\n✅ HTML report generated: {report_filename}")
    return report_filename

def generate_pdf_report(df, company_summary, product_summary, gst_summary, performers, viz_files, output_dir='outputs', report_type=''):
    """Generate comprehensive PDF report with all visualizations and metrics"""

    pdf_filename = f'{output_dir}/Sales_Analysis_Report.pdf'
    report_title = f'{report_type} Sales Analysis Report' if report_type else 'Sales Analysis Report'

    with PdfPages(pdf_filename) as pdf:
        # Page 1: Cover Page with Executive Summary
        fig = plt.figure(figsize=(11, 8.5))
        fig.patch.set_facecolor('white')
        ax = fig.add_subplot(111)
        ax.axis('off')

        # Title
        fig.text(0.5, 0.95, report_title, ha='center', fontsize=28,
                fontweight='bold', color='#667eea')
        fig.text(0.5, 0.91, f'Generated on {datetime.now().strftime("%B %d, %Y at %H:%M:%S")}',
                ha='center', fontsize=10, color='#888')

        # Executive Summary Title
        fig.text(0.08, 0.85, '📊 Executive Summary', ha='left', fontsize=18,
                fontweight='bold', color='#667eea')

        # Define metric cards data with more space below title
        metrics = [
            ('TOTAL PRODUCTS', f"{len(df):,}", 0.08, 0.69),
            ('TOTAL QUANTITY SOLD', f"{int(df['QTY'].sum()):,}", 0.29, 0.69),
            ('TOTAL REVENUE', f"Rs.{df['TAXBLEAMT'].sum():,.0f}", 0.50, 0.69),
            ('TOTAL GST', f"Rs.{df['GST_AMOUNT'].sum():,.0f}", 0.71, 0.69),
            ('GRAND TOTAL', f"Rs.{df['TOTAL_WITH_GST'].sum():,.0f}", 0.08, 0.54),
            ('TOTAL COMPANIES', f"{df['COMPANY'].nunique()}", 0.29, 0.54),
        ]

        # Create clean gradient colored cards
        card_width = 0.19
        card_height = 0.12

        # Single fusion gradient from light purple to light blue
        for idx, (label, value, x, y) in enumerate(metrics):
            # Subtle gradient - mostly blue with slight variations
            if idx < 4:
                card_color = '#7B8FF7'  # Slightly purple-blue
            else:
                card_color = '#6BA3F9'  # Slightly lighter blue

            # Draw card with subtle rounded corners
            fancy_box = FancyBboxPatch((x, y), card_width, card_height,
                                      boxstyle="round,pad=0.005",
                                      facecolor=card_color, edgecolor='none',
                                      transform=fig.transFigure, zorder=2,
                                      alpha=0.95)
            fig.patches.append(fancy_box)

            # Add label
            fig.text(x + card_width/2, y + card_height - 0.02, label,
                    ha='center', va='top', fontsize=8.5, color='white',
                    fontweight='normal', alpha=0.95)

            # Add value
            fig.text(x + card_width/2, y + 0.035, value,
                    ha='center', va='bottom', fontsize=16, color='white',
                    fontweight='bold')

        # Top Performers Section
        fig.text(0.08, 0.46, '🏆 Top Performers', ha='left', fontsize=20,
                fontweight='bold', color='#667eea')

        performers_text = f"""Best Product by Revenue:
  • {performers['top_by_revenue'].iloc[0]['ITNAME']}
  • Revenue: Rs.{performers['top_by_revenue'].iloc[0]['TAXBLEAMT']:,.2f}
  • Company: {performers['top_by_revenue'].iloc[0]['COMPANY']}

Best Product by Quantity:
  • {performers['top_by_quantity'].iloc[0]['ITNAME']}
  • Quantity: {int(performers['top_by_quantity'].iloc[0]['QTY']):,} units
  • Company: {performers['top_by_quantity'].iloc[0]['COMPANY']}

Top Company by Revenue:
  • {company_summary.index[0]}
  • Revenue: Rs.{company_summary.iloc[0]['Total_Revenue']:,.2f}
  • Market Share: {company_summary.iloc[0]['Market_Share_%']:.2f}%"""

        fig.text(0.5, 0.25, performers_text, ha='center', va='center', fontsize=10,
                family='monospace', bbox=dict(boxstyle='round', facecolor='#f8f9ff',
                edgecolor='#667eea', linewidth=1.5, pad=1))

        pdf.savefig(fig, bbox_inches='tight')
        plt.close()

        # Page 2: Company Analysis Summary
        fig = plt.figure(figsize=(11, 8.5))
        fig.patch.set_facecolor('white')
        ax = fig.add_subplot(111)
        ax.axis('off')

        fig.text(0.5, 0.95, 'Company Analysis Summary', ha='center', fontsize=20,
                fontweight='bold', color='#667eea')

        # Create table data - show all companies
        all_companies = company_summary
        table_data = []
        table_data.append(['Rank', 'Company', 'Revenue', 'Products'])

        for idx, (company, row) in enumerate(all_companies.iterrows(), 1):
            table_data.append([
                str(idx),
                company[:25],  # Truncate long names
                f"Rs.{row['Total_Revenue']:,.0f}",
                str(int(row['Product_Count']))
            ])

        # Create table
        table = ax.table(cellText=table_data, cellLoc='left',
                        loc='center', bbox=[0.1, 0.05, 0.8, 0.85])
        table.auto_set_font_size(False)
        table.set_fontsize(8)
        table.scale(1, 1.5)

        # Style header row
        for i in range(4):
            cell = table[(0, i)]
            cell.set_facecolor('#667eea')
            cell.set_text_props(weight='bold', color='white')

        # Alternate row colors
        for i in range(1, len(table_data)):
            for j in range(4):
                cell = table[(i, j)]
                if i % 2 == 0:
                    cell.set_facecolor('#f9f9f9')

        pdf.savefig(fig, bbox_inches='tight')
        plt.close()

        # Page 3: Top Products by Revenue
        fig = plt.figure(figsize=(11, 8.5))
        fig.patch.set_facecolor('white')
        ax = fig.add_subplot(111)
        ax.axis('off')

        fig.text(0.5, 0.95, 'Top Products by Revenue', ha='center', fontsize=18,
                fontweight='bold', color='#667eea')

        # Top 10 products table
        top_products = performers['top_by_revenue'].head(10)
        table_data = []
        table_data.append(['Rank', 'Product', 'Company', 'Revenue', 'Qty'])

        for idx, row in enumerate(top_products.itertuples(), 1):
            table_data.append([
                str(idx),
                row.ITNAME[:25],
                row.COMPANY[:15],
                f"Rs.{row.TAXBLEAMT:,.0f}",
                str(int(row.QTY))
            ])

        table = ax.table(cellText=table_data, cellLoc='left',
                        loc='upper center', bbox=[0.05, 0.45, 0.9, 0.45])
        table.auto_set_font_size(False)
        table.set_fontsize(9)
        table.scale(1, 1.8)

        # Style header
        for i in range(5):
            cell = table[(0, i)]
            cell.set_facecolor('#667eea')
            cell.set_text_props(weight='bold', color='white')

        # Alternate rows
        for i in range(1, len(table_data)):
            for j in range(5):
                cell = table[(i, j)]
                if i % 2 == 0:
                    cell.set_facecolor('#f9f9f9')

        pdf.savefig(fig, bbox_inches='tight')
        plt.close()

        # Page 4+: Add all visualization images
        for viz_file in viz_files:
            fig = plt.figure(figsize=(11, 8.5))
            img = plt.imread(viz_file)
            plt.imshow(img)
            plt.axis('off')
            plt.tight_layout(pad=0)
            pdf.savefig(fig, bbox_inches='tight')
            plt.close()

        # Set PDF metadata
        d = pdf.infodict()
        d['Title'] = report_title
        d['Author'] = 'Enhanced Sales Analyzer'
        d['Subject'] = 'Sales Data Analysis'
        d['Keywords'] = 'Sales, Analysis, Report'
        d['CreationDate'] = datetime.now()

    print(f"\n✅ PDF report generated: {pdf_filename}")
    return pdf_filename

# ============================================
# Main Analysis Function
# ============================================

def process_file(filename, output_dir, report_type):
    """Process a single data file and generate reports"""

    try:
        print(f"\n{'='*60}")
        print(f"📊 Processing {report_type} Data: {filename}")
        print(f"{'='*60}")

        # Load data
        df = load_data(filename)

        # Add company categorization
        print("\n🏷️  Categorizing products by company...")
        df = categorize_by_company(df)

        # Calculate advanced metrics
        print("📊 Calculating advanced metrics...")
        df = calculate_advanced_metrics(df)

        # Generate analyses
        print("🔍 Generating company analysis...")
        company_summary = generate_company_analysis(df)

        print("🔍 Generating product analysis...")
        product_summary = generate_product_analysis(df)

        print("🔍 Generating GST analysis...")
        gst_summary = generate_gst_analysis(df)

        print("🔍 Identifying star performers...")
        performers = identify_star_performers(df, top_n=10)

        # Create visualizations
        print("\n📈 Creating visualizations...")
        viz_files = create_all_visualizations(df, company_summary, gst_summary,
                                              output_dir=output_dir, report_type=report_type)

        # Generate PDF report
        print("\n📄 Generating comprehensive PDF report...")
        report_file = generate_pdf_report(df, company_summary, product_summary,
                                         gst_summary, performers, viz_files,
                                         output_dir=output_dir, report_type=report_type)

        # Save summary data to Excel
        print("\n💾 Saving analysis data to Excel...")
        with pd.ExcelWriter(f'{output_dir}/Analysis_Summary.xlsx', engine='openpyxl') as writer:
            company_summary.to_excel(writer, sheet_name='Company Analysis')
            product_summary.to_excel(writer, sheet_name='Product Analysis')
            gst_summary.to_excel(writer, sheet_name='GST Analysis')
            performers['top_by_revenue'].to_excel(writer, sheet_name='Top Revenue', index=False)
            performers['top_by_quantity'].to_excel(writer, sheet_name='Top Quantity', index=False)
            performers['top_by_performance'].to_excel(writer, sheet_name='Top Performance', index=False)

        print(f"\n{'='*60}")
        print(f"✅ {report_type} ANALYSIS COMPLETE!")
        print(f"{'='*60}")
        print(f"\n📊 Results Summary:")
        print(f"   • PDF Report: {output_dir}/Sales_Analysis_Report.pdf")
        print(f"   • Excel Summary: {output_dir}/Analysis_Summary.xlsx")
        print(f"   • Visualizations: {len(viz_files)} charts in {output_dir}/ folder")

        return True

    except Exception as e:
        print(f"\n❌ Error during {report_type} analysis: {str(e)}")
        import traceback
        print("\n📋 Full error trace:")
        print(traceback.format_exc())
        return False

def main():
    """Main execution function"""

    print("="*60)
    print("🚀 ENHANCED SALES DATA ANALYZER")
    print("="*60)

    # Look for specific data files
    pesticide_file = 'SALANAL_PS.XLS'
    fertilizer_file = 'SALANAL_FS.XLS'

    files_to_process = []

    if os.path.exists(pesticide_file):
        files_to_process.append({
            'filename': pesticide_file,
            'output_dir': 'outputs_pesticides',
            'report_type': 'Pesticides'
        })

    if os.path.exists(fertilizer_file):
        files_to_process.append({
            'filename': fertilizer_file,
            'output_dir': 'outputs_fertilizers',
            'report_type': 'Fertilizers'
        })

    if not files_to_process:
        print("\n❌ No data files found!")
        print("📝 Looking for: SALANAL_PS.XLS (Pesticides) and/or SALANAL_FS.XLS (Fertilizers)")
        return

    print(f"\n📁 Found {len(files_to_process)} data file(s) to process:")
    for item in files_to_process:
        print(f"   • {item['filename']} → {item['report_type']}")

    # Process each file
    results = []
    for item in files_to_process:
        success = process_file(item['filename'], item['output_dir'], item['report_type'])
        results.append({
            'type': item['report_type'],
            'success': success,
            'output_dir': item['output_dir']
        })

    # Final summary
    print("\n" + "="*60)
    print("🎉 ALL ANALYSIS COMPLETE!")
    print("="*60)
    print("\n📊 Final Summary:")
    for result in results:
        if result['success']:
            print(f"\n✅ {result['type']}:")
            print(f"   • PDF: {result['output_dir']}/Sales_Analysis_Report.pdf")
            print(f"   • Excel: {result['output_dir']}/Analysis_Summary.xlsx")
        else:
            print(f"\n❌ {result['type']}: Failed to generate reports")

    print("\n💡 Open the PDF reports to view detailed analysis!")
    print("="*60)

if __name__ == "__main__":
    main()
