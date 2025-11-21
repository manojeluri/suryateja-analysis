"""
Vercel Serverless Function for Sales Data Analysis
Receives JSON data from n8n and returns PDF report as base64
"""

import json
import base64
import io
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')
import os
import sys
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import FancyBboxPatch

# Add parent directory to path to import helper modules
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

# Set style for better visualizations
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (12, 6)
plt.rcParams['font.size'] = 10

# Load company products mapping
def load_company_products_from_csv(folder_path="Company Wise Products"):
    """Load company-product mappings from CSV files"""
    company_products = {}

    base_path = os.path.dirname(os.path.dirname(__file__))
    full_path = os.path.join(base_path, folder_path)

    if not os.path.exists(full_path):
        return company_products

    csv_files = [f for f in os.listdir(full_path) if f.lower().endswith('.csv')]

    for csv_file in csv_files:
        company_name = csv_file.replace('_Products.csv', '').replace('_Product_Names.csv', '')
        company_name = company_name.replace('_', ' ')

        if company_name == 'BestAgrolife':
            company_name = 'Best Agrolife'
        elif company_name == 'NovaAgriScience':
            company_name = 'Nova Agri Science'

        file_path = os.path.join(full_path, csv_file)
        try:
            df = pd.read_csv(file_path)
            if 'Product Name' in df.columns:
                products = df['Product Name'].dropna().str.strip().tolist()
            else:
                products = df.iloc[:, 0].dropna().str.strip().tolist()
            products = [p for p in products if p]
            company_products[company_name] = products
        except Exception as e:
            print(f"Error reading {csv_file}: {str(e)}")

    return company_products

COMPANY_PRODUCTS = load_company_products_from_csv()
PRODUCT_TO_COMPANY = {}
for company, products in COMPANY_PRODUCTS.items():
    for product in products:
        PRODUCT_TO_COMPANY[product] = company

def categorize_by_company(df):
    """Add company categorization to dataframe"""
    df['COMPANY'] = df['ITNAME'].map(PRODUCT_TO_COMPANY).fillna('Other')
    return df

def calculate_advanced_metrics(df):
    """Calculate advanced business metrics"""
    df['PRICE_PER_UNIT'] = df['TAXBLEAMT'] / df['QTY']
    df['GST_AMOUNT'] = df['TAXBLEAMT'] * df['GST'] / 100
    df['TOTAL_WITH_GST'] = df['TAXBLEAMT'] + df['GST_AMOUNT']
    df['REVENUE_PERCENTILE'] = df['TAXBLEAMT'].rank(pct=True) * 100
    df['QTY_PERCENTILE'] = df['QTY'].rank(pct=True) * 100

    revenue_norm = (df['TAXBLEAMT'] - df['TAXBLEAMT'].min()) / (df['TAXBLEAMT'].max() - df['TAXBLEAMT'].min())
    qty_norm = (df['QTY'] - df['QTY'].min()) / (df['QTY'].max() - df['QTY'].min())
    df['PERFORMANCE_SCORE'] = (revenue_norm * 0.6 + qty_norm * 0.4) * 100

    return df

def generate_company_analysis(df):
    """Generate comprehensive company-wise analysis"""
    company_summary = df.groupby('COMPANY').agg({
        'ITNAME': 'count',
        'QTY': ['sum', 'mean'],
        'TAXBLEAMT': ['sum', 'mean'],
        'GST_AMOUNT': 'sum',
        'TOTAL_WITH_GST': 'sum',
        'PRICE_PER_UNIT': 'mean'
    }).round(2)

    company_summary.columns = ['_'.join(col).strip() for col in company_summary.columns.values]
    company_summary.rename(columns={
        'ITNAME_count': 'Product_Count',
        'QTY_sum': 'Total_Quantity',
        'QTY_mean': 'Avg_Quantity',
        'TAXBLEAMT_sum': 'Total_Revenue',
        'TAXBLEAMT_mean': 'Avg_Revenue',
        'GST_AMOUNT_sum': 'Total_GST',
        'TOTAL_WITH_GST_sum': 'Total_With_GST',
        'PRICE_PER_UNIT_mean': 'Avg_Price_Per_Unit'
    }, inplace=True)

    total_revenue = company_summary['Total_Revenue'].sum()
    company_summary['Market_Share_%'] = (company_summary['Total_Revenue'] / total_revenue * 100).round(2)
    company_summary = company_summary.sort_values('Total_Revenue', ascending=False)

    return company_summary

def generate_pdf_report(df, company_summary):
    """Generate comprehensive PDF report"""

    pdf_buffer = io.BytesIO()

    with PdfPages(pdf_buffer) as pdf:
        # Page 1: Executive Summary
        fig = plt.figure(figsize=(11, 8.5))
        fig.patch.set_facecolor('white')
        ax = fig.add_subplot(111)
        ax.axis('off')

        # Title
        fig.text(0.5, 0.95, 'Sales Analysis Report', ha='center', fontsize=28,
                fontweight='bold', color='#667eea')
        fig.text(0.5, 0.91, f'Generated on {datetime.now().strftime("%B %d, %Y at %H:%M:%S")}',
                ha='center', fontsize=10, color='#888')

        # Executive Summary Title
        fig.text(0.08, 0.85, '📊 Executive Summary', ha='left', fontsize=18,
                fontweight='bold', color='#667eea')

        # Metric cards
        metrics = [
            ('TOTAL PRODUCTS', f"{len(df):,}", 0.08, 0.69),
            ('TOTAL QUANTITY', f"{int(df['QTY'].sum()):,}", 0.29, 0.69),
            ('TOTAL REVENUE', f"Rs.{df['TAXBLEAMT'].sum():,.0f}", 0.50, 0.69),
            ('TOTAL GST', f"Rs.{df['GST_AMOUNT'].sum():,.0f}", 0.71, 0.69),
            ('GRAND TOTAL', f"Rs.{df['TOTAL_WITH_GST'].sum():,.0f}", 0.08, 0.54),
            ('COMPANIES', f"{df['COMPANY'].nunique()}", 0.29, 0.54),
        ]

        card_width = 0.19
        card_height = 0.12

        for idx, (label, value, x, y) in enumerate(metrics):
            card_color = '#7B8FF7' if idx < 4 else '#6BA3F9'

            fancy_box = FancyBboxPatch((x, y), card_width, card_height,
                                      boxstyle="round,pad=0.005",
                                      facecolor=card_color, edgecolor='none',
                                      transform=fig.transFigure, zorder=2,
                                      alpha=0.95)
            fig.patches.append(fancy_box)

            fig.text(x + card_width/2, y + card_height - 0.02, label,
                    ha='center', va='top', fontsize=8.5, color='white',
                    fontweight='normal', alpha=0.95)

            fig.text(x + card_width/2, y + 0.035, value,
                    ha='center', va='bottom', fontsize=16, color='white',
                    fontweight='bold')

        # Top Performers
        fig.text(0.08, 0.46, '🏆 Top Performers', ha='left', fontsize=20,
                fontweight='bold', color='#667eea')

        top_product = df.nlargest(1, 'TAXBLEAMT').iloc[0]
        top_qty_product = df.nlargest(1, 'QTY').iloc[0]

        performers_text = f"""Best Product by Revenue:
  • {top_product['ITNAME']}
  • Revenue: Rs.{top_product['TAXBLEAMT']:,.2f}
  • Company: {top_product['COMPANY']}

Best Product by Quantity:
  • {top_qty_product['ITNAME']}
  • Quantity: {int(top_qty_product['QTY']):,} units
  • Company: {top_qty_product['COMPANY']}

Top Company by Revenue:
  • {company_summary.index[0]}
  • Revenue: Rs.{company_summary.iloc[0]['Total_Revenue']:,.2f}
  • Market Share: {company_summary.iloc[0]['Market_Share_%']:.2f}%"""

        fig.text(0.5, 0.25, performers_text, ha='center', va='center', fontsize=10,
                family='monospace', bbox=dict(boxstyle='round', facecolor='#f8f9ff',
                edgecolor='#667eea', linewidth=1.5, pad=1))

        pdf.savefig(fig, bbox_inches='tight')
        plt.close()

        # Page 2: Top Products by Revenue Chart
        fig = plt.figure(figsize=(11, 8.5))
        top_products = df.nlargest(15, 'TAXBLEAMT')
        colors = plt.cm.Spectral(np.linspace(0, 1, len(top_products)))
        plt.barh(range(len(top_products)), top_products['TAXBLEAMT'], color=colors)
        plt.yticks(range(len(top_products)), top_products['ITNAME'])
        plt.xlabel('Revenue (Rs.)', fontsize=12, fontweight='bold')
        plt.title('Top 15 Products by Revenue', fontsize=16, fontweight='bold', pad=20)
        plt.gca().invert_yaxis()
        plt.tight_layout()
        pdf.savefig(fig, bbox_inches='tight')
        plt.close()

        # Page 3: Top Products by Quantity Chart
        fig = plt.figure(figsize=(11, 8.5))
        top_qty = df.nlargest(15, 'QTY')
        colors = plt.cm.viridis(np.linspace(0, 1, len(top_qty)))
        plt.barh(range(len(top_qty)), top_qty['QTY'], color=colors)
        plt.yticks(range(len(top_qty)), top_qty['ITNAME'])
        plt.xlabel('Quantity Sold', fontsize=12, fontweight='bold')
        plt.title('Top 15 Products by Quantity', fontsize=16, fontweight='bold', pad=20)
        plt.gca().invert_yaxis()
        plt.tight_layout()
        pdf.savefig(fig, bbox_inches='tight')
        plt.close()

        # Page 4: Company Revenue Chart
        fig = plt.figure(figsize=(11, 8.5))
        top_companies = company_summary.nlargest(15, 'Total_Revenue')
        colors = plt.cm.Set3(np.linspace(0, 1, len(top_companies)))
        plt.barh(range(len(top_companies)), top_companies['Total_Revenue'], color=colors)
        plt.yticks(range(len(top_companies)), top_companies.index)
        plt.xlabel('Total Revenue (Rs.)', fontsize=12, fontweight='bold')
        plt.title('Top 15 Companies by Revenue', fontsize=16, fontweight='bold', pad=20)
        plt.gca().invert_yaxis()

        for i, v in enumerate(top_companies['Total_Revenue']):
            plt.text(v, i, f' Rs.{v:,.0f}', va='center', fontsize=10, fontweight='bold')

        plt.tight_layout()
        pdf.savefig(fig, bbox_inches='tight')
        plt.close()

        # Set PDF metadata
        d = pdf.infodict()
        d['Title'] = 'Sales Analysis Report'
        d['Author'] = 'Sales Analyzer'
        d['Subject'] = 'Sales Data Analysis'
        d['CreationDate'] = datetime.now()

    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()

def analyze_data(json_data):
    """Main analysis function that processes JSON data"""
    try:
        # Convert JSON to DataFrame
        df = pd.DataFrame(json_data)

        # Map column names if needed
        column_mapping = {
            'NAMT': 'TAXBLEAMT',
            'PER': 'GST'
        }
        df.rename(columns=column_mapping, inplace=True)

        # Ensure required columns exist
        required_cols = ['ITNAME', 'QTY', 'TAXBLEAMT', 'GST']
        missing_cols = [col for col in required_cols if col not in df.columns]

        if missing_cols:
            return {
                'success': False,
                'error': f'Missing required columns: {", ".join(missing_cols)}',
                'available_columns': df.columns.tolist()
            }

        # Categorize by company
        df = categorize_by_company(df)

        # Calculate metrics
        df = calculate_advanced_metrics(df)

        # Generate analyses
        company_summary = generate_company_analysis(df)

        # Generate PDF
        pdf_bytes = generate_pdf_report(df, company_summary)

        # Return PDF bytes and metadata
        return {
            'success': True,
            'pdf_bytes': pdf_bytes,
            'filename': f'Sales_Analysis_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf',
            'summary': {
                'total_rows': len(df),
                'total_revenue': float(df['TAXBLEAMT'].sum()),
                'total_quantity': int(df['QTY'].sum()),
                'total_gst': float(df['GST_AMOUNT'].sum()),
                'grand_total': float(df['TOTAL_WITH_GST'].sum()),
                'total_companies': int(df['COMPANY'].nunique())
            }
        }

    except Exception as e:
        import traceback
        return {
            'success': False,
            'error': str(e),
            'traceback': traceback.format_exc()
        }

# Vercel serverless handler
def handler(request):
    """Main handler for Vercel serverless function"""

    # Handle CORS
    if request.method == 'OPTIONS':
        return {
            'statusCode': 200,
            'headers': {
                'Access-Control-Allow-Origin': '*',
                'Access-Control-Allow-Methods': 'POST, OPTIONS',
                'Access-Control-Allow-Headers': 'Content-Type',
            },
            'body': ''
        }

    if request.method != 'POST':
        return {
            'statusCode': 405,
            'headers': {'Content-Type': 'application/json'},
            'body': json.dumps({'error': 'Method not allowed'})
        }

    try:
        # Parse request body
        if isinstance(request.body, bytes):
            body = json.loads(request.body.decode('utf-8'))
        elif isinstance(request.body, str):
            body = json.loads(request.body)
        else:
            body = request.body

        # Get data from request - handle different formats
        data = None

        # Try to get data from different possible locations
        if 'data' in body:
            data = body['data']
        elif isinstance(body, list):
            data = body

        if not data:
            return {
                'statusCode': 400,
                'headers': {'Content-Type': 'application/json'},
                'body': json.dumps({
                    'success': False,
                    'error': 'No data provided in request body',
                    'received_keys': list(body.keys()) if isinstance(body, dict) else 'body is not a dict'
                })
            }

        # Analyze data
        result = analyze_data(data)

        # Check if analysis was successful
        if not result.get('success'):
            return {
                'statusCode': 400,
                'headers': {
                    'Content-Type': 'application/json',
                    'Access-Control-Allow-Origin': '*'
                },
                'body': json.dumps({
                    'success': False,
                    'error': result.get('error'),
                    'traceback': result.get('traceback')
                })
            }

        # Return PDF as binary
        pdf_bytes = result['pdf_bytes']
        filename = result['filename']

        return {
            'statusCode': 200,
            'headers': {
                'Content-Type': 'application/pdf',
                'Content-Disposition': f'attachment; filename="{filename}"',
                'Access-Control-Allow-Origin': '*'
            },
            'body': base64.b64encode(pdf_bytes).decode('utf-8'),
            'isBase64Encoded': True
        }

    except Exception as e:
        import traceback
        return {
            'statusCode': 500,
            'headers': {'Content-Type': 'application/json'},
            'body': json.dumps({
                'success': False,
                'error': str(e),
                'traceback': traceback.format_exc()
            })
        }
