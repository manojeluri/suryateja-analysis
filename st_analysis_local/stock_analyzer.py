"""
Enhanced Stock & Inventory Analyzer
Analyzes stock movement and inventory health
Period: Apr 2025 to Oct 2025
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
plt.rcParams['font.weight'] = 'bold'
plt.rcParams['axes.labelweight'] = 'bold'
plt.rcParams['axes.titleweight'] = 'bold'

# ============================================
# Company Product Mappings - Loaded from CSV files
# ============================================
def load_company_products_from_csv(folder_path="Company Wise Products"):
    """
    Load company-product mappings from CSV files in the specified folder
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
        company_name = csv_file.replace('_Products.csv', '').replace('_Product_Names.csv', '')

        # Clean up company name
        company_name = company_name.replace('_', ' ')
        if company_name == 'BestAgrolife':
            company_name = 'Best Agrolife'
        elif company_name == 'NovaAgriScience':
            company_name = 'Nova Agri Science'

        # Read CSV file
        file_path = os.path.join(folder_path, csv_file)
        try:
            df = pd.read_csv(file_path)

            # Get product names
            if 'Product Name' in df.columns:
                products = df['Product Name'].dropna().str.strip().tolist()
            else:
                products = df.iloc[:, 0].dropna().str.strip().tolist()

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
# Stock Analysis Functions
# ============================================

def load_stock_data(filename):
    """Load stock data from XLS file"""
    print(f"\n📂 Loading stock data from: {filename}")

    if not os.path.exists(filename):
        raise FileNotFoundError(f"File not found: {filename}")

    # Read Excel file
    try:
        df = pd.read_excel(filename, engine='xlrd')
    except:
        try:
            df = pd.read_excel(filename, engine='openpyxl')
        except Exception as e:
            raise Exception(f"Failed to read Excel file: {str(e)}")

    print(f"✅ Loaded {len(df)} rows and {len(df.columns)} columns")
    print(f"📋 Columns: {', '.join(df.columns.tolist())}")

    return df

def categorize_by_company(df):
    """Add company categorization to dataframe"""
    df['COMPANY'] = df['ITNAME'].map(PRODUCT_TO_COMPANY).fillna('Other')
    return df

def calculate_stock_metrics(df):
    """Calculate comprehensive stock and inventory metrics"""

    # Stock turnover ratio (what % of total stock moved out)
    df['TURNOVER_RATIO'] = np.where(df['TOTAL'] > 0,
                                     (df['OUTWARD'] / df['TOTAL'] * 100), 0)

    # Stock efficiency (how well stock is utilized)
    df['STOCK_EFFICIENCY'] = np.where(df['TOTAL'] > 0,
                                       (df['OUTWARD'] / df['TOTAL'] * 100), 0)

    # Movement indicator
    df['HAS_MOVEMENT'] = (df['INWARD'] > 0) | (df['OUTWARD'] > 0)

    # Stock status categorization
    def categorize_stock_status(row):
        if row['CLST'] <= 0:
            return 'Out of Stock'
        elif row['CLST'] <= 5:
            return 'Low Stock'
        elif row['CLST'] <= 20:
            return 'Medium Stock'
        elif row['CLST'] <= 50:
            return 'Good Stock'
        else:
            return 'High Stock'

    df['STOCK_STATUS'] = df.apply(categorize_stock_status, axis=1)

    # Movement categorization
    def categorize_movement(row):
        if row['OUTWARD'] == 0:
            return 'No Movement'
        elif row['TURNOVER_RATIO'] >= 70:
            return 'Fast Moving'
        elif row['TURNOVER_RATIO'] >= 40:
            return 'Medium Moving'
        else:
            return 'Slow Moving'

    df['MOVEMENT_TYPE'] = df.apply(categorize_movement, axis=1)

    # Problem indicators
    df['HAS_NEGATIVE_STOCK'] = df['CLST'] < 0
    df['IS_DEAD_STOCK'] = (df['INWARD'] == 0) & (df['OUTWARD'] == 0) & (df['CLST'] > 0)
    df['IS_OVERSTOCKED'] = (df['CLST'] > 100) & (df['TURNOVER_RATIO'] < 30)

    return df

def generate_stock_overview(df):
    """Generate overall stock overview metrics"""

    overview = {
        'Total_Products': len(df),
        'Total_Opening_Stock': df['OPST'].sum(),
        'Total_Inward': df['INWARD'].sum(),
        'Total_Available': df['TOTAL'].sum(),
        'Total_Outward': df['OUTWARD'].sum(),
        'Total_Closing_Stock': df['CLST'].sum(),
        'Avg_Turnover_Ratio': df['TURNOVER_RATIO'].mean(),
        'Products_In_Stock': len(df[df['CLST'] > 0]),
        'Products_Out_Of_Stock': len(df[df['CLST'] <= 0]),
        'Products_With_Negative_Stock': len(df[df['CLST'] < 0]),
        'Dead_Stock_Products': len(df[df['IS_DEAD_STOCK']]),
        'Fast_Moving_Products': len(df[df['MOVEMENT_TYPE'] == 'Fast Moving']),
        'Slow_Moving_Products': len(df[df['MOVEMENT_TYPE'] == 'Slow Moving']),
        'No_Movement_Products': len(df[df['MOVEMENT_TYPE'] == 'No Movement']),
        'Overstocked_Products': len(df[df['IS_OVERSTOCKED']])
    }

    return overview

def generate_company_stock_analysis(df):
    """Generate company-wise stock analysis"""

    company_summary = df.groupby('COMPANY').agg({
        'ITNAME': 'count',
        'OPST': 'sum',
        'INWARD': 'sum',
        'TOTAL': 'sum',
        'OUTWARD': 'sum',
        'CLST': 'sum',
        'TURNOVER_RATIO': 'mean',
        'HAS_NEGATIVE_STOCK': 'sum',
        'IS_DEAD_STOCK': 'sum',
        'IS_OVERSTOCKED': 'sum'
    }).round(2)

    company_summary.columns = [
        'Product_Count', 'Opening_Stock', 'Inward', 'Total_Available',
        'Outward', 'Closing_Stock', 'Avg_Turnover_Ratio',
        'Negative_Stock_Count', 'Dead_Stock_Count', 'Overstocked_Count'
    ]

    # Calculate stock availability rate
    company_summary['Stock_Availability_%'] = (
        (company_summary['Product_Count'] - (company_summary['Closing_Stock'] <= 0).sum()) /
        company_summary['Product_Count'] * 100
    ).round(2)

    company_summary = company_summary.sort_values('Closing_Stock', ascending=False)

    return company_summary

def identify_fast_movers(df, top_n=15):
    """Identify fast-moving products"""
    fast_movers = df[df['OUTWARD'] > 0].nlargest(top_n, 'OUTWARD')[
        ['ITNAME', 'COMPANY', 'OUTWARD', 'CLST', 'TURNOVER_RATIO', 'STOCK_STATUS']
    ]
    return fast_movers

def identify_slow_movers(df, top_n=15):
    """Identify slow-moving products (have stock but low movement)"""
    slow_movers = df[
        (df['CLST'] > 0) & (df['OUTWARD'] > 0)
    ].nsmallest(top_n, 'TURNOVER_RATIO')[
        ['ITNAME', 'COMPANY', 'CLST', 'OUTWARD', 'TURNOVER_RATIO', 'STOCK_STATUS']
    ]
    return slow_movers

def identify_dead_stock(df):
    """Identify dead stock (no movement at all)"""
    dead_stock = df[df['IS_DEAD_STOCK']][
        ['ITNAME', 'COMPANY', 'CLST', 'STOCK_STATUS']
    ].sort_values('CLST', ascending=False)
    return dead_stock

def identify_problem_areas(df):
    """Identify various problem areas in inventory"""

    problems = {
        'negative_stock': df[df['HAS_NEGATIVE_STOCK']][
            ['ITNAME', 'COMPANY', 'OPST', 'INWARD', 'OUTWARD', 'CLST']
        ].sort_values('CLST'),

        'out_of_stock': df[df['CLST'] == 0][
            ['ITNAME', 'COMPANY', 'OUTWARD', 'STOCK_STATUS']
        ].sort_values('OUTWARD', ascending=False),

        'overstocked': df[df['IS_OVERSTOCKED']][
            ['ITNAME', 'COMPANY', 'CLST', 'OUTWARD', 'TURNOVER_RATIO']
        ].sort_values('CLST', ascending=False),

        'low_stock_high_demand': df[
            (df['CLST'] <= 10) & (df['CLST'] > 0) & (df['TURNOVER_RATIO'] >= 50)
        ][
            ['ITNAME', 'COMPANY', 'CLST', 'OUTWARD', 'TURNOVER_RATIO']
        ].sort_values('TURNOVER_RATIO', ascending=False)
    }

    return problems

def analyze_stock_distribution(df):
    """Analyze how stock is distributed across categories"""

    status_dist = df['STOCK_STATUS'].value_counts()
    movement_dist = df['MOVEMENT_TYPE'].value_counts()

    return {
        'status_distribution': status_dist,
        'movement_distribution': movement_dist
    }

# ============================================
# Visualization Functions
# ============================================

def create_stock_visualizations(df, company_summary, overview, output_dir='outputs_stock', report_type=''):
    """Create comprehensive stock visualizations"""

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    title_prefix = f"{report_type} - " if report_type else ""
    viz_files = []

    # 1. Top 15 Fast-Moving Products
    fig, ax = plt.subplots(figsize=(15, 9))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#f8f9fa')

    fast_movers = df[df['OUTWARD'] > 0].nlargest(15, 'OUTWARD')
    # Beautiful gradient from light green to dark green
    colors = plt.cm.YlGn(np.linspace(0.5, 0.95, len(fast_movers)))

    bars = ax.barh(range(len(fast_movers)), fast_movers['OUTWARD'],
                   color=colors, alpha=0.85, height=0.7)

    ax.set_yticks(range(len(fast_movers)))
    ax.set_yticklabels(fast_movers['ITNAME'], fontsize=12, fontweight='bold')
    ax.set_xlabel('Outward Quantity (Units)', fontsize=14, fontweight='bold', color='#2c3e50')
    ax.set_title(f'{title_prefix}Top 15 Fast-Moving Products',
                 fontsize=18, fontweight='bold', pad=20, color='#27ae60')
    ax.invert_yaxis()

    # Add value labels with better formatting
    for i, v in enumerate(fast_movers['OUTWARD']):
        ax.text(v, i, f'  {int(v)}', va='center', fontsize=11,
                fontweight='bold', color='#2c3e50')

    ax.grid(axis='x', alpha=0.25, linestyle='--', linewidth=0.8)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#dee2e6')
    ax.spines['bottom'].set_color('#dee2e6')

    plt.tight_layout()
    filename = f'{output_dir}/01_fast_movers.png'
    plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white')
    viz_files.append(filename)
    plt.close()

    # 2. Stock Status Distribution
    fig = plt.figure(figsize=(13, 10))
    fig.patch.set_facecolor('white')
    ax = fig.add_subplot(111)

    status_counts = df['STOCK_STATUS'].value_counts()
    # Beautiful color palette - softer and more modern
    colors_status = ['#e74c3c', '#fd79a8', '#fdcb6e', '#55efc4', '#00b894']
    explode = [0.08 if x == 'Out of Stock' else 0.02 for x in status_counts.index]

    # Create pie chart
    wedges, texts, autotexts = ax.pie(status_counts.values, labels=status_counts.index,
                                        autopct='%1.1f%%', startangle=140, colors=colors_status,
                                        explode=explode,
                                        textprops={'fontsize': 13, 'fontweight': 'bold'})

    # Style percentage text
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontsize(14)
        autotext.set_fontweight('bold')

    # Style labels
    for text in texts:
        text.set_fontsize(13)
        text.set_fontweight('bold')
        text.set_color('#2c3e50')

    ax.set_title(f'{title_prefix}Stock Level Distribution',
                 fontsize=20, fontweight='bold', pad=25, color='#2c3e50')

    # Add definitions at the bottom with better styling
    definitions = """
Stock Level Definitions:
• Out of Stock: Closing stock ≤ 0 (No inventory available)
• Low Stock: Closing stock 1-5 units (Critical - Reorder needed)
• Medium Stock: Closing stock 6-20 units (Moderate inventory)
• Good Stock: Closing stock 21-50 units (Healthy inventory level)
• High Stock: Closing stock > 50 units (Abundant inventory)
    """

    fig.text(0.5, 0.06, definitions, ha='center', va='top', fontsize=10.5,
             fontweight='bold', family='sans-serif', color='#2c3e50',
             bbox=dict(boxstyle='round,pad=1', facecolor='#f8f9fa',
                      edgecolor='#95a5a6', linewidth=1.5, alpha=0.9))

    plt.tight_layout()
    filename = f'{output_dir}/02_stock_distribution.png'
    plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white')
    viz_files.append(filename)
    plt.close()

    # 3. Company-wise Stock Summary
    fig, ax = plt.subplots(figsize=(15, 9))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#f8f9fa')

    top_companies = company_summary.nlargest(15, 'Closing_Stock')
    # Beautiful blue gradient - dark at top to light at bottom
    colors = plt.cm.Blues(np.linspace(0.95, 0.5, len(top_companies)))

    bars = ax.barh(range(len(top_companies)), top_companies['Closing_Stock'],
                   color=colors, alpha=0.85, height=0.7)

    ax.set_yticks(range(len(top_companies)))
    ax.set_yticklabels(top_companies.index, fontsize=12, fontweight='bold')
    ax.set_xlabel('Closing Stock (Units)', fontsize=14, fontweight='bold', color='#2c3e50')
    ax.set_title(f'{title_prefix}Top 15 Companies by Closing Stock',
                 fontsize=18, fontweight='bold', pad=20, color='#3498db')
    ax.invert_yaxis()

    for i, v in enumerate(top_companies['Closing_Stock']):
        ax.text(v, i, f'  {int(v)}', va='center', fontsize=11,
                fontweight='bold', color='#2c3e50')

    ax.grid(axis='x', alpha=0.25, linestyle='--', linewidth=0.8)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#dee2e6')
    ax.spines['bottom'].set_color('#dee2e6')

    plt.tight_layout()
    filename = f'{output_dir}/03_company_stock.png'
    plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white')
    viz_files.append(filename)
    plt.close()

    # 4. Movement Type Distribution
    fig = plt.figure(figsize=(13, 9))
    fig.patch.set_facecolor('white')
    ax = fig.add_subplot(111)

    movement_counts = df['MOVEMENT_TYPE'].value_counts()
    # Beautiful modern color palette
    colors_movement = ['#00b894', '#fdcb6e', '#e17055', '#b2bec3']
    explode = [0.05 if x == 'Fast Moving' else 0.02 for x in movement_counts.index]

    wedges, texts, autotexts = ax.pie(movement_counts.values, labels=movement_counts.index,
                                        autopct='%1.1f%%', startangle=45, colors=colors_movement,
                                        explode=explode,
                                        textprops={'fontsize': 13, 'fontweight': 'bold'})

    # Style percentage text
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontsize(14)
        autotext.set_fontweight('bold')

    # Style labels
    for text in texts:
        text.set_fontsize(13)
        text.set_fontweight('bold')
        text.set_color('#2c3e50')

    ax.set_title(f'{title_prefix}Product Movement Distribution',
                 fontsize=20, fontweight='bold', pad=25, color='#2c3e50')

    plt.tight_layout()
    filename = f'{output_dir}/04_movement_distribution.png'
    plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white')
    viz_files.append(filename)
    plt.close()

    # 5. Stock Turnover Analysis (Top 20 companies)
    fig, ax = plt.subplots(figsize=(15, 10))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#f8f9fa')

    top_turnover = company_summary.nlargest(20, 'Avg_Turnover_Ratio')
    # Red to Yellow to Green gradient based on performance
    colors = plt.cm.RdYlGn(np.linspace(0.3, 0.95, len(top_turnover)))

    bars = ax.barh(range(len(top_turnover)), top_turnover['Avg_Turnover_Ratio'],
                   color=colors, alpha=0.85, height=0.7)

    ax.set_yticks(range(len(top_turnover)))
    ax.set_yticklabels(top_turnover.index, fontsize=11, fontweight='bold')
    ax.set_xlabel('Average Turnover Ratio (%)', fontsize=14, fontweight='bold', color='#2c3e50')
    ax.set_title(f'{title_prefix}Top 20 Companies by Stock Turnover Ratio',
                 fontsize=18, fontweight='bold', pad=20, color='#e67e22')
    ax.invert_yaxis()

    for i, v in enumerate(top_turnover['Avg_Turnover_Ratio']):
        ax.text(v, i, f'  {v:.1f}%', va='center', fontsize=10,
                fontweight='bold', color='#2c3e50')

    ax.grid(axis='x', alpha=0.25, linestyle='--', linewidth=0.8)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#dee2e6')
    ax.spines['bottom'].set_color('#dee2e6')

    plt.tight_layout()
    filename = f'{output_dir}/05_turnover_ratio.png'
    plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white')
    viz_files.append(filename)
    plt.close()

    # 6. Inward vs Outward Comparison (Top 15 products)
    fig, ax = plt.subplots(figsize=(15, 9))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#f8f9fa')

    top_activity = df.nlargest(15, 'OUTWARD')
    x = np.arange(len(top_activity))
    width = 0.38

    # Beautiful contrasting colors
    bars1 = ax.barh(x - width/2, top_activity['INWARD'], width, label='Inward',
                    color='#74b9ff', alpha=0.9)
    bars2 = ax.barh(x + width/2, top_activity['OUTWARD'], width, label='Outward',
                    color='#ff7675', alpha=0.9)

    ax.set_yticks(x)
    ax.set_yticklabels(top_activity['ITNAME'], fontsize=12, fontweight='bold')
    ax.set_xlabel('Quantity (Units)', fontsize=14, fontweight='bold', color='#2c3e50')
    ax.set_title(f'{title_prefix}Inward vs Outward - Top 15 Products',
                 fontsize=18, fontweight='bold', pad=20, color='#6c5ce7')
    ax.invert_yaxis()

    # Create beautiful legend
    legend = ax.legend(fontsize=13, loc='lower right', framealpha=0.95,
                      edgecolor='#95a5a6', fancybox=True, shadow=True)
    for text in legend.get_texts():
        text.set_fontweight('bold')
        text.set_color('#2c3e50')

    ax.grid(axis='x', alpha=0.25, linestyle='--', linewidth=0.8)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#dee2e6')
    ax.spines['bottom'].set_color('#dee2e6')

    plt.tight_layout()
    filename = f'{output_dir}/06_inward_vs_outward.png'
    plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white')
    viz_files.append(filename)
    plt.close()

    print(f"\n✅ Generated {len(viz_files)} visualizations in '{output_dir}/' directory")

    return viz_files

def generate_stock_pdf_report(df, company_summary, overview, problems, fast_movers, slow_movers,
                              dead_stock, viz_files, output_dir='outputs_stock', report_type=''):
    """Generate comprehensive PDF report for stock analysis"""

    pdf_filename = f'{output_dir}/Stock_Analysis_Report.pdf'
    report_title = f'{report_type} Stock & Inventory Analysis Report' if report_type else 'Stock & Inventory Analysis Report'

    with PdfPages(pdf_filename) as pdf:
        # Page 1: Cover Page with Executive Summary
        fig = plt.figure(figsize=(11, 8.5))
        fig.patch.set_facecolor('white')
        ax = fig.add_subplot(111)
        ax.axis('off')

        # Title
        fig.text(0.5, 0.95, report_title, ha='center', fontsize=24,
                fontweight='bold', color='#2c3e50')
        fig.text(0.5, 0.91, 'Period: April 2025 - October 2025',
                ha='center', fontsize=11, color='#7f8c8d', style='italic')
        fig.text(0.5, 0.88, f'Generated on {datetime.now().strftime("%B %d, %Y at %H:%M:%S")}',
                ha='center', fontsize=9, color='#95a5a6')

        # Executive Summary Title
        fig.text(0.08, 0.82, '📊 Stock Overview', ha='left', fontsize=18,
                fontweight='bold', color='#2c3e50')

        # Define metric cards
        metrics = [
            ('TOTAL PRODUCTS', f"{overview['Total_Products']:,}", 0.08, 0.68),
            ('CLOSING STOCK', f"{int(overview['Total_Closing_Stock']):,}", 0.29, 0.68),
            ('TOTAL OUTWARD', f"{int(overview['Total_Outward']):,}", 0.50, 0.68),
            ('AVG TURNOVER', f"{overview['Avg_Turnover_Ratio']:.1f}%", 0.71, 0.68),
            ('IN STOCK', f"{overview['Products_In_Stock']:,}", 0.08, 0.53),
            ('OUT OF STOCK', f"{overview['Products_Out_Of_Stock']:,}", 0.29, 0.53),
            ('FAST MOVERS', f"{overview['Fast_Moving_Products']:,}", 0.50, 0.53),
            ('SLOW MOVERS', f"{overview['Slow_Moving_Products']:,}", 0.71, 0.53),
        ]

        card_width = 0.19
        card_height = 0.12

        for idx, (label, value, x, y) in enumerate(metrics):
            # Color coding based on metric type
            if 'OUT OF STOCK' in label or 'SLOW' in label:
                card_color = '#e74c3c'
            elif 'FAST' in label or 'IN STOCK' in label:
                card_color = '#27ae60'
            else:
                card_color = '#3498db'

            fancy_box = FancyBboxPatch((x, y), card_width, card_height,
                                      boxstyle="round,pad=0.005",
                                      facecolor=card_color, edgecolor='none',
                                      transform=fig.transFigure, zorder=2,
                                      alpha=0.9)
            fig.patches.append(fancy_box)

            fig.text(x + card_width/2, y + card_height - 0.02, label,
                    ha='center', va='top', fontsize=8, color='white',
                    fontweight='normal', alpha=0.95)

            fig.text(x + card_width/2, y + 0.035, value,
                    ha='center', va='bottom', fontsize=15, color='white',
                    fontweight='bold')

        # Key Insights Section
        fig.text(0.08, 0.45, '💡 Key Insights', ha='left', fontsize=18,
                fontweight='bold', color='#2c3e50')

        insights_text = f"""Inventory Health:
  • {overview['Products_In_Stock']} products in stock ({overview['Products_In_Stock']/overview['Total_Products']*100:.1f}%)
  • {overview['Products_Out_Of_Stock']} products out of stock
  • {overview['Products_With_Negative_Stock']} products with negative stock (data issues)

Movement Analysis:
  • {overview['Fast_Moving_Products']} fast-moving products
  • {overview['Slow_Moving_Products']} slow-moving products
  • {overview['No_Movement_Products']} products with no movement
  • Average turnover ratio: {overview['Avg_Turnover_Ratio']:.1f}%

Problem Areas:
  • {overview['Dead_Stock_Products']} dead stock items
  • {overview['Overstocked_Products']} overstocked products
  • Total inward: {int(overview['Total_Inward']):,} units
  • Total outward: {int(overview['Total_Outward']):,} units"""

        fig.text(0.5, 0.23, insights_text, ha='center', va='center', fontsize=9.5,
                family='monospace', bbox=dict(boxstyle='round', facecolor='#ecf0f1',
                edgecolor='#34495e', linewidth=1.5, pad=1))

        pdf.savefig(fig, bbox_inches='tight')
        plt.close()

        # Page 2: Company Stock Analysis
        fig = plt.figure(figsize=(11, 8.5))
        fig.patch.set_facecolor('white')
        ax = fig.add_subplot(111)
        ax.axis('off')

        fig.text(0.5, 0.95, 'Company Stock Analysis', ha='center', fontsize=20,
                fontweight='bold', color='#2c3e50')

        # Create table data
        table_data = []
        table_data.append(['Company', 'Products', 'Closing Stock', 'Outward', 'Turnover %'])

        for idx, (company, row) in enumerate(company_summary.head(20).iterrows(), 1):
            table_data.append([
                company[:20],
                str(int(row['Product_Count'])),
                str(int(row['Closing_Stock'])),
                str(int(row['Outward'])),
                f"{row['Avg_Turnover_Ratio']:.1f}%"
            ])

        table = ax.table(cellText=table_data, cellLoc='left',
                        loc='center', bbox=[0.05, 0.05, 0.9, 0.85])
        table.auto_set_font_size(False)
        table.set_fontsize(9)
        table.scale(1, 1.6)

        # Style header row
        for i in range(5):
            cell = table[(0, i)]
            cell.set_facecolor('#34495e')
            cell.set_text_props(weight='bold', color='white', size=10)

        # Alternate row colors and make text bold
        for i in range(1, len(table_data)):
            for j in range(5):
                cell = table[(i, j)]
                cell.set_text_props(weight='bold', size=9)
                if i % 2 == 0:
                    cell.set_facecolor('#ecf0f1')

        pdf.savefig(fig, bbox_inches='tight')
        plt.close()

        # Page 3: Fast Movers
        fig = plt.figure(figsize=(11, 8.5))
        fig.patch.set_facecolor('white')
        ax = fig.add_subplot(111)
        ax.axis('off')

        fig.text(0.5, 0.95, 'Top Fast-Moving Products', ha='center', fontsize=18,
                fontweight='bold', color='#27ae60')

        table_data = []
        table_data.append(['Rank', 'Product', 'Company', 'Outward', 'Closing', 'Turnover %'])

        for idx, row in enumerate(fast_movers.head(15).itertuples(), 1):
            table_data.append([
                str(idx),
                row.ITNAME[:25],
                row.COMPANY[:15],
                str(int(row.OUTWARD)),
                str(int(row.CLST)),
                f"{row.TURNOVER_RATIO:.1f}%"
            ])

        table = ax.table(cellText=table_data, cellLoc='left',
                        loc='upper center', bbox=[0.05, 0.35, 0.9, 0.55])
        table.auto_set_font_size(False)
        table.set_fontsize(9)
        table.scale(1, 1.9)

        for i in range(6):
            cell = table[(0, i)]
            cell.set_facecolor('#27ae60')
            cell.set_text_props(weight='bold', color='white', size=10)

        for i in range(1, len(table_data)):
            for j in range(6):
                cell = table[(i, j)]
                cell.set_text_props(weight='bold', size=9)
                if i % 2 == 0:
                    cell.set_facecolor('#d5f4e6')

        pdf.savefig(fig, bbox_inches='tight')
        plt.close()

        # Page 4: Problem Areas
        fig = plt.figure(figsize=(11, 8.5))
        fig.patch.set_facecolor('white')
        ax = fig.add_subplot(111)
        ax.axis('off')

        fig.text(0.5, 0.95, 'Problem Areas - Attention Required', ha='center', fontsize=18,
                fontweight='bold', color='#e74c3c')

        # Negative Stock
        fig.text(0.08, 0.88, f'⚠️  Negative Stock ({len(problems["negative_stock"])} products)',
                ha='left', fontsize=12, fontweight='bold', color='#c0392b')

        if len(problems["negative_stock"]) > 0:
            neg_table_data = [['Product', 'Company', 'Closing Stock']]
            for idx, row in enumerate(problems["negative_stock"].head(8).itertuples(), 1):
                neg_table_data.append([
                    row.ITNAME[:30],
                    row.COMPANY[:15],
                    str(int(row.CLST))
                ])

            neg_table = ax.table(cellText=neg_table_data, cellLoc='left',
                               loc='upper left', bbox=[0.05, 0.58, 0.9, 0.28])
            neg_table.auto_set_font_size(False)
            neg_table.set_fontsize(8.5)

            for i in range(3):
                cell = neg_table[(0, i)]
                cell.set_facecolor('#e74c3c')
                cell.set_text_props(weight='bold', color='white', size=9)

            for i in range(1, len(neg_table_data)):
                for j in range(3):
                    cell = neg_table[(i, j)]
                    cell.set_text_props(weight='bold', size=8.5)

        # Out of Stock
        fig.text(0.08, 0.52, f'📦 Out of Stock ({len(problems["out_of_stock"])} products)',
                ha='left', fontsize=12, fontweight='bold', color='#d35400')

        if len(problems["out_of_stock"]) > 0:
            oos_table_data = [['Product', 'Company', 'Outward']]
            for idx, row in enumerate(problems["out_of_stock"].head(8).itertuples(), 1):
                oos_table_data.append([
                    row.ITNAME[:30],
                    row.COMPANY[:15],
                    str(int(row.OUTWARD))
                ])

            oos_table = ax.table(cellText=oos_table_data, cellLoc='left',
                                loc='upper left', bbox=[0.05, 0.22, 0.9, 0.28])
            oos_table.auto_set_font_size(False)
            oos_table.set_fontsize(8.5)

            for i in range(3):
                cell = oos_table[(0, i)]
                cell.set_facecolor('#e67e22')
                cell.set_text_props(weight='bold', color='white', size=9)

            for i in range(1, len(oos_table_data)):
                for j in range(3):
                    cell = oos_table[(i, j)]
                    cell.set_text_props(weight='bold', size=8.5)

        pdf.savefig(fig, bbox_inches='tight')
        plt.close()

        # Add all visualization images
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
        d['Author'] = 'Stock & Inventory Analyzer'
        d['Subject'] = 'Stock & Inventory Analysis'
        d['Keywords'] = 'Stock, Inventory, Analysis, Report'
        d['CreationDate'] = datetime.now()

    print(f"\n✅ PDF report generated: {pdf_filename}")
    return pdf_filename

# ============================================
# Main Processing Function
# ============================================

def process_stock_file(filename, output_dir, report_type):
    """Process a single stock file and generate reports"""

    try:
        print(f"\n{'='*60}")
        print(f"📊 Processing {report_type} Stock Data: {filename}")
        print(f"{'='*60}")

        # Load data
        df = load_stock_data(filename)

        # Add company categorization
        print("\n🏷️  Categorizing products by company...")
        df = categorize_by_company(df)

        # Calculate stock metrics
        print("📊 Calculating stock metrics...")
        df = calculate_stock_metrics(df)

        # Generate analyses
        print("🔍 Generating stock overview...")
        overview = generate_stock_overview(df)

        print("🔍 Generating company stock analysis...")
        company_summary = generate_company_stock_analysis(df)

        print("🔍 Identifying fast-moving products...")
        fast_movers = identify_fast_movers(df)

        print("🔍 Identifying slow-moving products...")
        slow_movers = identify_slow_movers(df)

        print("🔍 Identifying dead stock...")
        dead_stock = identify_dead_stock(df)

        print("🔍 Identifying problem areas...")
        problems = identify_problem_areas(df)

        print("🔍 Analyzing stock distribution...")
        distributions = analyze_stock_distribution(df)

        # Create visualizations
        print("\n📈 Creating visualizations...")
        viz_files = create_stock_visualizations(df, company_summary, overview,
                                               output_dir=output_dir, report_type=report_type)

        # Generate PDF report
        print("\n📄 Generating comprehensive PDF report...")
        report_file = generate_stock_pdf_report(df, company_summary, overview, problems,
                                                fast_movers, slow_movers, dead_stock, viz_files,
                                                output_dir=output_dir, report_type=report_type)

        # Save detailed data to Excel
        print("\n💾 Saving analysis data to Excel...")
        with pd.ExcelWriter(f'{output_dir}/Stock_Analysis_Summary.xlsx', engine='openpyxl') as writer:
            company_summary.to_excel(writer, sheet_name='Company Analysis')
            fast_movers.to_excel(writer, sheet_name='Fast Movers', index=False)
            slow_movers.to_excel(writer, sheet_name='Slow Movers', index=False)
            dead_stock.to_excel(writer, sheet_name='Dead Stock', index=False)
            problems['negative_stock'].to_excel(writer, sheet_name='Negative Stock', index=False)
            problems['out_of_stock'].to_excel(writer, sheet_name='Out of Stock', index=False)
            problems['overstocked'].to_excel(writer, sheet_name='Overstocked', index=False)
            problems['low_stock_high_demand'].to_excel(writer, sheet_name='Low Stock High Demand', index=False)

            # Overview summary
            overview_df = pd.DataFrame([overview])
            overview_df.to_excel(writer, sheet_name='Overview', index=False)

        print(f"\n{'='*60}")
        print(f"✅ {report_type} STOCK ANALYSIS COMPLETE!")
        print(f"{'='*60}")
        print(f"\n📊 Results Summary:")
        print(f"   • PDF Report: {output_dir}/Stock_Analysis_Report.pdf")
        print(f"   • Excel Summary: {output_dir}/Stock_Analysis_Summary.xlsx")
        print(f"   • Visualizations: {len(viz_files)} charts in {output_dir}/ folder")

        return True

    except Exception as e:
        print(f"\n❌ Error during {report_type} stock analysis: {str(e)}")
        import traceback
        print("\n📋 Full error trace:")
        print(traceback.format_exc())
        return False

def main():
    """Main execution function"""

    print("="*60)
    print("🚀 STOCK & INVENTORY ANALYZER")
    print("="*60)
    print("📅 Analysis Period: April 2025 - October 2025")

    # Look for stock files in Stock Files folder
    pesticide_file = 'Stock Files/STKST_PS.XLS'
    fertilizer_file = 'Stock Files/STKST_FS.XLS'

    files_to_process = []

    if os.path.exists(pesticide_file):
        files_to_process.append({
            'filename': pesticide_file,
            'output_dir': 'outputs_stock_pesticides',
            'report_type': 'Pesticides'
        })

    if os.path.exists(fertilizer_file):
        files_to_process.append({
            'filename': fertilizer_file,
            'output_dir': 'outputs_stock_fertilizers',
            'report_type': 'Fertilizers'
        })

    if not files_to_process:
        print("\n❌ No stock files found!")
        print("📝 Looking for: STKST_PS.XLS (Pesticides) and/or STKST_FS.XLS (Fertilizers)")
        print("📁 Expected location: Stock Files/ folder")
        return

    print(f"\n📁 Found {len(files_to_process)} stock file(s) to process:")
    for item in files_to_process:
        print(f"   • {item['filename']} → {item['report_type']}")

    # Process each file
    results = []
    for item in files_to_process:
        success = process_stock_file(item['filename'], item['output_dir'], item['report_type'])
        results.append({
            'type': item['report_type'],
            'success': success,
            'output_dir': item['output_dir']
        })

    # Final summary
    print("\n" + "="*60)
    print("🎉 ALL STOCK ANALYSIS COMPLETE!")
    print("="*60)
    print("\n📊 Final Summary:")
    for result in results:
        if result['success']:
            print(f"\n✅ {result['type']}:")
            print(f"   • PDF: {result['output_dir']}/Stock_Analysis_Report.pdf")
            print(f"   • Excel: {result['output_dir']}/Stock_Analysis_Summary.xlsx")
        else:
            print(f"\n❌ {result['type']}: Failed to generate reports")

    print("\n💡 Open the PDF reports to view detailed stock analysis!")
    print("="*60)

if __name__ == "__main__":
    main()
