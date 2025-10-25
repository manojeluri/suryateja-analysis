from flask import Flask, request, jsonify
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend for server
import matplotlib.pyplot as plt
import seaborn as sns
import base64
from io import BytesIO
import json
import traceback
import os

app = Flask(__name__)

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
# Data Analysis Functions
# ============================================
def categorize_by_company(df):
    """
    Categorize products by company and create company-wise summaries
    """
    # Add company column to dataframe
    df['COMPANY'] = df['ITNAME'].map(PRODUCT_TO_COMPANY).fillna('Other')
    
    # Create company-wise summary
    company_summary = df.groupby('COMPANY').agg({
        'ITNAME': 'count',
        'QTY': 'sum',
        'TAXBLEAMT': 'sum'
    }).rename(columns={
        'ITNAME': 'Product Count',
        'QTY': 'Total Quantity',
        'TAXBLEAMT': 'Total Revenue'
    }).sort_values('Total Revenue', ascending=False)
    
    return df, company_summary

def analyze_excel_data(df):
    """
    Analyze sales data from SALANAL.XLS
    Columns: HSNCODE, ITNAME, QTY, TAXBLEAMT, GST
    """
    analysis_results = {
        "summary": {},
        "statistics": {},
        "insights": [],
        "top_products": {},
        "tax_analysis": {},
        "company_analysis": {}
    }

    # Add company categorization
    df, company_summary = categorize_by_company(df)

    # Basic summary
    analysis_results["summary"] = {
        "total_items": len(df),
        "total_columns": len(df.columns),
        "column_names": df.columns.tolist(),
        "total_companies": int(df['COMPANY'].nunique()),
        "categorized_products": int((df['COMPANY'] != 'Other').sum()),
        "uncategorized_products": int((df['COMPANY'] == 'Other').sum())
    }

    # Sales Statistics
    if 'QTY' in df.columns:
        total_quantity = df['QTY'].sum()
        analysis_results["statistics"]["total_quantity_sold"] = int(total_quantity)

    if 'TAXBLEAMT' in df.columns:
        total_amount = df['TAXBLEAMT'].sum()
        avg_amount = df['TAXBLEAMT'].mean()
        analysis_results["statistics"]["total_sales_amount"] = f"₹{total_amount:,.2f}"
        analysis_results["statistics"]["average_sale_value"] = f"₹{avg_amount:,.2f}"

    if 'GST' in df.columns:
        total_gst = (df['TAXBLEAMT'] * df['GST'] / 100).sum()
        analysis_results["statistics"]["total_gst_collected"] = f"₹{total_gst:,.2f}"

    # Top Products by Quantity
    if 'ITNAME' in df.columns and 'QTY' in df.columns:
        top_5_qty = df.nlargest(5, 'QTY')[['ITNAME', 'QTY', 'COMPANY']].to_dict('records')
        analysis_results["top_products"]["by_quantity"] = top_5_qty

    # Top Products by Revenue
    if 'ITNAME' in df.columns and 'TAXBLEAMT' in df.columns:
        top_5_revenue = df.nlargest(5, 'TAXBLEAMT')[['ITNAME', 'TAXBLEAMT', 'COMPANY']].to_dict('records')
        analysis_results["top_products"]["by_revenue"] = top_5_revenue

    # Company Analysis
    if not company_summary.empty:
        analysis_results["company_analysis"]["summary"] = company_summary.to_dict('index')
        
        # Top 5 companies by revenue
        top_companies = company_summary.nlargest(5, 'Total Revenue')
        analysis_results["company_analysis"]["top_companies"] = top_companies.to_dict('index')

    # GST Rate Distribution
    if 'GST' in df.columns:
        gst_distribution = df.groupby('GST').agg({
            'TAXBLEAMT': 'sum',
            'QTY': 'sum'
        }).to_dict()
        analysis_results["tax_analysis"]["gst_rate_distribution"] = gst_distribution

    # Generate insights
    insights = []
    if 'TAXBLEAMT' in df.columns and 'QTY' in df.columns:
        total_items = len(df)
        total_qty = df['QTY'].sum()
        total_revenue = df['TAXBLEAMT'].sum()

        insights.append(f"Total Products: {total_items} different items")
        insights.append(f"Total Quantity Sold: {int(total_qty)} units")
        insights.append(f"Total Revenue: ₹{total_revenue:,.2f}")
        
        # Company insights
        if 'COMPANY' in df.columns:
            total_companies = df['COMPANY'].nunique()
            insights.append(f"Products from {total_companies} companies")

            # Company-wise sales amounts
            if not company_summary.empty:
                insights.append("\n--- Company-wise Sales ---")
                for company, row in company_summary.iterrows():
                    revenue = row['Total Revenue']
                    insights.append(f"{company}: ₹{revenue:,.2f}")

                # Top company summary
                top_company = company_summary.index[0]
                top_company_revenue = company_summary.iloc[0]['Total Revenue']
                insights.append(f"\nTop Company: {top_company} (₹{top_company_revenue:,.2f})")

        if 'GST' in df.columns:
            total_gst = (df['TAXBLEAMT'] * df['GST'] / 100).sum()
            grand_total = total_revenue + total_gst
            insights.append(f"Grand Total (with GST): ₹{grand_total:,.2f}")
            insights.append(f"Total GST: ₹{total_gst:,.2f}")

        # Find best seller
        if 'ITNAME' in df.columns:
            best_seller_idx = df['QTY'].idxmax()
            best_seller = df.loc[best_seller_idx, 'ITNAME']
            best_seller_qty = df['QTY'].max()
            best_seller_company = df.loc[best_seller_idx, 'COMPANY']
            insights.append(f"Best Seller: {best_seller} from {best_seller_company} ({int(best_seller_qty)} units)")

            highest_revenue_idx = df['TAXBLEAMT'].idxmax()
            highest_revenue = df.loc[highest_revenue_idx, 'ITNAME']
            highest_revenue_amt = df['TAXBLEAMT'].max()
            highest_revenue_company = df.loc[highest_revenue_idx, 'COMPANY']
            insights.append(f"Highest Revenue Product: {highest_revenue} from {highest_revenue_company} (₹{highest_revenue_amt:,.2f})")

    analysis_results["insights"] = insights
    return analysis_results

def create_visualizations(df):
    """
    Create sales visualizations including company-wise analysis
    """
    visualizations = []
    
    # Add company categorization
    df, company_summary = categorize_by_company(df)

    # Visualization 1: Top 10 Products by Quantity Sold
    if 'ITNAME' in df.columns and 'QTY' in df.columns:
        fig, ax = plt.subplots(figsize=(12, 6))
        top_10_qty = df.nlargest(10, 'QTY')
        colors = plt.cm.Set3(range(len(top_10_qty)))
        ax.barh(top_10_qty['ITNAME'], top_10_qty['QTY'], color=colors)
        ax.set_xlabel('Quantity Sold')
        ax.set_title('Top 10 Products by Quantity Sold')
        ax.invert_yaxis()
        plt.tight_layout()

        buffer = BytesIO()
        plt.savefig(buffer, format='png', bbox_inches='tight', dpi=150)
        buffer.seek(0)
        img_base64 = base64.b64encode(buffer.read()).decode()
        visualizations.append({
            "title": "Top 10 Products by Quantity",
            "image": img_base64
        })
        plt.close()

    # Visualization 2: Revenue by Product (Top 10)
    if 'ITNAME' in df.columns and 'TAXBLEAMT' in df.columns:
        fig, ax = plt.subplots(figsize=(12, 6))
        top_10_revenue = df.nlargest(10, 'TAXBLEAMT')
        colors = plt.cm.Set2(range(len(top_10_revenue)))
        ax.barh(top_10_revenue['ITNAME'], top_10_revenue['TAXBLEAMT'], color=colors)
        ax.set_xlabel('Revenue (₹)')
        ax.set_title('Top 10 Products by Revenue')
        ax.invert_yaxis()
        plt.tight_layout()

        buffer = BytesIO()
        plt.savefig(buffer, format='png', bbox_inches='tight', dpi=150)
        buffer.seek(0)
        img_base64 = base64.b64encode(buffer.read()).decode()
        visualizations.append({
            "title": "Top 10 Products by Revenue",
            "image": img_base64
        })
        plt.close()

    # Visualization 3: Company-wise Revenue Distribution
    if not company_summary.empty and 'Total Revenue' in company_summary.columns:
        fig, ax = plt.subplots(figsize=(14, 8))
        
        # Get top 10 companies
        top_companies = company_summary.nlargest(10, 'Total Revenue')
        
        colors = plt.cm.Spectral(range(len(top_companies)))
        ax.barh(top_companies.index, top_companies['Total Revenue'], color=colors)
        ax.set_xlabel('Total Revenue (₹)')
        ax.set_title('Top 10 Companies by Revenue')
        ax.invert_yaxis()
        
        # Format x-axis to show values in lakhs/crores
        ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'₹{x/100000:.1f}L' if x < 10000000 else f'₹{x/10000000:.2f}Cr'))
        
        plt.tight_layout()

        buffer = BytesIO()
        plt.savefig(buffer, format='png', bbox_inches='tight', dpi=150)
        buffer.seek(0)
        img_base64 = base64.b64encode(buffer.read()).decode()
        visualizations.append({
            "title": "Company-wise Revenue Distribution",
            "image": img_base64
        })
        plt.close()

    # Visualization 4: Product Count by Company
    if not company_summary.empty and 'Product Count' in company_summary.columns:
        fig, ax = plt.subplots(figsize=(12, 6))
        
        top_companies_count = company_summary.nlargest(10, 'Product Count')
        
        colors = plt.cm.Paired(range(len(top_companies_count)))
        ax.bar(top_companies_count.index, top_companies_count['Product Count'], color=colors)
        ax.set_ylabel('Number of Products')
        ax.set_title('Top 10 Companies by Product Count')
        ax.tick_params(axis='x', rotation=45)
        
        plt.tight_layout()

        buffer = BytesIO()
        plt.savefig(buffer, format='png', bbox_inches='tight', dpi=150)
        buffer.seek(0)
        img_base64 = base64.b64encode(buffer.read()).decode()
        visualizations.append({
            "title": "Product Count by Company",
            "image": img_base64
        })
        plt.close()

    return visualizations

# ============================================
# Flask Routes
# ============================================
@app.route('/', methods=['GET'])
def home():
    """Root endpoint"""
    return jsonify({
        "message": "Excel Analysis API is running!",
        "version": "2.0 - with Company Analysis",
        "endpoints": {
            "health": "/health",
            "analyze": "/analyze (POST with JSON data)"
        }
    }), 200

@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({"status": "healthy", "message": "API is running"}), 200

@app.route('/analyze', methods=['POST'])
def analyze_data():
    """
    Main endpoint to receive data and return analysis
    Handles JSON data from n8n
    """
    try:
        data = request.get_json()

        print("Received request data keys:", data.keys() if data else "None")

        if 'data' not in data:
            return jsonify({"error": "No data provided in request"}), 400

        # Handle different data formats
        received_data = data['data']

        # If it's a string, try to parse it as JSON
        if isinstance(received_data, str):
            print("Data is string, attempting to parse...")
            # Remove the leading '=' if present (from n8n expression)
            if received_data.startswith('='):
                received_data = received_data[1:]
            # Parse the JSON string
            received_data = json.loads(received_data)

        # Convert to DataFrame
        df = pd.DataFrame(received_data)

        print(f"✅ Successfully loaded data with {len(df)} rows and {len(df.columns)} columns")
        print(f"Columns: {df.columns.tolist()}")
        print(f"First row: {df.iloc[0].to_dict() if len(df) > 0 else 'No data'}")

        # Perform analysis
        analysis = analyze_excel_data(df)

        # Create visualizations
        visualizations = create_visualizations(df)

        # Prepare response
        response = {
            "success": True,
            "analysis": analysis,
            "visualizations": visualizations,
            "message": f"Analysis complete! Processed {len(df)} rows with {len(df.columns)} columns across {analysis['summary'].get('total_companies', 0)} companies."
        }

        return jsonify(response), 200

    except Exception as e:
        error_trace = traceback.format_exc()
        print(f"❌ Error: {str(e)}")
        print("Full traceback:")
        print(error_trace)
        return jsonify({
            "success": False,
            "error": str(e),
            "trace": error_trace,
            "message": "Failed to process data"
        }), 500

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
