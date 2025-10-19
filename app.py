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

app = Flask(__name__)

# ============================================
# Data Analysis Functions (from your Colab)
# ============================================
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
        "tax_analysis": {}
    }

    # Basic summary
    analysis_results["summary"] = {
        "total_items": len(df),
        "total_columns": len(df.columns),
        "column_names": df.columns.tolist()
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
        top_5_qty = df.nlargest(5, 'QTY')[['ITNAME', 'QTY']].to_dict('records')
        analysis_results["top_products"]["by_quantity"] = top_5_qty

    # Top Products by Revenue
    if 'ITNAME' in df.columns and 'TAXBLEAMT' in df.columns:
        top_5_revenue = df.nlargest(5, 'TAXBLEAMT')[['ITNAME', 'TAXBLEAMT']].to_dict('records')
        analysis_results["top_products"]["by_revenue"] = top_5_revenue

    # GST Rate Distribution
    if 'GST' in df.columns:
        gst_distribution = df.groupby('GST').agg({
            'TAXBLEAMT': 'sum',
            'QTY': 'sum'
        }).to_dict()
        analysis_results["tax_analysis"]["gst_rate_distribution"] = gst_distribution

    # Generate insights
    if 'TAXBLEAMT' in df.columns and 'QTY' in df.columns:
        total_items = len(df)
        total_qty = df['QTY'].sum()
        total_revenue = df['TAXBLEAMT'].sum()

        analysis_results["insights"].append(f"Total Products: {total_items} different items")
        analysis_results["insights"].append(f"Total Quantity Sold: {int(total_qty)} units")
        analysis_results["insights"].append(f"Total Revenue: ₹{total_revenue:,.2f}")

        if 'GST' in df.columns:
            total_gst = (df['TAXBLEAMT'] * df['GST'] / 100).sum()
            grand_total = total_revenue + total_gst
            analysis_results["insights"].append(f"Grand Total (with GST): ₹{grand_total:,.2f}")
            analysis_results["insights"].append(f"Total GST: ₹{total_gst:,.2f}")

        # Find best seller
        if 'ITNAME' in df.columns:
            best_seller = df.loc[df['QTY'].idxmax(), 'ITNAME']
            best_seller_qty = df['QTY'].max()
            analysis_results["insights"].append(f"Best Seller: {best_seller} ({int(best_seller_qty)} units)")

            highest_revenue = df.loc[df['TAXBLEAMT'].idxmax(), 'ITNAME']
            highest_revenue_amt = df['TAXBLEAMT'].max()
            analysis_results["insights"].append(f"Highest Revenue Product: {highest_revenue} (₹{highest_revenue_amt:,.2f})")

    return analysis_results

def create_visualizations(df):
    """
    Create sales visualizations tailored for SALANAL.XLS data
    """
    visualizations = []

    # Visualization 1: Top 10 Products by Quantity Sold
    if 'ITNAME' in df.columns and 'QTY' in df.columns:
        fig, ax = plt.subplots(figsize=(12, 6))
        top_10_qty = df.nlargest(10, 'QTY')
        ax.barh(top_10_qty['ITNAME'], top_10_qty['QTY'], color='steelblue')
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
        ax.barh(top_10_revenue['ITNAME'], top_10_revenue['TAXBLEAMT'], color='green')
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

    return visualizations

# ============================================
# Flask Routes
# ============================================
@app.route('/', methods=['GET'])
def home():
    """Root endpoint"""
    return jsonify({
        "message": "Excel Analysis API is running!",
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
            "message": f"Analysis complete! Processed {len(df)} rows with {len(df.columns)} columns."
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
