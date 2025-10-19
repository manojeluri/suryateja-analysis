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
# Company Product Mappings
# ============================================
COMPANY_PRODUCTS = {
    "Gharda": [
        "Aniloguard 1 Lt", "Aniloguard 500ml", "Bakeel 1 Lt", "Bakeel 250ml", "Bakeel 500ml",
        "Black 500ml", "Boxer 100ml", "Boxer 250ml", "Chamatkar 250ml", "Chamatkar 500gms",
        "Dunet 100ml", "Guru 1kg", "Guru 500gms", "Samarth 500ml", "Urja 200ml",
        "Vow 1 Lt", "Vow 500ml", "Zolt 500ml"
    ],
    "Adama": [
        "Agas 250gms", "Agas 25gms", "Agas 500gms", "Agil 250ml", "Agil 500ml",
        "Confidor 200ml", "Confidor 70wg 35gms", "Confidor 70wg 70gms", "Nomolt 250ml",
        "Nomolt 500ml", "Pegasus 1kg", "Ronstar 250ml", "Ronstar 500ml", "San Hero 100ml",
        "Tilt 250ml", "Topsin 250gms", "Topsin 500gms"
    ],
    "Best Agrolife": [
        "Axeman 333", "Bestline 200 Gm", "Cubax Power 1 Lt", "Cubax Power 250ml",
        "Cubax Power 500ml", "Dynamita 333ml", "Moximate 1 Lt", "Moximate 250ml",
        "Moximate 500ml", "Pasto 50ml", "Pasto 250ml", "Zygant 1 Lt"
    ],
    "Godrej": [
        "Armurox 300ml", "Asatrus 500gms", "Double 100ml", "Double 500ml",
        "Gracia 160ml", "Gracia 400ml", "Greeva 50gms", "Greeva 250gms",
        "Nominee Gold 250ml", "Nominee Gold 500ml", "Starban 1 Lt", "Starban 250ml"
    ],
    "Indofil": [
        "Alecto 100ml", "Alecto 50 Ml", "Avatar 1kg", "Avatar 500gms", "Baan 120gms",
        "Baan 240gms", "Baan 30gms", "Baan 60gms", "Carbomain 5kg 5kg", "Dost 1kg",
        "Dost 500gms", "Evito 1 Lt", "Evito 250ml", "Evito 500ml", "M-45 1kg",
        "M-45 500gms", "Padiwix 600ml", "Score 1 Lt", "Score 250ml", "Score 50 Ml",
        "Unison 100ml", "Unison 250ml", "Unison 500ml"
    ],
    "Nichino": [
        "Akio 1 Lt", "Akio 500ml", "Bifors 1 Lt", "Bifors 500gms", "Chandika 1 Lt",
        "Chandika 500ml", "Fluxy 1 Lt", "Fluxy 250ml", "Fluxy 500ml", "Fluxyr 1 Lt",
        "Fluxyr 500ml", "Imishot 1 Lt", "Imishot 500ml", "Takaf 1 Lt", "Takaf 500ml"
    ],
    "Nova Agri Science": [
        "Atrino 1kg", "Atrino 500gms", "Azoman 1kg", "Azoman 250gms", "Azoman 500gms",
        "Ecotin 1kg", "Ecotin 500gms", "Instomin 500ml", "Instomin 250ml", "Mantomin 1kg",
        "Micromin 500gms", "Zinkmin 1kg"
    ],
    "Nova Agri Tech": [
        "Nova Feret 19;19;19 1kg", "Nova Feret 19;19;19 25kg", "Nova Fert Cn 1kg",
        "Nova Fert Cn 25kg", "Nova Googley 10kg", "Nova Potash 1kg", "Nova Proto 1kg",
        "Nova Proto 250gms"
    ],
    "Superior": [
        "Benot 250ml", "Dollar 250ml", "Dollar 500ml", "Dollar[gr] 5kg", "Dyeton 250ml",
        "Dyeton 500ml", "Galben 1 Lt", "Kohinoor 250ml", "Kohinoor 500ml", "Movento 250ml"
    ],
    "Swal": [
        "Arryan 400gms", "Arryan 800gms", "Delma 300gms", "Delma 600gms", "Dsp99 1 Lt",
        "Dsp99 250ml", "Dsp99 500ml", "Durnet 250ml", "Durnet 500ml", "Filia 240ml",
        "Filia 480ml", "Juva 250ml", "Juva 500ml", "Kadett 1 Lt", "Kadett 500ml",
        "Lancer Gold 100ml", "Lancer Gold 250ml", "Lancer Gold 500ml", "Montaf 1 Lt",
        "Montaf 500ml", "Radiant 250ml", "Radiant 500ml", "Rimostar 1 Lt", "Rimostar 500ml",
        "Solomon 1 Lt", "Solomon 250ml", "Viraat 1 Lt", "Viraat 500ml"
    ],
    "VNR": [
        "Almix 8gr", "Azaka Duo 200ml", "Benevia 240ml", "Cabriotop 300gms", "Cabriotop 600gms",
        "Rifit 1 Lt", "Rifit 500ml", "Sumilex 500gms"
    ],
    "Balaji": [
        "Acerbo 1 Lt", "Acerbo 200ml", "Aliette 250gms", "Aliette 500gms", "Amicus 1 Lt",
        "Amicus 500ml", "Ampligo 250ml", "Ampligo 80ml", "Areva 500ml", "Benzer 1 Lt",
        "Benzer 250ml", "Coragen 60ml", "Double Power 1 Lt", "Double Power 250ml",
        "Double Power 500ml", "Fame 500ml", "Flint 200gms", "Flint 500gms", "Kamikazi 250ml",
        "Kamikazi 500ml", "Prime Power 200ml", "Prime Power 500ml", "Regent 250ml",
        "Ridomil 1kg", "Ridomil 500gms", "Tata Mida 1 Lt", "Tata Mida 250ml", "Tata Mida 500ml",
        "Virtako 250ml", "Volium Flexi 250ml", "Volium Flexi 500ml", "Volium Targour 250ml"
    ],
    "Chennakesava": [
        "A S A Plus 5kg", "Agripro 1kg", "Agripro 250gms", "Agripro 500gms", "Agrofer 100gms",
        "Agrofer 250gms", "Agromin Gold 1 Lt", "Agromin Gold 500ml", "Agromin Max 1kg",
        "Agromin Max 250gms", "Agromin Max 500gms", "Boran[aries] 1kg", "Boran[aries] 250gms",
        "Boran[aries] 500gms", "Chelamin 100gms", "Cil 10;26;26 50kg", "Cil 14;35;14 50kg",
        "Cil 20;20;0;13 50kg", "Cil Dap 50kg", "Cil Mop 50kg", "Cil Ssp 50kg", "Cil Urea 45kg",
        "Endomyco 125gms", "F 20;20;0;13 50kg", "Godavari Tri Gold 50kg", "Gromor Nano Dap 1 Lt",
        "Gsfc A/s 50kg", "Macrofert 20;20;20 1kg", "Mag Mix 1kg"
    ],
    "PI": [
        "Biovita 1 Lt", "Biovita 500ml", "Biovita Granuls 10kg", "Biovita X 25kg",
        "Biovita X 4kg", "Green Gold 10kg", "Pi Bio Super 10kg", "Pi Force 100ml",
        "Pi Force 250ml", "Pi Force 500ml", "Pi Lanz 100ml", "Pi Lanz 250ml",
        "Pi Lanz 500ml", "Pi Multi Micro 25kg", "Pi Multi Trace 25kg", "Pi Vita 200ml",
        "Pyra Gold 250gms", "Pyra Gold 500gms"
    ],
    "Srikar": [
        "Prime Granules 5kg", "Ripkil Top 250ml", "Ripkil Top 500ml", "Sixer 250ml",
        "Sixer 500ml"
    ],
    "Anjaneya": [
        "Acemain 1kg", "Acemain 250gms", "Acemain 500gms", "Argyle 250gms", "Bahaar 1 Lt",
        "Bahaar 500ml", "Bomber 200ml", "Bulldo 250ml", "Bulldo 500ml", "Chooper 500ml",
        "Dorito 250gms", "Dorito 500gms", "Ekka 500ml", "Green Stone 250ml", "Green Stone 500ml",
        "Insta 1 Lt", "Insta 500ml", "Kabuto 250ml", "Kabuto 500ml", "Kabuto Plus 500ml",
        "Lancer 250ml", "Lancer 500ml", "Mastak 250gms", "Mastak 500gms", "Meto 1 Lt",
        "Meto 500ml", "Monocil 250ml", "Monocil 500ml", "Nisarga 500ml", "Power Point 250ml",
        "Power Point 500ml", "Ruler 1 Lt", "Ruler 500ml", "Samrudhi 500ml", "Seimlet 240gms",
        "Semarai 1 Lt", "Serang 500gms", "Suraksha 250ml", "Super Lava 5g 250ml",
        "Super Max 500gms", "Zenetra 1 Lt", "Zenetra 250ml", "Zenetra 500ml"
    ],
    "Sairam": [
        "Amistar 1 Lt", "Amistar 500ml", "Bavistin 500gms", "Calaris Xtra 1.4lt",
        "Cythion 1 Lt", "Cythion 500ml", "Exel Mera 1 Lt", "Karait 1 Lt", "Karait 500ml",
        "Kasumin 1 Lt", "Kasumin 500ml", "Monocil[e] 250ml", "Monocil[e] 500ml",
        "Nativo 300gms", "Providar 100ml", "Providar 250ml", "Providar 500ml",
        "Sulpher 90% Wp 1kg", "Taqat 1 Lt", "Taqat 250ml", "Taqat 500ml", "Trassid 500ml"
    ],
    "T Stanes": [
        "Biocure 1kg", "Biocure-b 1kg", "Fytovita 1 Lt", "Fytovita 500ml", "Kurax 1 Lt",
        "Kurax 500ml", "S.o.p 1kg", "S.o.p 25kg", "Stanomax 1 Lt", "Stanomax 250ml",
        "Stanomax 500ml", "Stenomin 500gms", "Super Lava 5g 1 Lt", "Super Lava 5g 250ml",
        "Super Max 500gms"
    ],
    "Syngenta": [
        "Actara 100gms", "Actara 250gms", "Amistar 1 Lt", "Capcadis 250gms", "Capcadis 50gms",
        "Curacron 500ml", "Pegasus 500gms", "Quantis 400ml", "Redomil Gold 100gms",
        "Redomil Gold 250gms", "Redomil Gold 500gms"
    ]
}

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
            
            # Top company
            if not company_summary.empty:
                top_company = company_summary.index[0]
                top_company_revenue = company_summary.iloc[0]['Total Revenue']
                insights.append(f"Top Company: {top_company} (₹{top_company_revenue:,.2f})")

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
