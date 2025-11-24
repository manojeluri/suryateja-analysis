"""
Local runner for Sales Analysis
Processes Excel/CSV data and generates PDF report
"""

import sys
import os
import pandas as pd
import json
from datetime import datetime

# Add api directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'api'))
from analyze import analyze_data

def load_data_from_file(filepath):
    """Load data from Excel or CSV file"""
    ext = os.path.splitext(filepath)[1].lower()

    if ext == '.xlsx':
        df = pd.read_excel(filepath, engine='openpyxl')
    elif ext == '.xls':
        df = pd.read_excel(filepath, engine='xlrd')
    elif ext == '.csv':
        df = pd.read_csv(filepath)
    else:
        raise ValueError(f"Unsupported file type: {ext}. Please use .xlsx, .xls, or .csv")

    # Convert DataFrame to list of dicts
    data = df.to_dict('records')
    return data

def main():
    print("=" * 60)
    print("Sales Analysis - Local Runner")
    print("=" * 60)

    # Get data file path from command line or use default
    if len(sys.argv) > 1:
        data_file = sys.argv[1]
    else:
        # Look for common file names in current directory
        possible_files = ['data.xlsx', 'sales_data.xlsx', 'data.csv', 'sales_data.csv']
        data_file = None
        for f in possible_files:
            if os.path.exists(f):
                data_file = f
                break

        if not data_file:
            print("\nUsage: python run_local.py <path_to_data_file>")
            print("\nSupported formats: .xlsx, .xls, .csv")
            print("\nExpected columns:")
            print("  - ITNAME (or Item Name) - Product name")
            print("  - QTY (or Quantity) - Quantity sold")
            print("  - TAXBLEAMT (or NAMT) - Taxable amount")
            print("  - GST (or PER) - GST percentage")
            print("  - HSNCODE (optional) - HSN code")
            sys.exit(1)

    # Check if file exists
    if not os.path.exists(data_file):
        print(f"\n❌ Error: File not found: {data_file}")
        sys.exit(1)

    print(f"\n📁 Loading data from: {data_file}")

    try:
        # Load data
        data = load_data_from_file(data_file)
        print(f"✅ Loaded {len(data)} records")

        # Show first few columns to verify
        if data:
            print(f"\n📋 Available columns: {list(data[0].keys())}")
            print(f"📊 First record sample:")
            for key, value in list(data[0].items())[:5]:
                print(f"   {key}: {value}")

        # Analyze data
        print(f"\n🔄 Processing data and generating PDF report...")
        result = analyze_data(data)

        if result['success']:
            # Save PDF
            output_filename = f"Sales_Analysis_{datetime.now().strftime('%Y_%m_%d')}.pdf"
            with open(output_filename, 'wb') as f:
                f.write(result['pdf_bytes'])

            print(f"\n✅ SUCCESS! PDF report generated")
            print(f"📄 Saved as: {output_filename}")
            print(f"📊 Summary:")
            print(f"   Total Products: {result['summary']['total_rows']}")
            print(f"   Total Quantity: {result['summary']['total_quantity']:,}")
            print(f"   Total Revenue: Rs.{result['summary']['total_revenue']:,.2f}")
            print(f"   Total GST: Rs.{result['summary']['total_gst']:,.2f}")
            print(f"   Grand Total: Rs.{result['summary']['grand_total']:,.2f}")
            print(f"   Companies: {result['summary']['total_companies']}")
            print(f"\n✨ Report ready! Open {output_filename} to view.")

        else:
            print(f"\n❌ Analysis failed!")
            print(f"Error: {result['error']}")
            if 'traceback' in result:
                print(f"\nTraceback:\n{result['traceback']}")
            if 'available_columns' in result:
                print(f"\nAvailable columns in your file: {result['available_columns']}")

    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
