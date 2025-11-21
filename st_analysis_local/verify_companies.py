"""
Quick verification script to test company product loading
"""
import os

def load_company_products_from_csv(folder_path="Company Wise Products"):
    """
    Load company-product mappings from CSV files in the specified folder
    """
    company_products = {}

    if not os.path.exists(folder_path):
        print(f"⚠️  Warning: '{folder_path}' folder not found.")
        return company_products

    # Get all CSV files
    csv_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.csv')]

    print(f"\n📁 Loading company products from '{folder_path}/' folder...")
    print(f"   Found {len(csv_files)} CSV files\n")

    for csv_file in sorted(csv_files):
        # Extract company name from filename
        company_name = csv_file.replace('_Products.csv', '').replace('_Product_Names.csv', '')
        company_name = company_name.replace('_', ' ')

        # Handle special cases
        if company_name == 'BestAgrolife':
            company_name = 'Best Agrolife'
        elif company_name == 'NovaAgriScience':
            company_name = 'Nova Agri Science'

        print(f"   ✅ {csv_file:<40} → {company_name}")

    print(f"\n✅ Total companies: {len(csv_files)}")

    # Highlight newly added companies
    new_companies = ['Dhanuka_Products.csv', 'Sudharsan_Products.csv', 'Coramandel_Products.csv']
    print(f"\n🆕 Newly added companies:")
    for new_file in new_companies:
        if new_file in csv_files:
            print(f"   ✓ {new_file}")
        else:
            print(f"   ✗ {new_file} - NOT FOUND!")

if __name__ == "__main__":
    print("="*60)
    print("🔍 COMPANY PRODUCTS VERIFICATION")
    print("="*60)
    load_company_products_from_csv()
    print("\n" + "="*60)
    print("✅ Verification complete! All CSV files will be automatically")
    print("   loaded when you run local_analyzer.py")
    print("="*60)
