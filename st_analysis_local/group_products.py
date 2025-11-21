"""
Script to group products with same name but different packages together
"""
import pandas as pd
import os
import re
from collections import defaultdict

def extract_base_product_name(product_name):
    """
    Extract base product name by removing package size information
    Examples:
        "Aniloguard 1 Lt" -> "Aniloguard"
        "Bakeel 250ml" -> "Bakeel"
        "Guru 500gms" -> "Guru"
    """
    # Remove package size patterns (numbers followed by units)
    # Common patterns: ml, Lt, L, kg, gms, gr, Gms
    pattern = r'\s*\d+[\.]?\d*\s*(ml|ML|Lt|LT|L|kg|KG|gms|Gms|GMS|gr|GR|Gr)\s*$'
    base_name = re.sub(pattern, '', product_name).strip()
    return base_name

def extract_package_size(product_name):
    """
    Extract package size for sorting
    Returns a tuple (numeric_value, unit) for proper sorting
    """
    # Find package size pattern
    pattern = r'(\d+[\.]?\d*)\s*(ml|ML|Lt|LT|L|kg|KG|gms|Gms|GMS|gr|GR|Gr)$'
    match = re.search(pattern, product_name)

    if match:
        value = float(match.group(1))
        unit = match.group(2).lower()

        # Normalize units to ml/gms for comparison
        if unit in ['lt', 'l']:
            value = value * 1000  # Convert to ml
            unit = 'ml'
        elif unit in ['kg']:
            value = value * 1000  # Convert to gms
            unit = 'gms'

        return (value, unit)

    # If no package size found, return high value to put at end
    return (999999, 'zzz')

def group_and_sort_products(csv_file_path):
    """
    Read CSV, group products by base name, sort by package size, and save back
    """
    print(f"\n📂 Processing: {os.path.basename(csv_file_path)}")

    # Read CSV
    df = pd.read_csv(csv_file_path)

    if 'Product Name' not in df.columns:
        print(f"   ⚠️  Warning: 'Product Name' column not found")
        return

    # Get product names and remove empty/NaN values
    products = df['Product Name'].dropna().tolist()
    original_count = len(products)

    # Remove duplicates while preserving order
    seen = set()
    unique_products = []
    duplicates_removed = 0

    for product in products:
        product = str(product).strip()
        if product and product not in seen:
            seen.add(product)
            unique_products.append(product)
        elif product in seen:
            duplicates_removed += 1

    # Group products by base name
    grouped = defaultdict(list)
    for product in unique_products:
        base_name = extract_base_product_name(product)
        grouped[base_name].append(product)

    # Sort within each group by package size
    sorted_products = []
    for base_name in sorted(grouped.keys()):
        group = grouped[base_name]
        # Sort by package size (smallest to largest)
        group_sorted = sorted(group, key=extract_package_size)
        sorted_products.extend(group_sorted)

    # Create new dataframe
    new_df = pd.DataFrame({'Product Name': sorted_products})

    # Save back to CSV
    new_df.to_csv(csv_file_path, index=False)

    print(f"   ✅ Processed {original_count} products")
    if duplicates_removed > 0:
        print(f"   🔄 Removed {duplicates_removed} duplicates")
    print(f"   📦 Grouped into {len(grouped)} product families")
    print(f"   💾 Saved {len(sorted_products)} unique products")

def main():
    """Process all CSV files in Company Wise Products folder"""

    folder_path = "Company Wise Products"

    if not os.path.exists(folder_path):
        print(f"❌ Folder '{folder_path}' not found!")
        return

    # Get all CSV files
    csv_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.csv')]

    print("="*60)
    print("🔄 GROUPING PRODUCTS BY NAME")
    print("="*60)
    print(f"\n📁 Found {len(csv_files)} CSV files in '{folder_path}/'")

    # Process each file
    for csv_file in sorted(csv_files):
        file_path = os.path.join(folder_path, csv_file)
        try:
            group_and_sort_products(file_path)
        except Exception as e:
            print(f"   ❌ Error: {str(e)}")

    print("\n" + "="*60)
    print("✅ ALL FILES PROCESSED!")
    print("="*60)
    print("\n💡 Products with the same name are now grouped together")
    print("   and sorted by package size (smallest to largest)")

if __name__ == "__main__":
    main()
