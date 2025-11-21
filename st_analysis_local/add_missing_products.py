import csv
import os

# Define company name mappings
company_mapping = {
    'Syngenta': 'Syngenta_Products.csv',
    'Nova science': 'NovaAgriScience_Products.csv',
    'Swal': 'Swal_Products.csv',
    'Rallis': 'Rallis_Products.csv',
    'Superior': 'Superior_Products.csv',
    'Godrej': 'Godrej_Products.csv',
    'Bestagro': 'BestAgrolife_Products.csv',
    'Tstanes': 'T_Stanes_Products.csv',
    'Balaji': 'Balaji_Products.csv',
    'Adama': 'Adama_Products.csv',
    'VNR': 'VNR_Products.csv',
    'Gharda': 'Gharda_Products.csv',
    'PI': 'PI_Products.csv',
    'Chennakesava': 'Chennakesava_Products.csv',
    'Indofil': 'Indofil_Products.csv',
    'Nichino': 'Nichino_Products.csv',
    'Nova tech': 'Nova_Agri_TechProduct_Names.csv',
}

# Read the missing products CSV
missing_products = []
with open('/Users/manojeluri/Downloads/Missing Products Companies - missing_products (1).csv', 'r') as f:
    reader = csv.DictReader(f)
    for row in reader:
        missing_products.append(row)

# Organize products by company
products_by_company = {}
products_not_added = []

for product in missing_products:
    product_name = product['ITNAME']
    company = product['Company']

    if company in company_mapping:
        if company not in products_by_company:
            products_by_company[company] = []
        products_by_company[company].append(product_name)
    else:
        products_not_added.append({
            'product': product_name,
            'company': company,
            'reason': 'Company file not found'
        })

# Add products to respective files
base_path = '/Users/manojeluri/Desktop/root/Work/Build/AI Coding/st_analysis_local/Company Wise Products/'
added_count = 0
duplicate_count = 0

for company, products in products_by_company.items():
    file_path = os.path.join(base_path, company_mapping[company])

    # Read existing products
    existing_products = set()
    with open(file_path, 'r') as f:
        reader = csv.reader(f)
        next(reader)  # Skip header
        for row in reader:
            if row and row[0]:
                existing_products.add(row[0].strip())

    # Add new products
    new_products = []
    for product in products:
        if product not in existing_products:
            new_products.append(product)
            added_count += 1
        else:
            duplicate_count += 1
            products_not_added.append({
                'product': product,
                'company': company,
                'reason': 'Already exists in file'
            })

    # Append new products to file
    if new_products:
        with open(file_path, 'a', newline='') as f:
            writer = csv.writer(f)
            for product in new_products:
                writer.writerow([product])

# Print summary
print(f"\n{'='*60}")
print(f"SUMMARY")
print(f"{'='*60}")
print(f"✓ Products successfully added: {added_count}")
print(f"⊗ Products already existing (duplicates): {duplicate_count}")
print(f"✗ Products not added: {len(products_not_added)}")
print(f"{'='*60}\n")

if products_not_added:
    print("Products that couldn't be added:\n")

    # Group by reason
    by_reason = {}
    for item in products_not_added:
        reason = item['reason']
        if reason not in by_reason:
            by_reason[reason] = []
        by_reason[reason].append(item)

    for reason, items in by_reason.items():
        print(f"\n{reason}:")
        print("-" * 60)
        for item in items:
            company_display = f"({item['company']})" if item['company'] else "(No company)"
            print(f"  • {item['product']} {company_display}")
