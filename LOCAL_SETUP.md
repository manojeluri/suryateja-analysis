# Local Setup Guide

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Prepare Your Data File

Place your Excel (.xlsx, .xls) or CSV (.csv) file in this folder. The file should have these columns:

**Required columns:**
- `ITNAME` - Product name
- `QTY` - Quantity sold
- `TAXBLEAMT` (or `NAMT`) - Taxable amount
- `GST` (or `PER`) - GST percentage

**Optional columns:**
- `HSNCODE` - HSN code

### 3. Run the Analysis

```bash
# If your file is named data.xlsx, sales_data.xlsx, data.csv, or sales_data.csv:
python run_local.py

# Or specify your file explicitly:
python run_local.py your_file.xlsx
```

The script will generate a PDF report named `Sales_Analysis_YYYYMMDD_HHMMSS.pdf` in the same folder.

## Example

```bash
# Install dependencies
pip install -r requirements.txt

# Run with your data file
python run_local.py sales_data.xlsx

# Output:
# ============================================================
# Sales Analysis - Local Runner
# ============================================================
#
# 📁 Loading data from: sales_data.xlsx
# ✅ Loaded 150 records
#
# 🔄 Processing data and generating PDF report...
#
# ✅ SUCCESS! PDF report generated
# 📄 Saved as: Sales_Analysis_20251125_143022.pdf
# 📊 Summary:
#    Total Products: 150
#    Total Quantity: 5,432
#    Total Revenue: Rs.234,567.89
#    Total GST: Rs.42,222.22
#    Grand Total: Rs.276,790.11
#    Companies: 12
```

## What Gets Generated

The PDF report includes:

1. **Executive Summary Page**
   - Total products, quantity, revenue, GST
   - Top performers by revenue and quantity
   - Top company by market share

2. **Top 15 Products by Revenue**
   - Colorful bar chart showing highest revenue products

3. **Top 15 Products by Quantity**
   - Visualization of most sold products

4. **Top 15 Companies by Revenue**
   - Company-wise revenue breakdown with values

## Data File Format Examples

### Excel Format (.xlsx)
```
| HSNCODE   | ITNAME        | QTY | TAXBLEAMT | GST |
|-----------|---------------|-----|-----------|-----|
| 38089199  | Agas 250gms   | 1   | 600       | 18  |
| 38089199  | Alecto 50 Ml  | 3   | 4440      | 18  |
```

### CSV Format (.csv)
```csv
HSNCODE,ITNAME,QTY,TAXBLEAMT,GST
38089199,Agas 250gms,1,600,18
38089199,Alecto 50 Ml,3,4440,18
```

## Troubleshooting

### Error: Missing required columns
**Problem:** Your data file doesn't have the expected column names.

**Solution:** The script accepts alternative column names:
- `NAMT` instead of `TAXBLEAMT`
- `PER` instead of `GST`

Make sure your columns match one of these naming conventions.

### Error: Module not found
**Problem:** Dependencies not installed.

**Solution:**
```bash
pip install -r requirements.txt
```

### Error: File not found
**Problem:** Specified file doesn't exist.

**Solution:**
```bash
# Check your current directory
ls -la

# Make sure file path is correct
python run_local.py /full/path/to/your/file.xlsx
```

## Testing with Sample Data

Use the included test script to verify everything works:

```bash
python test_local.py
```

This will use sample data and generate a test PDF.

## Company Product Mapping

The analysis automatically categorizes products by company using CSV files in the `Company Wise Products` folder. Currently includes:

- Adama
- Anjaneya
- Balaji
- Best Agrolife
- Chennakesava
- Coramandel
- Dhanuka
- Gharda
- Godrej
- Indofil
- Nichino
- Nova Agri Science
- PI Industries
- Rallis
- Sairam
- Srikar
- Sudharsan
- Superior
- Swal
- Syngenta
- T Stanes
- VNR

Products not in these lists are categorized as "Other".

## Advanced Usage

### Process Multiple Files

```bash
for file in *.xlsx; do
    python run_local.py "$file"
done
```

### Custom Output Location

Modify `run_local.py` to change the output directory:

```python
output_filename = f"./reports/Sales_Analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
```

## Requirements

- Python 3.8+
- pandas
- matplotlib
- seaborn
- openpyxl (for Excel files)
- flask
- numpy

All requirements are in `requirements.txt`.
