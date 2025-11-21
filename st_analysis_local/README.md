# Enhanced Sales Data Analyzer

A comprehensive local sales analysis tool that reads Excel files and generates detailed analytics with beautiful visualizations and HTML reports.

## Features

### 📊 Advanced Analytics
- **Company-wise Analysis**: Revenue, market share, product count, and pricing metrics
- **Product Performance**: Rankings by revenue, quantity, and performance scores
- **GST Analysis**: Tax collection breakdown and distribution
- **Performance Scoring**: Normalized metrics combining revenue and quantity
- **Statistical Insights**: Mean, median, standard deviation, and percentile rankings

### 📈 Visualizations (10 Charts)
1. Top 15 Products by Revenue
2. Top 15 Products by Quantity
3. Company Revenue Distribution
4. Company Market Share (Pie Chart)
5. Product Count by Company
6. GST Collection by Tax Rate
7. Revenue vs Quantity Scatter Plot
8. Price Per Unit Distribution
9. Top 20 Products by Performance Score
10. Average Price Per Unit by Company

### 📄 Outputs
- **HTML Report**: Comprehensive interactive report with all metrics and charts
- **Excel Summary**: Multi-sheet workbook with detailed analyses
- **PNG Charts**: High-resolution visualizations (300 DPI)

## Installation

### 1. Install Python Dependencies

```bash
pip install -r requirements.txt
```

### 2. Place Your Data File

Copy your Excel file (e.g., `SALANAL.XLS`) into this directory:

```
st_analysis_local/
├── local_analyzer.py
├── requirements.txt
├── README.md
└── SALANAL.XLS  <-- Your data file here
```

## Usage

### Run the Analysis

```bash
python local_analyzer.py
```

The script will:
1. 🔍 Automatically detect .xls or .xlsx files
2. 📊 Perform comprehensive analysis
3. 📈 Generate 10 visualizations
4. 📄 Create HTML report
5. 💾 Export Excel summary

### View Results

All outputs are saved in the `outputs/` folder:

```
outputs/
├── Sales_Analysis_Report.html  <-- Open this in browser
├── Analysis_Summary.xlsx
└── 01-10 visualization PNG files
```

## Data Format

The script expects an Excel file with the following columns:

| Column     | Description                          |
|------------|--------------------------------------|
| HSNCODE    | HSN product code                     |
| ITNAME     | Product name                         |
| QTY        | Quantity sold                        |
| TAXBLEAMT  | Taxable amount (pre-GST revenue)     |
| GST        | GST rate (%)                         |

## Company Mappings

The analyzer includes pre-configured mappings for 22 companies:

- Gharda, Adama, Best Agrolife, Godrej, Indofil
- Nichino, Nova Agri Science, Nova Agri Tech, Superior
- Swal, VNR, Balaji, Chennakesava, PI
- Srikar, Anjaneya, Sairam, T Stanes, Syngenta
- And more...

Products not in the mapping are categorized as "Other".

## Key Metrics Calculated

### Basic Metrics
- Total revenue, quantity, and GST
- Average price per unit
- Average order value

### Advanced Metrics
- **Performance Score**: Weighted combination (60% revenue + 40% quantity)
- **Market Share**: Company-wise revenue percentage
- **Percentile Rankings**: Revenue and quantity rankings
- **Price Distribution**: Statistical analysis of pricing

### Company Analysis
- Total revenue and market share
- Product count and diversity
- Average pricing and quantity metrics
- Standard deviation for consistency analysis

## Customization

### Add New Companies

Edit the `COMPANY_PRODUCTS` dictionary in `local_analyzer.py`:

```python
COMPANY_PRODUCTS = {
    "Your Company": [
        "Product 1",
        "Product 2",
        ...
    ],
    ...
}
```

### Adjust Top N Rankings

Change the `top_n` parameter in the `identify_star_performers()` function:

```python
performers = identify_star_performers(df, top_n=20)  # Show top 20 instead of 10
```

### Modify Visualizations

Edit the `create_all_visualizations()` function to:
- Change colors: Modify `plt.cm.ColorMap`
- Adjust sizes: Change `figsize` parameters
- Add new charts: Create new visualization blocks

## Troubleshooting

### File Not Found Error
**Problem**: "No .xls or .xlsx files found"

**Solution**: Ensure your Excel file is in the same directory as `local_analyzer.py`

### Import Errors
**Problem**: "ModuleNotFoundError: No module named 'pandas'"

**Solution**: Install dependencies:
```bash
pip install -r requirements.txt
```

### Excel Read Error
**Problem**: "Failed to read Excel file"

**Solution**: The script tries both `xlrd` and `openpyxl` engines. Make sure both are installed:
```bash
pip install xlrd openpyxl
```

For newer .xlsx files, use:
```bash
pip install openpyxl
```

For older .xls files, use:
```bash
pip install xlrd==2.0.1
```

## Output Examples

### Console Output
```
====================================================================
🚀 ENHANCED SALES DATA ANALYZER
====================================================================

📂 Loading data from: SALANAL.XLS
✅ Loaded 150 rows and 5 columns
📋 Columns: HSNCODE, ITNAME, QTY, TAXBLEAMT, GST

🏷️  Categorizing products by company...
📊 Calculating advanced metrics...
🔍 Generating company analysis...
🔍 Generating product analysis...
🔍 Generating GST analysis...
🔍 Identifying star performers...

📈 Creating visualizations...
✅ Generated 10 visualizations in 'outputs/' directory

📄 Generating comprehensive HTML report...
✅ HTML report generated: outputs/Sales_Analysis_Report.html

====================================================================
✅ ANALYSIS COMPLETE!
====================================================================

📊 Results Summary:
   • HTML Report: outputs/Sales_Analysis_Report.html
   • Excel Summary: outputs/Analysis_Summary.xlsx
   • Visualizations: 10 charts in outputs/ folder

💡 Open 'outputs/Sales_Analysis_Report.html' in your browser to view the full report!
```

## Comparison with Production Version

### Improvements Over API Version:

1. **Local File Support**: No need for API calls or JSON formatting
2. **Enhanced Analytics**:
   - Performance scoring system
   - Percentile rankings
   - Statistical analysis
   - Price distribution metrics
3. **Better Visualizations**: 10 charts vs 4 charts
4. **HTML Report**: Professional, interactive report
5. **Excel Export**: Multi-sheet summary workbook
6. **No Dependencies**: No Flask server needed

## Requirements

- Python 3.8+
- pandas 2.1+
- matplotlib 3.8+
- seaborn 0.13+
- numpy 1.26+
- openpyxl 3.1+ (for .xlsx files)
- xlrd 2.0+ (for .xls files)

## License

Free to use and modify for your business needs.

## Support

For issues or questions:
1. Check the Troubleshooting section
2. Review the console output for error messages
3. Ensure your data file matches the expected format

---

**Happy Analyzing! 📊📈**
