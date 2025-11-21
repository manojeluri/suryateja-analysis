"""
Quick test script to verify the analyze function works
"""
import json

# Sample test data matching your n8n workflow format
test_data = [
    {"HSNCODE":"38089199","ITNAME":"Agas 250gms","QTY":1,"TAXBLEAMT":600,"GST":18},
    {"HSNCODE":"38089199","ITNAME":"Alecto 50 Ml","QTY":3,"TAXBLEAMT":4440,"GST":18},
    {"HSNCODE":"38089199","ITNAME":"Antracol 1kg","QTY":1,"TAXBLEAMT":800,"GST":18}
]

print("Testing analyze function...")
print(f"Test data has {len(test_data)} items")

try:
    # Import the module
    import sys
    sys.path.insert(0, 'api')
    from analyze import analyze_data
    
    result = analyze_data(test_data)
    
    if result['success']:
        print("✅ Analysis SUCCESS!")
        print(f"   - Total rows: {result['summary']['total_rows']}")
        print(f"   - Total revenue: Rs.{result['summary']['total_revenue']:.2f}")
        print(f"   - PDF bytes: {len(result['pdf_bytes'])} bytes")
        print(f"   - Filename: {result['filename']}")
    else:
        print("❌ Analysis FAILED!")
        print(f"   Error: {result['error']}")
        if 'traceback' in result:
            print(f"   Traceback: {result['traceback']}")
except Exception as e:
    print(f"❌ Test FAILED with exception: {str(e)}")
    import traceback
    traceback.print_exc()
