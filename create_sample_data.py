#!/usr/bin/env python3
"""
Sample Data Generator for Excel Field Analyzer
Creates a sample Excel file with multiple worksheets for testing purposes.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

def create_sample_excel():
    """Create a sample Excel file with multiple worksheets for testing."""
    
    # Create output filename
    output_file = "sample_data.xlsx"
    
    # Sample data for different worksheets
    worksheets = {
        "Orders": {
            "Purchase Order": [f"PO-{i:04d}" for i in range(1, 101)],
            "Order Details": [f"Order {i}" for i in range(1, 101)],
            "Due Date": [(datetime.now() + timedelta(days=random.randint(1, 30))).strftime("%Y-%m-%d") for _ in range(100)],
            "Product": [f"Product {chr(65 + i % 26)}{i//26 + 1}" for i in range(100)],
            "Description": [f"Description for product {i+1}" for i in range(100)],
            "Quantity": [random.randint(1, 50) for _ in range(100)],
            "Unit Price": [round(random.uniform(10, 500), 2) for _ in range(100)],
            "Total Price": [round(random.uniform(100, 5000), 2) for _ in range(100)],
            "Customer": [f"Customer {i+1}" for i in range(100)],
            "Shipping Code": [f"SHIP-{random.randint(1000, 9999)}" for _ in range(100)]
        },
        
        "Production": {
            "Purchase Order": [f"PO-{i:04d}" for i in range(1, 101)],
            "Product": [f"Product {chr(65 + i % 26)}{i//26 + 1}" for i in range(100)],
            "Build Time": [random.randint(1, 8) for _ in range(100)],
            "Cut Time": [random.randint(1, 4) for _ in range(100)],
            "Man Mins": [random.randint(30, 480) for _ in range(100)],
            "Total Man Mins": [random.randint(60, 960) for _ in range(100)],
            "Built By": [f"Worker {random.randint(1, 10)}" for _ in range(100)],
            "Build Information": [f"Build info {i+1}" for i in range(100)],
            "Production Date": [(datetime.now() + timedelta(days=random.randint(-30, 0))).strftime("%Y-%m-%d") for _ in range(100)],
            "Status": [random.choice(["In Progress", "Completed", "Pending"]) for _ in range(100)]
        },
        
        "Shipping": {
            "Purchase Order": [f"PO-{i:04d}" for i in range(1, 101)],
            "Product": [f"Product {chr(65 + i % 26)}{i//26 + 1}" for i in range(100)],
            "Shipping Code": [f"SHIP-{random.randint(1000, 9999)}" for _ in range(100)],
            "Carrier": [random.choice(["APC", "DX", "Van", "Royal Mail"]) for _ in range(100)],
            "Tracking Number": [f"TRK{random.randint(100000, 999999)}" for _ in range(100)],
            "Ship Date": [(datetime.now() + timedelta(days=random.randint(-15, 0))).strftime("%Y-%m-%d") for _ in range(100)],
            "Delivery Date": [(datetime.now() + timedelta(days=random.randint(1, 7))).strftime("%Y-%m-%d") for _ in range(100)],
            "Customer": [f"Customer {i+1}" for i in range(100)],
            "Address": [f"Address {i+1}, City, Postcode" for i in range(100)]
        },
        
        "Inventory": {
            "Product": [f"Product {chr(65 + i % 26)}{i//26 + 1}" for i in range(100)],
            "Description": [f"Description for product {i+1}" for i in range(100)],
            "Category": [random.choice(["Electronics", "Clothing", "Books", "Home"]) for _ in range(100)],
            "Stock Level": [random.randint(0, 1000) for _ in range(100)],
            "Reorder Point": [random.randint(10, 100) for _ in range(100)],
            "Unit Cost": [round(random.uniform(5, 200), 2) for _ in range(100)],
            "Supplier": [f"Supplier {random.randint(1, 20)}" for _ in range(100)],
            "Last Updated": [(datetime.now() + timedelta(days=random.randint(-30, 0))).strftime("%Y-%m-%d") for _ in range(100)]
        },
        
        "Customers": {
            "Customer ID": [f"CUST-{i:04d}" for i in range(1, 101)],
            "Customer Name": [f"Customer {i+1}" for i in range(100)],
            "Email": [f"customer{i+1}@example.com" for i in range(100)],
            "Phone": [f"+44 {random.randint(100000000, 999999999)}" for _ in range(100)],
            "Address": [f"Address {i+1}, City, Postcode" for i in range(100)],
            "Registration Date": [(datetime.now() + timedelta(days=random.randint(-365, 0))).strftime("%Y-%m-%d") for _ in range(100)],
            "Total Orders": [random.randint(1, 50) for _ in range(100)],
            "Total Spent": [round(random.uniform(100, 10000), 2) for _ in range(100)]
        }
    }
    
    # Create Excel writer
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, data in worksheets.items():
            df = pd.DataFrame(data)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"✅ Sample Excel file created: {output_file}")
    print(f"📊 Contains {len(worksheets)} worksheets:")
    for sheet_name in worksheets.keys():
        print(f"   • {sheet_name}")
    print(f"\n🎯 You can now test the analyzer with:")
    print(f"   python excel_analyzer_cli.py {output_file}")

if __name__ == "__main__":
    create_sample_excel() 