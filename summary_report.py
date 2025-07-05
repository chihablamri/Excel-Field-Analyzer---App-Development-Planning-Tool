#!/usr/bin/env python3
"""
Summary Report Generator for Excel Field Analysis
This script provides a clear summary of the field analysis results.
"""

import json
import pandas as pd
from pathlib import Path

def load_analysis_report():
    """Load the analysis report from JSON file."""
    report_file = Path("excel_analysis_results_improved/improved_analysis_report.json")
    if report_file.exists():
        with open(report_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        print("Analysis report not found. Please run the Excel field analyzer first.")
        return None

def print_summary_report():
    """Print a comprehensive summary report."""
    report = load_analysis_report()
    if not report:
        return
    
    print("=" * 80)
    print("PRODUCTION SCHEDULE FIELD ANALYSIS - SUMMARY REPORT")
    print("=" * 80)
    print(f"File Analyzed: {report['file_path']}")
    print(f"Analysis Date: {report['analysis_date']}")
    print()
    
    # Key Statistics
    print("📊 KEY STATISTICS:")
    print("-" * 40)
    print(f"• Total Worksheets: {report['total_sheets']}")
    print(f"• Total Unique Fields: {report['total_unique_fields']}")
    print(f"• Fields Used in Multiple Sheets: {len(report['common_fields'])}")
    print(f"• Fields Used in Single Sheet: {len(report['unique_fields'])}")
    print(f"• Universal Fields (All Sheets): {len(report['universal_fields'])}")
    print()
    
    # Most Common Fields
    print("🏆 MOST COMMON FIELDS (Used in 7+ sheets):")
    print("-" * 40)
    sorted_common = sorted(report['common_fields'].items(), key=lambda x: x[1], reverse=True)
    for field, count in sorted_common:
        if count >= 7:
            print(f"• {field} ({count} sheets)")
    print()
    
    # Field Categories
    print("📋 FIELD CATEGORIES:")
    print("-" * 40)
    for category, fields in report['field_categories'].items():
        if fields:
            print(f"\n{category} ({len(fields)} fields):")
            # Sort fields by usage
            sorted_fields = sorted(fields, key=lambda f: report['sheets_per_field'].get(f, 0), reverse=True)
            for field in sorted_fields[:5]:  # Show top 5 per category
                sheet_count = report['sheets_per_field'].get(field, 0)
                print(f"  • {field} ({sheet_count} sheets)")
            if len(sorted_fields) > 5:
                print(f"  • ... and {len(sorted_fields) - 5} more")
    print()
    
    # Sheet Analysis
    print("📄 WORKSHEET ANALYSIS:")
    print("-" * 40)
    for i, sheet_name in enumerate(report['sheet_names'], 1):
        field_count = report['fields_per_sheet'][sheet_name]
        print(f"{i:2d}. {sheet_name} ({field_count} fields)")
    print()
    
    # Recommendations
    print("💡 RECOMMENDATIONS FOR APP DEVELOPMENT:")
    print("-" * 40)
    print("1. CORE FIELDS (Include in all modules):")
    core_fields = [f for f, count in report['common_fields'].items() if count >= 7]
    for field in core_fields[:10]:  # Top 10 core fields
        print(f"   • {field}")
    print()
    
    print("2. MODULE-SPECIFIC FIELDS:")
    print("   • Order Management: Purchase Order, Order Details, Due Date")
    print("   • Production: Build Time, Cut Time, Man Mins, Total Man Mins")
    print("   • Product Info: Product, Description, Shipping Code")
    print("   • Build Tracking: Built By, Build Information, Mins Built")
    print("   • Despatch: APC, DX, Van, Label Printed?")
    print()
    
    print("3. DATA STRUCTURE SUGGESTIONS:")
    print("   • Use a flexible schema to accommodate varying field sets per sheet")
    print("   • Implement field mapping for different worksheet types")
    print("   • Consider dynamic form generation based on sheet type")
    print("   • Include field validation based on usage patterns")
    print()
    
    print("=" * 80)
    print("📁 OUTPUT FILES GENERATED:")
    print("-" * 40)
    print("• improved_field_matrix.xlsx - Complete field presence matrix")
    print("• improved_detailed_analysis.xlsx - Detailed analysis with categories")
    print("• improved_analysis_report.json - Raw analysis data")
    print()
    print("Use these files to guide your app development and database design.")
    print("=" * 80)

if __name__ == "__main__":
    print_summary_report() 