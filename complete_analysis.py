#!/usr/bin/env python3
"""
Complete Excel Analysis Workflow
Performs full analysis: Excel file → Field Matrix → Comprehensive Report → Open Report
"""

import subprocess
import sys
from pathlib import Path
import time

def run_command(command, description):
    """Run a command and handle errors."""
    print(f"\n{'='*60}")
    print(f"🚀 {description}")
    print(f"{'='*60}")
    print(f"Running: {command}")
    
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        print("✅ Success!")
        if result.stdout:
            print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Error: {e}")
        if e.stdout:
            print("STDOUT:", e.stdout)
        if e.stderr:
            print("STDERR:", e.stderr)
        return False

def main():
    """Run the complete analysis workflow."""
    print("🎯 COMPLETE EXCEL ANALYSIS WORKFLOW")
    print("="*60)
    
    # Check if Excel file is provided
    if len(sys.argv) < 2:
        print("❌ Please provide an Excel file path!")
        print("Usage: python complete_analysis.py \"path/to/your/file.xlsx\"")
        print("\nExample:")
        print("  python complete_analysis.py \"Editing Production Schedule - MW Version.xlsx\"")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    
    # Check if file exists
    if not Path(excel_file).exists():
        print(f"❌ File not found: {excel_file}")
        sys.exit(1)
    
    print(f"📁 Analyzing: {excel_file}")
    print(f"⏰ Started at: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Step 1: Run Excel field analysis
    if not run_command(f'python excel_analyzer_cli.py "{excel_file}"', 
                      "STEP 1: Analyzing Excel file and generating field matrix"):
        print("❌ Step 1 failed. Stopping workflow.")
        sys.exit(1)
    
    # Step 2: Generate comprehensive report
    if not run_command('python generate_comprehensive_report.py', 
                      "STEP 2: Generating comprehensive report with charts"):
        print("❌ Step 2 failed. Stopping workflow.")
        sys.exit(1)
    
    # Step 3: Open the report
    if not run_command('python open_report.py', 
                      "STEP 3: Opening HTML report in browser"):
        print("⚠️  Step 3 failed, but reports are still generated.")
    
    print("\n" + "="*60)
    print("🎉 COMPLETE ANALYSIS WORKFLOW FINISHED!")
    print("="*60)
    print("📁 Generated files:")
    print("   • Field Matrix: excel_analysis_results/improved_field_matrix.xlsx")
    print("   • Detailed Analysis: excel_analysis_results/improved_detailed_analysis.xlsx")
    print("   • Analysis Report: excel_analysis_results/improved_analysis_report.json")
    print("   • Comprehensive Excel Report: excel_analysis_results/comprehensive_report/comprehensive_analysis_report.xlsx")
    print("   • HTML Report: excel_analysis_results/comprehensive_report/comprehensive_analysis_report.html")
    print("   • Charts: excel_analysis_results/comprehensive_report/*.png")
    print("\n💡 The HTML report should have opened in your browser.")
    print("   If not, manually open: excel_analysis_results/comprehensive_report/comprehensive_analysis_report.html")
    print("="*60)

if __name__ == "__main__":
    main() 