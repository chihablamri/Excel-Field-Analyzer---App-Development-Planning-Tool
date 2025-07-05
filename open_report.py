#!/usr/bin/env python3
"""
Simple script to open the generated HTML report in the default browser.
"""

import webbrowser
from pathlib import Path
import os

def open_html_report():
    """Open the HTML report in the default browser."""
    report_path = Path("excel_analysis_results/comprehensive_report/comprehensive_analysis_report.html")
    
    if report_path.exists():
        # Convert to absolute path and file URL
        absolute_path = report_path.absolute()
        file_url = f"file:///{absolute_path.as_posix()}"
        
        print(f"🌐 Opening HTML report: {report_path}")
        print(f"📁 File location: {absolute_path}")
        
        try:
            webbrowser.open(file_url)
            print("✅ Report opened in your default browser!")
        except Exception as e:
            print(f"❌ Error opening report: {e}")
            print(f"💡 You can manually open: {absolute_path}")
    else:
        print("❌ HTML report not found!")
        print("💡 Please run 'python generate_comprehensive_report.py' first.")

if __name__ == "__main__":
    open_html_report() 