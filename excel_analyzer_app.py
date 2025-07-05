#!/usr/bin/env python3
"""
Excel Field Analyzer Application
A comprehensive GUI application for analyzing Excel files and generating field matrices.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
from pathlib import Path
import json
from datetime import datetime
import threading
import re
import os
import sys

class ExcelFieldAnalyzer:
    """Core analysis engine for Excel files."""
    
    def __init__(self, excel_file_path: str):
        self.excel_file_path = Path(excel_file_path)
        self.sheet_data = {}
        self.sheet_headers = {}
        self.all_fields = set()
        self.field_matrix = {}
        
    def load_excel_file(self) -> bool:
        """Load the Excel file and extract all worksheets."""
        try:
            if not self.excel_file_path.exists():
                return False, f"File '{self.excel_file_path}' not found."
                
            excel_file = pd.ExcelFile(self.excel_file_path)
            
            for sheet_name in excel_file.sheet_names:
                try:
                    df = pd.read_excel(self.excel_file_path, sheet_name=sheet_name)
                    self.sheet_data[sheet_name] = df
                except Exception as e:
                    return False, f"Could not load sheet '{sheet_name}': {e}"
                    
            return True, f"Successfully loaded {len(excel_file.sheet_names)} worksheets"
            
        except Exception as e:
            return False, f"Error loading Excel file: {e}"
    
    def extract_actual_headers(self) -> Dict[str, List[str]]:
        """Extract actual field names from the data rows."""
        sheet_headers = {}
        
        for sheet_name, df in self.sheet_data.items():
            headers = []
            
            for col_idx, col in enumerate(df.columns):
                header_found = False
                
                # Check first 5 rows for potential headers
                for row_idx in range(min(5, len(df))):
                    value = str(df.iloc[row_idx, col_idx]).strip()
                    
                    if self._is_likely_header(value, df, col_idx):
                        headers.append(value)
                        header_found = True
                        break
                
                # If no header found, use column name or create a generic one
                if not header_found:
                    if str(col).startswith('Unnamed:'):
                        headers.append(f"Column_{col_idx + 1}")
                    else:
                        headers.append(str(col))
            
            sheet_headers[sheet_name] = headers
            self.all_fields.update(headers)
            
        self.sheet_headers = sheet_headers
        return sheet_headers
    
    def _is_likely_header(self, value: str, df: pd.DataFrame, col_idx: int) -> bool:
        """Determine if a value is likely a header."""
        if not value or value == 'nan' or value == 'None':
            return False
        
        # Common header patterns
        header_patterns = [
            r'^[A-Z][a-z\s]+$',
            r'^[A-Z\s]+$',
            r'^[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*$',
            r'^[A-Z][a-z]+\s+[A-Z][a-z]+$',
        ]
        
        for pattern in header_patterns:
            if re.match(pattern, value):
                value_count = df.iloc[:, col_idx].astype(str).str.contains(value, regex=False, na=False).sum()
                if value_count <= 3:
                    return True
        
        # Check for common header keywords
        header_keywords = [
            'order', 'details', 'assigned', 'due', 'production', 'date', 'purchase', 
            'shipping', 'product', 'description', 'cut', 'build', 'time', 'man', 'mins',
            'quantity', 'total', 'information', 'built', 'by', 'despatch', 'pallet',
            'apc', 'dx', 'label', 'printed', 'van', 'notes', 'invoiced', 'capacity'
        ]
        
        value_lower = value.lower()
        for keyword in header_keywords:
            if keyword in value_lower:
                if not self._is_common_data_value(value, df, col_idx):
                    return True
        
        return False
    
    def _is_common_data_value(self, value: str, df: pd.DataFrame, col_idx: int) -> bool:
        """Check if a value is a common data value rather than a header."""
        value_count = df.iloc[:, col_idx].astype(str).str.contains(value, regex=False, na=False).sum()
        total_rows = len(df)
        
        if value_count > total_rows * 0.1:
            return True
        
        data_patterns = [
            r'^\d+$',
            r'^\d{4}-\d{2}-\d{2}',
            r'^[A-Z]{2}\d+\s+[A-Z0-9]',
            r'^PO\d+',
            r'^[A-Z]{2}\d+',
        ]
        
        for pattern in data_patterns:
            if re.match(pattern, value):
                return True
        
        return False
    
    def create_field_matrix(self) -> pd.DataFrame:
        """Create a matrix showing which fields are present in which worksheets."""
        sheet_headers = self.extract_actual_headers()
        
        matrix_data = {}
        for sheet_name, headers in sheet_headers.items():
            matrix_data[sheet_name] = {}
            for field in self.all_fields:
                matrix_data[sheet_name][field] = 1 if field in headers else 0
        
        self.field_matrix = pd.DataFrame(matrix_data).T
        return self.field_matrix
    
    def generate_summary_report(self) -> Dict:
        """Generate a comprehensive summary report."""
        if self.field_matrix.empty:
            self.create_field_matrix()
        
        total_sheets = len(self.sheet_data)
        total_fields = len(self.all_fields)
        
        fields_per_sheet = self.field_matrix.sum(axis=1).to_dict()
        sheets_per_field = self.field_matrix.sum(axis=0).to_dict()
        
        common_fields = {field: count for field, count in sheets_per_field.items() if count > 1}
        unique_fields = {field: count for field, count in sheets_per_field.items() if count == 1}
        universal_fields = {field: count for field, count in sheets_per_field.items() if count == total_sheets}
        
        field_categories = self._categorize_fields()
        
        report = {
            'file_path': str(self.excel_file_path),
            'analysis_date': datetime.now().isoformat(),
            'total_sheets': total_sheets,
            'total_unique_fields': total_fields,
            'fields_per_sheet': fields_per_sheet,
            'sheets_per_field': sheets_per_field,
            'common_fields': common_fields,
            'unique_fields': unique_fields,
            'universal_fields': universal_fields,
            'sheet_names': list(self.sheet_data.keys()),
            'all_field_names': sorted(list(self.all_fields)),
            'field_categories': field_categories
        }
        
        return report
    
    def _categorize_fields(self) -> Dict[str, List[str]]:
        """Categorize fields into logical groups."""
        categories = {
            'Order Information': [],
            'Production Details': [],
            'Timing': [],
            'Product Information': [],
            'Build Information': [],
            'Despatch Information': [],
            'Capacity & Planning': [],
            'Other': []
        }
        
        for field in self.all_fields:
            field_lower = field.lower()
            
            if any(word in field_lower for word in ['order', 'purchase', 'assigned']):
                categories['Order Information'].append(field)
            elif any(word in field_lower for word in ['production', 'build', 'cut', 'man', 'mins']):
                categories['Production Details'].append(field)
            elif any(word in field_lower for word in ['date', 'due', 'time']):
                categories['Timing'].append(field)
            elif any(word in field_lower for word in ['product', 'description']):
                categories['Product Information'].append(field)
            elif any(word in field_lower for word in ['build information', 'built by']):
                categories['Build Information'].append(field)
            elif any(word in field_lower for word in ['despatch', 'shipping', 'pallet', 'apc', 'dx', 'van', 'label']):
                categories['Despatch Information'].append(field)
            elif any(word in field_lower for word in ['capacity', 'planning', 'wc']):
                categories['Capacity & Planning'].append(field)
            else:
                categories['Other'].append(field)
        
        return categories
    
    def save_results(self, output_dir: str) -> Dict[str, str]:
        """Save the analysis results to files."""
        output_path = Path(output_dir)
        output_path.mkdir(exist_ok=True)
        
        saved_files = {}
        
        # Save field matrix as Excel
        matrix_file = output_path / "improved_field_matrix.xlsx"
        self.field_matrix.to_excel(matrix_file)
        saved_files['field_matrix'] = str(matrix_file)
        
        # Save summary report as JSON
        report = self.generate_summary_report()
        report_file = output_path / "improved_analysis_report.json"
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        saved_files['analysis_report'] = str(report_file)
        
        # Save detailed field information as Excel
        detailed_file = output_path / "improved_detailed_analysis.xlsx"
        with pd.ExcelWriter(detailed_file, engine='openpyxl') as writer:
            # Field matrix
            self.field_matrix.to_excel(writer, sheet_name='Field_Matrix')
            
            # Summary statistics
            summary_data = {
                'Metric': ['Total Sheets', 'Total Unique Fields', 'Common Fields', 'Unique Fields', 'Universal Fields'],
                'Count': [
                    report['total_sheets'],
                    report['total_unique_fields'],
                    len(report['common_fields']),
                    len(report['unique_fields']),
                    len(report['universal_fields'])
                ]
            }
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
            
            # Field details
            field_details = []
            for field, sheet_count in report['sheets_per_field'].items():
                field_details.append({
                    'Field_Name': field,
                    'Sheets_Present': sheet_count,
                    'Percentage_of_Sheets': f"{(sheet_count / report['total_sheets']) * 100:.1f}%"
                })
            pd.DataFrame(field_details).sort_values('Sheets_Present', ascending=False).to_excel(
                writer, sheet_name='Field_Details', index=False
            )
            
            # Field categories
            for category, fields in report['field_categories'].items():
                if fields:
                    category_data = []
                    for field in fields:
                        sheet_count = report['sheets_per_field'].get(field, 0)
                        category_data.append({
                            'Field_Name': field,
                            'Sheets_Present': sheet_count,
                            'Percentage_of_Sheets': f"{(sheet_count / report['total_sheets']) * 100:.1f}%"
                        })
                    sheet_name = category.replace(' ', '_')[:31]
                    pd.DataFrame(category_data).sort_values('Sheets_Present', ascending=False).to_excel(
                        writer, sheet_name=sheet_name, index=False
                    )
        
        saved_files['detailed_analysis'] = str(detailed_file)
        
        return saved_files

class ExcelAnalyzerApp:
    """Main GUI application for Excel field analysis."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Field Analyzer")
        self.root.geometry("800x600")
        self.root.configure(bg='#f0f0f0')
        
        self.excel_file_path = None
        self.output_dir = None
        self.analyzer = None
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the user interface."""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Excel Field Analyzer", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection
        ttk.Label(main_frame, text="Excel File:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.file_var = tk.StringVar()
        file_entry = ttk.Entry(main_frame, textvariable=self.file_var, width=50)
        file_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_file).grid(row=1, column=2, pady=5)
        
        # Output directory
        ttk.Label(main_frame, text="Output Directory:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.output_var = tk.StringVar()
        output_entry = ttk.Entry(main_frame, textvariable=self.output_var, width=50)
        output_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_output).grid(row=2, column=2, pady=5)
        
        # Analyze button
        self.analyze_button = ttk.Button(main_frame, text="Analyze Excel File", 
                                       command=self.analyze_file, style='Accent.TButton')
        self.analyze_button.grid(row=3, column=0, columnspan=3, pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # Results text area
        ttk.Label(main_frame, text="Analysis Results:").grid(row=5, column=0, sticky=tk.W, pady=(10, 5))
        self.results_text = scrolledtext.ScrolledText(main_frame, height=20, width=80)
        self.results_text.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Configure row weights
        main_frame.rowconfigure(6, weight=1)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready to analyze Excel files")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_label.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
    def browse_file(self):
        """Browse for Excel file."""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_file_path = file_path
            self.file_var.set(file_path)
            
            # Auto-set output directory
            output_dir = str(Path(file_path).parent / "excel_analysis_results")
            self.output_var.set(output_dir)
            self.output_dir = output_dir
            
    def browse_output(self):
        """Browse for output directory."""
        output_path = filedialog.askdirectory(title="Select Output Directory")
        if output_path:
            self.output_dir = output_path
            self.output_var.set(output_path)
            
    def analyze_file(self):
        """Analyze the selected Excel file."""
        if not self.excel_file_path:
            messagebox.showerror("Error", "Please select an Excel file first.")
            return
            
        if not self.output_dir:
            messagebox.showerror("Error", "Please select an output directory.")
            return
            
        # Start analysis in a separate thread
        self.analyze_button.config(state='disabled')
        self.progress.start()
        self.status_var.set("Analyzing Excel file...")
        
        thread = threading.Thread(target=self._run_analysis)
        thread.daemon = True
        thread.start()
        
    def _run_analysis(self):
        """Run the analysis in a separate thread."""
        try:
            # Create analyzer
            self.analyzer = ExcelFieldAnalyzer(self.excel_file_path)
            
            # Load file
            success, message = self.analyzer.load_excel_file()
            if not success:
                self.root.after(0, lambda: self._show_error(message))
                return
                
            # Create field matrix
            matrix = self.analyzer.create_field_matrix()
            
            # Generate report
            report = self.analyzer.generate_summary_report()
            
            # Save results
            saved_files = self.analyzer.save_results(self.output_dir)
            
            # Update UI with results
            self.root.after(0, lambda: self._show_results(report, saved_files))
            
        except Exception as e:
            self.root.after(0, lambda: self._show_error(f"Analysis failed: {str(e)}"))
            
    def _show_error(self, message):
        """Show error message."""
        self.analyze_button.config(state='normal')
        self.progress.stop()
        self.status_var.set("Analysis failed")
        messagebox.showerror("Error", message)
        
    def _show_results(self, report, saved_files):
        """Show analysis results."""
        self.analyze_button.config(state='normal')
        self.progress.stop()
        self.status_var.set("Analysis completed successfully")
        
        # Clear previous results
        self.results_text.delete(1.0, tk.END)
        
        # Display results
        results = f"""
EXCEL FIELD ANALYSIS RESULTS
{'='*50}

📊 KEY STATISTICS:
• Total Worksheets: {report['total_sheets']}
• Total Unique Fields: {report['total_unique_fields']}
• Fields Used in Multiple Sheets: {len(report['common_fields'])}
• Fields Used in Single Sheet: {len(report['unique_fields'])}
• Universal Fields (All Sheets): {len(report['universal_fields'])}

🏆 MOST COMMON FIELDS (Used in 7+ sheets):
"""
        
        sorted_common = sorted(report['common_fields'].items(), key=lambda x: x[1], reverse=True)
        for field, count in sorted_common:
            if count >= 7:
                results += f"• {field} ({count} sheets)\n"
                
        results += f"""

📄 WORKSHEET ANALYSIS:
"""
        for i, sheet_name in enumerate(report['sheet_names'], 1):
            field_count = report['fields_per_sheet'][sheet_name]
            results += f"{i:2d}. {sheet_name} ({field_count} fields)\n"
            
        results += f"""

📁 GENERATED FILES:
• Field Matrix: {saved_files['field_matrix']}
• Detailed Analysis: {saved_files['detailed_analysis']}
• Analysis Report: {saved_files['analysis_report']}

✅ Analysis completed successfully!
"""
        
        self.results_text.insert(1.0, results)
        
        # Show success message
        messagebox.showinfo("Success", 
                          f"Analysis completed!\n\nGenerated files:\n" +
                          f"• {saved_files['field_matrix']}\n" +
                          f"• {saved_files['detailed_analysis']}\n" +
                          f"• {saved_files['analysis_report']}")

def main():
    """Main function to run the application."""
    root = tk.Tk()
    app = ExcelAnalyzerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 