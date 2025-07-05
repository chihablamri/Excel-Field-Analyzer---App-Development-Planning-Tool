#!/usr/bin/env python3
"""
Excel Field Analyzer - Command Line Interface (Simple Version)
A simple command-line tool for analyzing Excel files and generating field matrices.
"""

import argparse
import sys
from pathlib import Path
import pandas as pd
import json
from datetime import datetime
import re

class ExcelFieldAnalyzer:
    """Core analysis engine for Excel files."""
    
    def __init__(self, excel_file_path: str):
        self.excel_file_path = Path(excel_file_path)
        self.sheet_data = {}
        self.sheet_headers = {}
        self.all_fields = set()
        self.field_matrix = {}
        
    def load_excel_file(self) -> tuple[bool, str]:
        """Load the Excel file and extract all worksheets."""
        try:
            if not self.excel_file_path.exists():
                return False, f"File '{self.excel_file_path}' not found."
                
            excel_file = pd.ExcelFile(self.excel_file_path)
            print(f"Found {len(excel_file.sheet_names)} worksheets")
            
            for sheet_name in excel_file.sheet_names:
                try:
                    df = pd.read_excel(self.excel_file_path, sheet_name=sheet_name)
                    self.sheet_data[sheet_name] = df
                    print(f"   Loaded '{sheet_name}' ({len(df)} rows, {len(df.columns)} columns)")
                except Exception as e:
                    print(f"   Failed to load '{sheet_name}': {e}")
                    
            return True, f"Successfully loaded {len(excel_file.sheet_names)} worksheets"
            
        except Exception as e:
            return False, f"Error loading Excel file: {e}"
    
    def extract_actual_headers(self) -> dict[str, list[str]]:
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
        print("Extracting field names from worksheets...")
        sheet_headers = self.extract_actual_headers()
        
        matrix_data = {}
        for sheet_name, headers in sheet_headers.items():
            matrix_data[sheet_name] = {}
            for field in self.all_fields:
                matrix_data[sheet_name][field] = 1 if field in headers else 0
        
        self.field_matrix = pd.DataFrame(matrix_data).T
        return self.field_matrix
    
    def generate_summary_report(self) -> dict:
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
    
    def _categorize_fields(self) -> dict[str, list[str]]:
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
    
    def save_results(self, output_dir: str) -> dict[str, str]:
        """Save the analysis results to files."""
        output_path = Path(output_dir)
        output_path.mkdir(exist_ok=True)
        
        saved_files = {}
        
        print("Saving analysis results...")
        
        # Save field matrix as Excel
        matrix_file = output_path / "improved_field_matrix.xlsx"
        self.field_matrix.to_excel(matrix_file)
        saved_files['field_matrix'] = str(matrix_file)
        print(f"   Field matrix saved: {matrix_file}")
        
        # Save summary report as JSON
        report = self.generate_summary_report()
        report_file = output_path / "improved_analysis_report.json"
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        saved_files['analysis_report'] = str(report_file)
        print(f"   Analysis report saved: {report_file}")
        
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
        print(f"   Detailed analysis saved: {detailed_file}")
        
        return saved_files
    
    def print_summary(self):
        """Print a summary of the analysis to the console."""
        if self.field_matrix.empty:
            self.create_field_matrix()
        
        report = self.generate_summary_report()
        
        print("\n" + "="*60)
        print("ANALYSIS SUMMARY")
        print("="*60)
        print(f"File: {Path(report['file_path']).name}")
        print(f"Analysis Date: {report['analysis_date']}")
        print(f"Total Sheets: {report['total_sheets']}")
        print(f"Total Unique Fields: {report['total_unique_fields']}")
        print(f"Common Fields (multiple sheets): {len(report['common_fields'])}")
        print(f"Unique Fields (single sheet): {len(report['unique_fields'])}")
        print(f"Universal Fields (all sheets): {len(report['universal_fields'])}")
        
        print("\n" + "-"*40)
        print("MOST COMMON FIELDS (7+ sheets):")
        print("-"*40)
        sorted_common = sorted(report['common_fields'].items(), key=lambda x: x[1], reverse=True)
        for field, count in sorted_common:
            if count >= 7:
                print(f"  * {field} ({count} sheets)")
        
        print("\n" + "-"*40)
        print("WORKSHEET ANALYSIS:")
        print("-"*40)
        for i, sheet_name in enumerate(report['sheet_names'], 1):
            field_count = report['fields_per_sheet'][sheet_name]
            print(f"{i:2d}. {sheet_name} ({field_count} fields)")
        
        print("\n" + "="*60)

def main():
    """Main function to run the command-line analyzer."""
    parser = argparse.ArgumentParser(
        description='Analyze Excel files and generate field matrices',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python excel_analyzer_cli_simple.py "my_file.xlsx"
  python excel_analyzer_cli_simple.py "my_file.xlsx" --output "my_results"
  python excel_analyzer_cli_simple.py "my_file.xlsx" --output "my_results" --no-summary
        """
    )
    
    parser.add_argument('excel_file', help='Path to the Excel file to analyze')
    parser.add_argument('--output', '-o', default=None, 
                       help='Output directory (default: excel_analysis_results)')
    parser.add_argument('--no-summary', action='store_true',
                       help='Skip printing summary to console')
    
    args = parser.parse_args()
    
    # Validate input file
    if not Path(args.excel_file).exists():
        print(f"ERROR: File '{args.excel_file}' not found.")
        sys.exit(1)
    
    # Set output directory
    if args.output:
        output_dir = args.output
    else:
        output_dir = str(Path(args.excel_file).parent / "excel_analysis_results")
    
    print("Starting Excel Field Analysis...")
    print(f"Input file: {args.excel_file}")
    print(f"Output directory: {output_dir}")
    
    # Create analyzer
    analyzer = ExcelFieldAnalyzer(args.excel_file)
    
    # Load file
    success, message = analyzer.load_excel_file()
    if not success:
        print(f"ERROR: {message}")
        sys.exit(1)
    
    # Create field matrix
    matrix = analyzer.create_field_matrix()
    
    # Save results
    saved_files = analyzer.save_results(output_dir)
    
    # Print summary
    if not args.no_summary:
        analyzer.print_summary()
    
    print("\n" + "="*60)
    print("ANALYSIS COMPLETED SUCCESSFULLY!")
    print("="*60)
    print("Generated files:")
    print(f"   * Field Matrix: {saved_files['field_matrix']}")
    print(f"   * Detailed Analysis: {saved_files['detailed_analysis']}")
    print(f"   * Analysis Report: {saved_files['analysis_report']}")
    print("\nUse these files to guide your app development and database design.")
    print("="*60)

if __name__ == "__main__":
    main() 