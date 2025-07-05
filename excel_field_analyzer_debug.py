import pandas as pd
import numpy as np
from pathlib import Path
import argparse
import sys
from typing import Dict, List, Set, Tuple
import json
from datetime import datetime

class ExcelFieldAnalyzerDebug:
    """
    A debug version of the Excel field analyzer that shows all columns including unnamed ones.
    """
    
    def __init__(self, excel_file_path: str):
        """
        Initialize the analyzer with an Excel file path.
        
        Args:
            excel_file_path (str): Path to the Excel file to analyze
        """
        self.excel_file_path = Path(excel_file_path)
        self.sheet_data = {}
        self.all_fields = set()
        self.field_matrix = {}
        
    def load_excel_file(self) -> bool:
        """
        Load the Excel file and extract all worksheets.
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            if not self.excel_file_path.exists():
                print(f"Error: File '{self.excel_file_path}' not found.")
                return False
                
            # Read all sheets from the Excel file
            excel_file = pd.ExcelFile(self.excel_file_path)
            print(f"Found {len(excel_file.sheet_names)} worksheets: {excel_file.sheet_names}")
            
            for sheet_name in excel_file.sheet_names:
                try:
                    # Read the sheet
                    df = pd.read_excel(self.excel_file_path, sheet_name=sheet_name)
                    self.sheet_data[sheet_name] = df
                    print(f"Loaded sheet '{sheet_name}' with {len(df)} rows and {len(df.columns)} columns")
                except Exception as e:
                    print(f"Warning: Could not load sheet '{sheet_name}': {e}")
                    
            return True
            
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            return False
    
    def analyze_all_columns(self):
        """
        Analyze all columns in all sheets, including unnamed ones.
        """
        print("\n" + "="*80)
        print("DETAILED COLUMN ANALYSIS (INCLUDING UNNAMED COLUMNS)")
        print("="*80)
        
        for sheet_name, df in self.sheet_data.items():
            print(f"\n--- SHEET: {sheet_name} ---")
            print(f"Total columns: {len(df.columns)}")
            print(f"Total rows: {len(df)}")
            
            print("\nColumn details:")
            for i, col in enumerate(df.columns):
                col_str = str(col)
                # Show first few non-null values
                non_null_values = df[col].dropna().head(3).tolist()
                sample_values = [str(v)[:50] for v in non_null_values]  # Truncate long values
                
                print(f"  {i+1:2d}. '{col_str}'")
                print(f"      Type: {df[col].dtype}")
                print(f"      Non-null count: {df[col].count()}")
                if sample_values:
                    print(f"      Sample values: {sample_values}")
                else:
                    print(f"      Sample values: [all null]")
                print()
    
    def extract_all_fields(self) -> Dict[str, Set[str]]:
        """
        Extract ALL field names from each worksheet, including unnamed columns.
        
        Returns:
            Dict[str, Set[str]]: Dictionary mapping sheet names to sets of field names
        """
        sheet_fields = {}
        
        for sheet_name, df in self.sheet_data.items():
            # Get ALL column names without filtering
            fields = set()
            for col in df.columns:
                col_str = str(col)
                fields.add(col_str)
            
            sheet_fields[sheet_name] = fields
            self.all_fields.update(fields)
            
        return sheet_fields
    
    def create_field_matrix(self) -> pd.DataFrame:
        """
        Create a matrix showing which fields are present in which worksheets.
        
        Returns:
            pd.DataFrame: Matrix with sheets as rows and fields as columns
        """
        # Extract fields from all sheets
        sheet_fields = self.extract_all_fields()
        
        # Create the matrix
        matrix_data = {}
        for sheet_name, fields in sheet_fields.items():
            matrix_data[sheet_name] = {}
            for field in self.all_fields:
                matrix_data[sheet_name][field] = 1 if field in fields else 0
        
        # Convert to DataFrame
        self.field_matrix = pd.DataFrame(matrix_data).T
        
        return self.field_matrix
    
    def generate_summary_report(self) -> Dict:
        """
        Generate a comprehensive summary report of the field analysis.
        
        Returns:
            Dict: Summary statistics and insights
        """
        if self.field_matrix.empty:
            self.create_field_matrix()
        
        total_sheets = len(self.sheet_data)
        total_fields = len(self.all_fields)
        
        # Count fields per sheet
        fields_per_sheet = self.field_matrix.sum(axis=1).to_dict()
        
        # Count sheets per field
        sheets_per_field = self.field_matrix.sum(axis=0).to_dict()
        
        # Find common fields (present in multiple sheets)
        common_fields = {field: count for field, count in sheets_per_field.items() if count > 1}
        
        # Find unique fields (present in only one sheet)
        unique_fields = {field: count for field, count in sheets_per_field.items() if count == 1}
        
        # Find universal fields (present in all sheets)
        universal_fields = {field: count for field, count in sheets_per_field.items() if count == total_sheets}
        
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
            'all_field_names': sorted(list(self.all_fields))
        }
        
        return report
    
    def print_summary(self):
        """
        Print a summary of the analysis to the console.
        """
        if self.field_matrix.empty:
            self.create_field_matrix()
        
        report = self.generate_summary_report()
        
        print("\n" + "="*60)
        print("EXCEL FIELD ANALYSIS SUMMARY (ALL COLUMNS)")
        print("="*60)
        print(f"File: {report['file_path']}")
        print(f"Analysis Date: {report['analysis_date']}")
        print(f"Total Sheets: {report['total_sheets']}")
        print(f"Total Unique Fields: {report['total_unique_fields']}")
        print(f"Common Fields (in multiple sheets): {len(report['common_fields'])}")
        print(f"Unique Fields (in single sheet): {len(report['unique_fields'])}")
        print(f"Universal Fields (in all sheets): {len(report['universal_fields'])}")
        
        print("\n" + "-"*40)
        print("SHEET NAMES:")
        print("-"*40)
        for i, sheet_name in enumerate(report['sheet_names'], 1):
            field_count = report['fields_per_sheet'][sheet_name]
            print(f"{i:2d}. {sheet_name} ({field_count} fields)")
        
        if report['universal_fields']:
            print("\n" + "-"*40)
            print("UNIVERSAL FIELDS (present in all sheets):")
            print("-"*40)
            for field in sorted(report['universal_fields'].keys()):
                print(f"  • {field}")
        
        if report['common_fields']:
            print("\n" + "-"*40)
            print("COMMON FIELDS (present in multiple sheets):")
            print("-"*40)
            sorted_common = sorted(report['common_fields'].items(), key=lambda x: x[1], reverse=True)
            for field, count in sorted_common:
                print(f"  • {field} ({count} sheets)")
        
        print("\n" + "-"*40)
        print("ALL UNIQUE FIELDS:")
        print("-"*40)
        for i, field in enumerate(report['all_field_names'], 1):
            sheet_count = report['sheets_per_field'][field]
            print(f"{i:3d}. {field} ({sheet_count} sheets)")
        
        print("\n" + "="*60)
    
    def save_results(self, output_dir: str = "excel_analysis_results_debug"):
        """
        Save the analysis results to files.
        
        Args:
            output_dir (str): Directory to save results
        """
        output_path = Path(output_dir)
        output_path.mkdir(exist_ok=True)
        
        # Save field matrix as Excel
        matrix_file = output_path / "field_matrix_all_columns.xlsx"
        self.field_matrix.to_excel(matrix_file)
        print(f"Field matrix saved to: {matrix_file}")
        
        # Save summary report as JSON
        report = self.generate_summary_report()
        report_file = output_path / "analysis_report_all_columns.json"
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        print(f"Analysis report saved to: {report_file}")
        
        # Save detailed field information as Excel
        detailed_file = output_path / "detailed_field_analysis_all_columns.xlsx"
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
        
        print(f"Detailed analysis saved to: {detailed_file}")


def main():
    """
    Main function to run the Excel field analyzer debug version.
    """
    parser = argparse.ArgumentParser(description='Analyze Excel file fields across multiple worksheets (debug version)')
    parser.add_argument('excel_file', help='Path to the Excel file to analyze')
    parser.add_argument('--output-dir', default='excel_analysis_results_debug', 
                       help='Directory to save results (default: excel_analysis_results_debug)')
    parser.add_argument('--no-save', action='store_true', 
                       help='Skip saving results to files')
    
    args = parser.parse_args()
    
    # Create analyzer instance
    analyzer = ExcelFieldAnalyzerDebug(args.excel_file)
    
    # Load the Excel file
    if not analyzer.load_excel_file():
        sys.exit(1)
    
    # Analyze all columns in detail
    analyzer.analyze_all_columns()
    
    # Create field matrix
    matrix = analyzer.create_field_matrix()
    
    # Print summary
    analyzer.print_summary()
    
    # Save results if requested
    if not args.no_save:
        analyzer.save_results(args.output_dir)
    
    print(f"\nAnalysis complete! Field matrix shape: {matrix.shape}")


if __name__ == "__main__":
    main() 