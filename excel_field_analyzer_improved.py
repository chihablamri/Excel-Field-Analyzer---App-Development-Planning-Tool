import pandas as pd
import numpy as np
from pathlib import Path
import argparse
import sys
from typing import Dict, List, Set, Tuple
import json
from datetime import datetime
import re

class ImprovedExcelFieldAnalyzer:
    """
    An improved Excel field analyzer that can extract actual field names from data rows
    and create a proper field matrix for production schedules.
    """
    
    def __init__(self, excel_file_path: str):
        """
        Initialize the analyzer with an Excel file path.
        
        Args:
            excel_file_path (str): Path to the Excel file to analyze
        """
        self.excel_file_path = Path(excel_file_path)
        self.sheet_data = {}
        self.sheet_headers = {}
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
    
    def extract_actual_headers(self) -> Dict[str, List[str]]:
        """
        Extract actual field names from the data rows by looking for header-like values.
        
        Returns:
            Dict[str, List[str]]: Dictionary mapping sheet names to lists of field names
        """
        sheet_headers = {}
        
        for sheet_name, df in self.sheet_data.items():
            headers = []
            
            # Look for header-like values in the first few rows
            for col_idx, col in enumerate(df.columns):
                header_found = False
                
                # Check first 5 rows for potential headers
                for row_idx in range(min(5, len(df))):
                    value = str(df.iloc[row_idx, col_idx]).strip()
                    
                    # Check if this looks like a header
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
        """
        Determine if a value is likely a header based on various criteria.
        
        Args:
            value (str): The value to check
            df (pd.DataFrame): The dataframe
            col_idx (int): Column index
            
        Returns:
            bool: True if likely a header
        """
        if not value or value == 'nan' or value == 'None':
            return False
        
        # Common header patterns
        header_patterns = [
            r'^[A-Z][a-z\s]+$',  # Title case words
            r'^[A-Z\s]+$',       # All caps
            r'^[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*$',  # CamelCase or TitleCase
            r'^[A-Z][a-z]+\s+[A-Z][a-z]+$',  # Two title case words
        ]
        
        # Check if value matches header patterns
        for pattern in header_patterns:
            if re.match(pattern, value):
                # Additional check: see if this value appears only once or very few times
                # (headers typically don't repeat much in data)
                value_count = df.iloc[:, col_idx].astype(str).str.contains(value, regex=False, na=False).sum()
                if value_count <= 3:  # Header appears 3 or fewer times
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
                # Check if it's not a common data value
                if not self._is_common_data_value(value, df, col_idx):
                    return True
        
        return False
    
    def _is_common_data_value(self, value: str, df: pd.DataFrame, col_idx: int) -> bool:
        """
        Check if a value is a common data value rather than a header.
        
        Args:
            value (str): The value to check
            df (pd.DataFrame): The dataframe
            col_idx (int): Column index
            
        Returns:
            bool: True if it's a common data value
        """
        # Check if this value appears frequently in the column
        value_count = df.iloc[:, col_idx].astype(str).str.contains(value, regex=False, na=False).sum()
        total_rows = len(df)
        
        # If value appears in more than 10% of rows, it's likely data, not header
        if value_count > total_rows * 0.1:
            return True
        
        # Check for common data patterns
        data_patterns = [
            r'^\d+$',  # Just numbers
            r'^\d{4}-\d{2}-\d{2}',  # Date format
            r'^[A-Z]{2}\d+\s+[A-Z0-9]',  # Shipping codes like "AB1 0BE"
            r'^PO\d+',  # Purchase order numbers
            r'^[A-Z]{2}\d+',  # Product codes
        ]
        
        for pattern in data_patterns:
            if re.match(pattern, value):
                return True
        
        return False
    
    def create_field_matrix(self) -> pd.DataFrame:
        """
        Create a matrix showing which fields are present in which worksheets.
        
        Returns:
            pd.DataFrame: Matrix with sheets as rows and fields as columns
        """
        # Extract actual headers from all sheets
        sheet_headers = self.extract_actual_headers()
        
        # Create the matrix
        matrix_data = {}
        for sheet_name, headers in sheet_headers.items():
            matrix_data[sheet_name] = {}
            for field in self.all_fields:
                matrix_data[sheet_name][field] = 1 if field in headers else 0
        
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
        
        # Group fields by category
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
        """
        Categorize fields into logical groups.
        
        Returns:
            Dict[str, List[str]]: Dictionary of field categories
        """
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
    
    def print_summary(self):
        """
        Print a summary of the analysis to the console.
        """
        if self.field_matrix.empty:
            self.create_field_matrix()
        
        report = self.generate_summary_report()
        
        print("\n" + "="*80)
        print("IMPROVED EXCEL FIELD ANALYSIS SUMMARY")
        print("="*80)
        print(f"File: {report['file_path']}")
        print(f"Analysis Date: {report['analysis_date']}")
        print(f"Total Sheets: {report['total_sheets']}")
        print(f"Total Unique Fields: {report['total_unique_fields']}")
        print(f"Common Fields (in multiple sheets): {len(report['common_fields'])}")
        print(f"Unique Fields (in single sheet): {len(report['unique_fields'])}")
        print(f"Universal Fields (in all sheets): {len(report['universal_fields'])}")
        
        print("\n" + "-"*50)
        print("SHEET NAMES WITH FIELD COUNTS:")
        print("-"*50)
        for i, sheet_name in enumerate(report['sheet_names'], 1):
            field_count = report['fields_per_sheet'][sheet_name]
            print(f"{i:2d}. {sheet_name} ({field_count} fields)")
        
        if report['universal_fields']:
            print("\n" + "-"*50)
            print("UNIVERSAL FIELDS (present in all sheets):")
            print("-"*50)
            for field in sorted(report['universal_fields'].keys()):
                print(f"  • {field}")
        
        if report['common_fields']:
            print("\n" + "-"*50)
            print("COMMON FIELDS (present in multiple sheets):")
            print("-"*50)
            sorted_common = sorted(report['common_fields'].items(), key=lambda x: x[1], reverse=True)
            for field, count in sorted_common[:20]:  # Show top 20
                print(f"  • {field} ({count} sheets)")
            if len(sorted_common) > 20:
                print(f"  ... and {len(sorted_common) - 20} more")
        
        print("\n" + "-"*50)
        print("FIELD CATEGORIES:")
        print("-"*50)
        for category, fields in report['field_categories'].items():
            if fields:
                print(f"\n{category} ({len(fields)} fields):")
                for field in sorted(fields):
                    sheet_count = report['sheets_per_field'].get(field, 0)
                    print(f"  • {field} ({sheet_count} sheets)")
        
        print("\n" + "-"*50)
        print("ALL UNIQUE FIELDS (sorted by usage):")
        print("-"*50)
        sorted_fields = sorted(report['sheets_per_field'].items(), key=lambda x: x[1], reverse=True)
        for i, (field, sheet_count) in enumerate(sorted_fields, 1):
            print(f"{i:3d}. {field} ({sheet_count} sheets)")
        
        print("\n" + "="*80)
    
    def save_results(self, output_dir: str = "excel_analysis_results_improved"):
        """
        Save the analysis results to files.
        
        Args:
            output_dir (str): Directory to save results
        """
        output_path = Path(output_dir)
        output_path.mkdir(exist_ok=True)
        
        # Save field matrix as Excel
        matrix_file = output_path / "improved_field_matrix.xlsx"
        self.field_matrix.to_excel(matrix_file)
        print(f"Improved field matrix saved to: {matrix_file}")
        
        # Save summary report as JSON
        report = self.generate_summary_report()
        report_file = output_path / "improved_analysis_report.json"
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        print(f"Improved analysis report saved to: {report_file}")
        
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
                    pd.DataFrame(category_data).sort_values('Sheets_Present', ascending=False).to_excel(
                        writer, sheet_name=category.replace(' ', '_')[:31], index=False
                    )
        
        print(f"Improved detailed analysis saved to: {detailed_file}")


def main():
    """
    Main function to run the improved Excel field analyzer.
    """
    parser = argparse.ArgumentParser(description='Analyze Excel file fields with improved header detection')
    parser.add_argument('excel_file', help='Path to the Excel file to analyze')
    parser.add_argument('--output-dir', default='excel_analysis_results_improved', 
                       help='Directory to save results (default: excel_analysis_results_improved)')
    parser.add_argument('--no-save', action='store_true', 
                       help='Skip saving results to files')
    
    args = parser.parse_args()
    
    # Create analyzer instance
    analyzer = ImprovedExcelFieldAnalyzer(args.excel_file)
    
    # Load the Excel file
    if not analyzer.load_excel_file():
        sys.exit(1)
    
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