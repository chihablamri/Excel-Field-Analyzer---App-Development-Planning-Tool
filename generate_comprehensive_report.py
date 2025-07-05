#!/usr/bin/env python3
"""
Comprehensive Report Generator for Excel Field Analysis
Generates a detailed report with all collected information, charts, and recommendations.
"""

import json
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
from datetime import datetime
import argparse
import sys
from typing import Dict, List
import numpy as np

class ComprehensiveReportGenerator:
    """Generates comprehensive reports from Excel field analysis."""
    
    def __init__(self, analysis_dir: str = "excel_analysis_results"):
        self.analysis_dir = Path(analysis_dir)
        self.report_data = None
        self.output_dir = None
        
    def load_analysis_data(self) -> bool:
        """Load the analysis data from JSON file."""
        report_file = self.analysis_dir / "improved_analysis_report.json"
        
        if not report_file.exists():
            print(f"ERROR: Analysis report not found: {report_file}")
            print("Please run the Excel analyzer first.")
            return False
            
        try:
            with open(report_file, 'r', encoding='utf-8') as f:
                self.report_data = json.load(f)
            print(f"Loaded analysis data from: {report_file}")
            return True
        except Exception as e:
            print(f"Error loading analysis data: {e}")
            return False
    
    def create_charts(self, output_dir: Path):
        """Create visualizations for the report."""
        print("Creating charts and visualizations...")
        
        # Set style
        plt.style.use('default')
        sns.set_palette("husl")
        
        # 1. Field Usage Distribution
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
        
        # Field usage histogram
        usage_counts = list(self.report_data['sheets_per_field'].values())
        ax1.hist(usage_counts, bins=range(1, max(usage_counts) + 2), alpha=0.7, edgecolor='black')
        ax1.set_xlabel('Number of Sheets Using Field')
        ax1.set_ylabel('Number of Fields')
        ax1.set_title('Field Usage Distribution')
        ax1.grid(True, alpha=0.3)
        
        # Top fields by usage
        sorted_fields = sorted(self.report_data['sheets_per_field'].items(), 
                             key=lambda x: x[1], reverse=True)[:15]
        field_names = [f[:20] + '...' if len(f) > 20 else f for f, _ in sorted_fields]
        field_counts = [c for _, c in sorted_fields]
        
        bars = ax2.barh(range(len(field_names)), field_counts)
        ax2.set_yticks(range(len(field_names)))
        ax2.set_yticklabels(field_names)
        ax2.set_xlabel('Number of Sheets')
        ax2.set_title('Top 15 Most Used Fields')
        ax2.grid(True, alpha=0.3)
        
        # Add value labels on bars
        for i, bar in enumerate(bars):
            width = bar.get_width()
            ax2.text(width + 0.1, bar.get_y() + bar.get_height()/2, 
                    str(int(width)), ha='left', va='center')
        
        plt.tight_layout()
        plt.savefig(output_dir / 'field_usage_analysis.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        # 2. Sheet Analysis
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
        
        # Field count per sheet
        sheet_field_counts = list(self.report_data['fields_per_sheet'].values())
        sheet_names = list(self.report_data['fields_per_sheet'].keys())
        
        bars = ax1.bar(range(len(sheet_names)), sheet_field_counts)
        ax1.set_xlabel('Worksheets')
        ax1.set_ylabel('Number of Fields')
        ax1.set_title('Fields per Worksheet')
        ax1.set_xticks(range(len(sheet_names)))
        ax1.set_xticklabels([name[:15] + '...' if len(name) > 15 else name 
                            for name in sheet_names], rotation=45, ha='right')
        ax1.grid(True, alpha=0.3)
        
        # Add value labels on bars
        for i, bar in enumerate(bars):
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                    str(int(height)), ha='center', va='bottom')
        
        # Field categories pie chart
        category_counts = {}
        for category, fields in self.report_data['field_categories'].items():
            if fields:
                category_counts[category] = len(fields)
        
        if category_counts:
            ax2.pie(category_counts.values(), labels=category_counts.keys(), autopct='%1.1f%%')
            ax2.set_title('Field Categories Distribution')
        
        plt.tight_layout()
        plt.savefig(output_dir / 'sheet_and_category_analysis.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        print("   Charts saved to output directory")
    
    def generate_excel_report(self, output_dir: Path):
        """Generate a comprehensive Excel report."""
        print("Generating comprehensive Excel report...")
        
        with pd.ExcelWriter(output_dir / 'comprehensive_analysis_report.xlsx', engine='openpyxl') as writer:
            
            # 1. Executive Summary
            summary_data = {
                'Metric': [
                    'File Analyzed',
                    'Analysis Date',
                    'Total Worksheets',
                    'Total Unique Fields',
                    'Common Fields (Multiple Sheets)',
                    'Unique Fields (Single Sheet)',
                    'Universal Fields (All Sheets)',
                    'Most Used Field',
                    'Average Fields per Sheet',
                    'Field Usage Standard Deviation'
                ],
                'Value': [
                    Path(self.report_data['file_path']).name,
                    self.report_data['analysis_date'],
                    self.report_data['total_sheets'],
                    self.report_data['total_unique_fields'],
                    len(self.report_data['common_fields']),
                    len(self.report_data['unique_fields']),
                    len(self.report_data['universal_fields']),
                    max(self.report_data['sheets_per_field'].items(), key=lambda x: x[1])[0],
                    f"{np.mean(list(self.report_data['fields_per_sheet'].values())):.1f}",
                    f"{np.std(list(self.report_data['fields_per_sheet'].values())):.1f}"
                ]
            }
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='Executive_Summary', index=False)
            
            # 2. Field Usage Matrix
            # Load the field matrix
            matrix_file = self.analysis_dir / "improved_field_matrix.xlsx"
            if matrix_file.exists():
                field_matrix = pd.read_excel(matrix_file, index_col=0)
                field_matrix.to_excel(writer, sheet_name='Field_Usage_Matrix')
            
            # 3. Field Details
            field_details = []
            for field, sheet_count in self.report_data['sheets_per_field'].items():
                percentage = (sheet_count / self.report_data['total_sheets']) * 100
                field_details.append({
                    'Field_Name': field,
                    'Sheets_Present': sheet_count,
                    'Percentage_of_Sheets': f"{percentage:.1f}%",
                    'Usage_Level': 'High' if percentage >= 70 else 'Medium' if percentage >= 40 else 'Low'
                })
            
            pd.DataFrame(field_details).sort_values('Sheets_Present', ascending=False).to_excel(
                writer, sheet_name='Field_Details', index=False
            )
            
            # 4. Sheet Analysis
            sheet_analysis = []
            for sheet_name, field_count in self.report_data['fields_per_sheet'].items():
                sheet_analysis.append({
                    'Sheet_Name': sheet_name,
                    'Field_Count': field_count,
                    'Percentage_of_Total_Fields': f"{(field_count / self.report_data['total_unique_fields']) * 100:.1f}%"
                })
            
            pd.DataFrame(sheet_analysis).sort_values('Field_Count', ascending=False).to_excel(
                writer, sheet_name='Sheet_Analysis', index=False
            )
            
            # 5. Field Categories
            for category, fields in self.report_data['field_categories'].items():
                if fields:
                    category_data = []
                    for field in fields:
                        sheet_count = self.report_data['sheets_per_field'].get(field, 0)
                        percentage = (sheet_count / self.report_data['total_sheets']) * 100
                        category_data.append({
                            'Field_Name': field,
                            'Sheets_Present': sheet_count,
                            'Percentage_of_Sheets': f"{percentage:.1f}%",
                            'Importance_Level': 'Critical' if percentage >= 70 else 'Important' if percentage >= 40 else 'Optional'
                        })
                    
                    sheet_name = category.replace(' ', '_')[:31]
                    pd.DataFrame(category_data).sort_values('Sheets_Present', ascending=False).to_excel(
                        writer, sheet_name=sheet_name, index=False
                    )
            
            # 6. Recommendations
            recommendations = self._generate_recommendations()
            pd.DataFrame(recommendations).to_excel(writer, sheet_name='Recommendations', index=False)
            
            # 7. Data Quality Assessment
            quality_assessment = self._assess_data_quality()
            pd.DataFrame(quality_assessment).to_excel(writer, sheet_name='Data_Quality', index=False)
        
        print("   Excel report saved")
    
    def _generate_recommendations(self) -> List[Dict]:
        """Generate app development recommendations."""
        recommendations = []
        
        # Core fields recommendations
        core_fields = [f for f, count in self.report_data['sheets_per_field'].items() if count >= 7]
        recommendations.append({
            'Category': 'Core Fields',
            'Recommendation': 'Include these fields in all modules',
            'Fields': ', '.join(core_fields[:10]),
            'Priority': 'High',
            'Rationale': f'These {len(core_fields)} fields are used across 7+ sheets'
        })
        
        # Database design recommendations
        recommendations.append({
            'Category': 'Database Design',
            'Recommendation': 'Use flexible schema with field mapping',
            'Fields': 'N/A',
            'Priority': 'High',
            'Rationale': 'Field sets vary significantly between sheets'
        })
        
        # Module-specific recommendations
        for category, fields in self.report_data['field_categories'].items():
            if fields and len(fields) > 1:
                high_usage_fields = [f for f in fields 
                                   if self.report_data['sheets_per_field'].get(f, 0) >= 5]
                if high_usage_fields:
                    recommendations.append({
                        'Category': f'{category} Module',
                        'Recommendation': f'Create dedicated {category.lower()} module',
                        'Fields': ', '.join(high_usage_fields[:5]),
                        'Priority': 'Medium',
                        'Rationale': f'{len(high_usage_fields)} fields with high usage in this category'
                    })
        
        # UI/UX recommendations
        recommendations.append({
            'Category': 'User Interface',
            'Recommendation': 'Implement dynamic form generation',
            'Fields': 'N/A',
            'Priority': 'Medium',
            'Rationale': 'Field requirements vary by worksheet type'
        })
        
        return recommendations
    
    def _assess_data_quality(self) -> List[Dict]:
        """Assess data quality and consistency."""
        quality_metrics = []
        
        # Field consistency
        total_fields = self.report_data['total_unique_fields']
        common_fields = len(self.report_data['common_fields'])
        consistency_score = (common_fields / total_fields) * 100 if total_fields > 0 else 0
        
        quality_metrics.append({
            'Metric': 'Field Consistency',
            'Score': f"{consistency_score:.1f}%",
            'Description': f'{common_fields} out of {total_fields} fields used across multiple sheets',
            'Quality_Level': 'Good' if consistency_score >= 50 else 'Needs Improvement'
        })
        
        # Field standardization
        standardized_fields = sum(1 for field in self.report_data['all_field_names'] 
                                if not field.startswith('Column_') and not field.startswith('Unnamed:'))
        standardization_score = (standardized_fields / total_fields) * 100 if total_fields > 0 else 0
        
        quality_metrics.append({
            'Metric': 'Field Standardization',
            'Score': f"{standardization_score:.1f}%",
            'Description': f'{standardized_fields} out of {total_fields} fields have meaningful names',
            'Quality_Level': 'Good' if standardization_score >= 70 else 'Needs Improvement'
        })
        
        # Sheet coverage
        avg_fields_per_sheet = np.mean(list(self.report_data['fields_per_sheet'].values()))
        coverage_score = min(100, (avg_fields_per_sheet / 20) * 100)  # Assuming 20 fields is optimal
        
        quality_metrics.append({
            'Metric': 'Sheet Coverage',
            'Score': f"{coverage_score:.1f}%",
            'Description': f'Average {avg_fields_per_sheet:.1f} fields per sheet',
            'Quality_Level': 'Good' if coverage_score >= 60 else 'Needs Improvement'
        })
        
        return quality_metrics
    
    def generate_html_report(self, output_dir: Path):
        """Generate an HTML report with embedded charts."""
        print("Generating HTML report...")
        
        html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Field Analysis Report</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 40px; background-color: #f5f5f5; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 0 20px rgba(0,0,0,0.1); }}
        h1 {{ color: #2c3e50; text-align: center; border-bottom: 3px solid #3498db; padding-bottom: 10px; }}
        h2 {{ color: #34495e; margin-top: 30px; }}
        .summary-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin: 20px 0; }}
        .summary-card {{ background: #ecf0f1; padding: 20px; border-radius: 8px; text-align: center; }}
        .summary-card h3 {{ margin: 0; color: #2c3e50; }}
        .summary-card .value {{ font-size: 2em; font-weight: bold; color: #3498db; }}
        .chart-container {{ text-align: center; margin: 30px 0; }}
        .chart-container img {{ max-width: 100%; height: auto; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }}
        table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
        th, td {{ padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }}
        th {{ background-color: #3498db; color: white; }}
        tr:nth-child(even) {{ background-color: #f2f2f2; }}
        .recommendation {{ background: #e8f5e8; padding: 15px; margin: 10px 0; border-left: 4px solid #27ae60; border-radius: 4px; }}
        .warning {{ background: #fff3cd; padding: 15px; margin: 10px 0; border-left: 4px solid #ffc107; border-radius: 4px; }}
        .info {{ background: #d1ecf1; padding: 15px; margin: 10px 0; border-left: 4px solid #17a2b8; border-radius: 4px; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Excel Field Analysis Report</h1>
        
        <div class="info">
            <strong>File Analyzed:</strong> {Path(self.report_data['file_path']).name}<br>
            <strong>Analysis Date:</strong> {self.report_data['analysis_date']}<br>
            <strong>Generated by:</strong> Excel Field Analyzer v2.0
        </div>
        
        <h2>📈 Executive Summary</h2>
        <div class="summary-grid">
            <div class="summary-card">
                <h3>Worksheets</h3>
                <div class="value">{self.report_data['total_sheets']}</div>
            </div>
            <div class="summary-card">
                <h3>Unique Fields</h3>
                <div class="value">{self.report_data['total_unique_fields']}</div>
            </div>
            <div class="summary-card">
                <h3>Common Fields</h3>
                <div class="value">{len(self.report_data['common_fields'])}</div>
            </div>
            <div class="summary-card">
                <h3>Universal Fields</h3>
                <div class="value">{len(self.report_data['universal_fields'])}</div>
            </div>
        </div>
        
        <h2>📊 Field Usage Analysis</h2>
        <div class="chart-container">
            <img src="field_usage_analysis.png" alt="Field Usage Analysis">
        </div>
        
        <h2>📋 Sheet Analysis</h2>
        <div class="chart-container">
            <img src="sheet_and_category_analysis.png" alt="Sheet and Category Analysis">
        </div>
        
        <h2>🏆 Top 10 Most Used Fields</h2>
        <table>
            <tr><th>Rank</th><th>Field Name</th><th>Sheets Used</th><th>Usage %</th></tr>
        """
        
        sorted_fields = sorted(self.report_data['sheets_per_field'].items(), 
                             key=lambda x: x[1], reverse=True)[:10]
        
        for i, (field, count) in enumerate(sorted_fields, 1):
            percentage = (count / self.report_data['total_sheets']) * 100
            html_content += f"""
            <tr>
                <td>{i}</td>
                <td>{field}</td>
                <td>{count}</td>
                <td>{percentage:.1f}%</td>
            </tr>
            """
        
        html_content += """
        </table>
        
        <h2>💡 Key Recommendations</h2>
        """
        
        recommendations = self._generate_recommendations()
        for rec in recommendations[:5]:  # Show top 5 recommendations
            html_content += f"""
            <div class="recommendation">
                <strong>{rec['Category']}:</strong> {rec['Recommendation']}<br>
                <em>Priority: {rec['Priority']}</em><br>
                <small>{rec['Rationale']}</small>
            </div>
            """
        
        html_content += """
        <h2>📄 Field Categories</h2>
        """
        
        for category, fields in self.report_data['field_categories'].items():
            if fields:
                high_usage = [f for f in fields if self.report_data['sheets_per_field'].get(f, 0) >= 5]
                html_content += f"""
                <h3>{category} ({len(fields)} fields)</h3>
                <p><strong>High Usage Fields:</strong> {', '.join(high_usage[:5])}</p>
                """
        
        html_content += """
        <h2>📊 Data Quality Assessment</h2>
        """
        
        quality_metrics = self._assess_data_quality()
        for metric in quality_metrics:
            quality_class = 'recommendation' if 'Good' in metric['Quality_Level'] else 'warning'
            html_content += f"""
            <div class="{quality_class}">
                <strong>{metric['Metric']}:</strong> {metric['Score']} ({metric['Quality_Level']})<br>
                <small>{metric['Description']}</small>
            </div>
            """
        
        html_content += """
        <hr style="margin: 40px 0;">
        <p style="text-align: center; color: #7f8c8d;">
            Report generated by Excel Field Analyzer | 
            For app development and database design guidance
        </p>
    </div>
</body>
</html>
        """
        
        with open(output_dir / 'comprehensive_analysis_report.html', 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print("   HTML report saved")
    
    def generate_report(self, output_dir: str = None):
        """Generate all report formats."""
        if not self.load_analysis_data():
            return False
        
        if output_dir is None:
            output_dir = self.analysis_dir / "comprehensive_report"
        else:
            output_dir = Path(output_dir)
        
        output_dir.mkdir(exist_ok=True)
        self.output_dir = output_dir
        
        print(f"Generating comprehensive report in: {output_dir}")
        
        # Create charts
        self.create_charts(output_dir)
        
        # Generate Excel report
        self.generate_excel_report(output_dir)
        
        # Generate HTML report
        self.generate_html_report(output_dir)
        
        print("\n" + "="*60)
        print("COMPREHENSIVE REPORT GENERATED SUCCESSFULLY!")
        print("="*60)
        print("Generated files:")
        print(f"   - Excel Report: {output_dir / 'comprehensive_analysis_report.xlsx'}")
        print(f"   - HTML Report: {output_dir / 'comprehensive_analysis_report.html'}")
        print(f"   - Charts: {output_dir / 'field_usage_analysis.png'}")
        print(f"   - Charts: {output_dir / 'sheet_and_category_analysis.png'}")
        print("\nOpen the HTML report in your browser for the best viewing experience.")
        print("="*60)
        
        return True

def main():
    """Main function to run the report generator."""
    parser = argparse.ArgumentParser(description='Generate comprehensive reports from Excel field analysis')
    parser.add_argument('--input-dir', default='excel_analysis_results',
                       help='Directory containing analysis results (default: excel_analysis_results)')
    parser.add_argument('--output-dir', default=None,
                       help='Output directory for reports (default: input_dir/comprehensive_report)')
    
    args = parser.parse_args()
    
    generator = ComprehensiveReportGenerator(args.input_dir)
    success = generator.generate_report(args.output_dir)
    
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main() 