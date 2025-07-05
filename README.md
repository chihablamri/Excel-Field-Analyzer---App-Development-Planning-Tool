# Excel Field Analyzer - App Development Planning Tool

A comprehensive Python suite for analyzing Excel files with multiple worksheets to identify field patterns and guide database design for app development. Features smart field detection, categorization, and statistical analysis to help developers understand data structures across complex Excel workbooks.

## 🎯 Purpose

This tool analyzes Excel files with multiple worksheets to identify unique fields across all sheets. This helps determine the number of fields required when building an app and provides insights into data structure patterns. Perfect for developers who need to understand complex Excel data structures before designing database schemas or building applications.

## ✨ Key Features

- **Smart Field Detection**: Identifies actual field names from data rows using pattern recognition
- **Multi-Sheet Analysis**: Processes all worksheets in an Excel file simultaneously
- **Field Categorization**: Automatically categorizes fields by business function
- **Statistical Analysis**: Provides usage frequency and field distribution statistics
- **Multiple Interfaces**: CLI, GUI, and drag-and-drop options for different user preferences
- **Comprehensive Reporting**: Generates detailed Excel reports and JSON data for further processing

## 📁 Output Files

The application generates three key output files:

1. **`improved_field_matrix.xlsx`** - Complete matrix showing which fields are present in which worksheets
2. **`improved_detailed_analysis.xlsx`** - Detailed analysis with field categories and statistics  
3. **`improved_analysis_report.json`** - Raw analysis data for further processing

## 🚀 Quick Start

### Option 1: Generate Sample Data (Recommended for Testing)
```bash
python create_sample_data.py
python excel_analyzer_cli.py "sample_data.xlsx"
```

### Option 2: Command Line
```bash
python excel_analyzer_cli.py "your_excel_file.xlsx"
```

### Option 3: Drag & Drop (Windows)
Simply drag your Excel file onto `analyze_excel.bat`

### Option 4: GUI Application
```bash
python excel_analyzer_app.py
```

## 📋 Available Tools

### 1. `excel_analyzer_cli.py` - Command Line Interface
**Features:**
- Simple command-line interface
- Automatic output directory creation
- Progress indicators and detailed logging
- Summary report in console

**Usage:**
```bash
# Basic usage
python excel_analyzer_cli.py "my_file.xlsx"

# Custom output directory
python excel_analyzer_cli.py "my_file.xlsx" --output "my_results"

# Skip console summary
python excel_analyzer_cli.py "my_file.xlsx" --no-summary
```

### 2. `excel_analyzer_app.py` - GUI Application
**Features:**
- User-friendly graphical interface
- File browser for input selection
- Real-time progress tracking
- Results display in application window

**Usage:**
```bash
python excel_analyzer_app.py
```

### 3. `analyze_excel.bat` - Windows Batch File
**Features:**
- Drag and drop functionality
- No command line knowledge required
- Automatic execution

**Usage:**
- Drag any Excel file onto `analyze_excel.bat`
- Or run: `analyze_excel.bat "path\to\file.xlsx"`

### 4. `excel_field_analyzer_improved.py` - Advanced Analysis
**Features:**
- Advanced header detection
- Field categorization
- Comprehensive reporting

### 5. `excel_field_analyzer_debug.py` - Debug Version
**Features:**
- Shows all columns including unnamed ones
- Detailed column analysis
- Useful for troubleshooting

### 6. `summary_report.py` - Summary Generator
**Features:**
- Generates readable summary reports
- Field categorization
- App development recommendations

## 📊 Analysis Features

### Field Detection
- **Smart Header Detection**: Identifies actual field names from data rows
- **Pattern Recognition**: Recognizes common field naming patterns
- **Keyword Matching**: Identifies fields based on common business terms

### Field Categorization
- **Order Information**: Purchase Order, Order Details
- **Production Details**: Build Time, Cut Time, Man Mins
- **Timing**: Due Date, Due In, Production Date
- **Product Information**: Product, Description
- **Build Information**: Built By, Build Information
- **Despatch Information**: Shipping Code, APC, DX, Van
- **Capacity & Planning**: Capacity, Planning fields

### Statistical Analysis
- **Usage Frequency**: How many sheets use each field
- **Common Fields**: Fields used across multiple sheets
- **Unique Fields**: Fields used in only one sheet
- **Universal Fields**: Fields present in all sheets

## 📈 Example Output

```
📊 ANALYSIS SUMMARY
============================================================
📁 File: Production Schedule.xlsx
📅 Analysis Date: 2025-06-27T06:06:36.211513
📋 Total Sheets: 11
🔤 Total Unique Fields: 64
🔄 Common Fields (multiple sheets): 32
📌 Unique Fields (single sheet): 32
🌐 Universal Fields (all sheets): 0

🏆 MOST COMMON FIELDS (7+ sheets):
  • Shipping Code (9 sheets)
  • Production Date (9 sheets)
  • Product (9 sheets)
  • Due Date (9 sheets)
  • Purchase Order (9 sheets)
  • Description (9 sheets)
```

## 💡 App Development Recommendations

### Core Fields (Include in All Modules)
- Purchase Order, Order Details, Due Date
- Product, Description, Shipping Code
- Build Time, Cut Time, Man Mins, Total Man Mins
- Built By, Build Information

### Database Design Suggestions
- Use flexible schema for varying field sets
- Implement field mapping for different worksheet types
- Consider dynamic form generation
- Include field validation based on usage patterns

## 🔧 Installation

1. **Clone the repository:**
```bash
git clone https://github.com/yourusername/excel-field-analyzer.git
cd excel-field-analyzer
```

2. **Install Python Dependencies:**
```bash
pip install -r requirements.txt
```

3. **Required Packages:**
- pandas >= 1.5.0
- openpyxl >= 3.0.0
- numpy >= 1.21.0

## 📝 Usage Examples

### Test with Sample Data
```bash
# Generate sample Excel file with multiple worksheets
python create_sample_data.py

# Analyze the sample data
python excel_analyzer_cli.py "sample_data.xlsx"
```

### Basic Analysis
```bash
python excel_analyzer_cli.py "your_excel_file.xlsx"
```

### Custom Output Directory
```bash
python excel_analyzer_cli.py "my_file.xlsx" --output "my_analysis_results"
```

### Batch Processing
```bash
# Analyze multiple files
for file in *.xlsx; do
    python excel_analyzer_cli.py "$file" --output "results_${file%.*}"
done
```

## 🎯 Use Cases

- **App Development**: Determine required fields for database design
- **Data Migration**: Understand field mapping between systems
- **Process Analysis**: Identify common patterns across worksheets
- **Documentation**: Create field inventories for existing systems
- **Quality Assurance**: Verify field consistency across sheets

## 📞 Support

The application handles various Excel file structures:
- Multiple worksheets
- Unnamed columns
- Embedded headers in data rows
- Mixed data types
- Large datasets

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🔄 Version History

- **v1.0.0** - Initial release with CLI and GUI interfaces
- **v1.1.0** - Added advanced field categorization and statistical analysis
- **v1.2.0** - Improved header detection and pattern recognition
- **v1.3.0** - Added comprehensive reporting and app development recommendations 