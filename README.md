# Unified Excel Processing System

A comprehensive Python-based system that processes Excel files through an AI-powered analysis pipeline, converting them to HTML format, analyzing content with OpenAI's GPT models, and generating updated Excel files with filled data.

## 🎯 Project Overview

This system automates the process of filling empty cells in Excel spreadsheets by:
1. **Converting Excel files to HTML** with preserved formatting and structure
2. **Analyzing HTML tables** using OpenAI's GPT models to match text data to appropriate cells
3. **Generating multiple output formats** (CSV, JSON, Excel) with cell mappings
4. **Updating original Excel files** with the analyzed data

The system is designed for processing complex Excel workbooks with multiple worksheets, particularly useful for financial planning, project documentation, and data analysis workflows.

## 🏗️ System Architecture

```
Input_Folder/          →  Excel files (.xlsx)
    ↓ (Excel to HTML conversion)
html_outputs/          →  HTML files (.html)
    ↓ (AI Analysis with text data)
DATA_SOURCES/          →  Text data files (.txt)
    ↓ (Analysis results)
Output_folder/         →  Analysis outputs (.csv, .json, .txt, .xlsx)
    ↓ (Excel updating)
Updated_excel_workbooks/ → Updated Excel files (.xlsx)
```

## 📁 Project Structure

```
final_project/
├── 📂 Input_Folder/                    # Input Excel files
│   ├── ENGLISH_1A priedas_InoStartas en.xlsx
│   ├── ENGISH_1B priedas_InoStartas en.xlsx
│   ├── ENGISH_Rekomenduojama forma...xlsx
│   └── ENGLISH_Finansinis planas en.xlsx
├── 📂 DATA_SOURCES/                    # Text data for analysis
│   ├── ENGLISH_1A priedas_InoStartas en_1.txt
│   ├── ENGLISH_1A priedas_InoStartas en_2.txt
│   ├── ENGISH_1B priedas_InoStartas en_1_Fic_amounts_patenting.txt
│   └── ... (18 text files)
├── 📂 html_outputs/                    # Generated HTML files
│   ├── ENGLISH_1A priedas_InoStartas en_1.html
│   ├── ENGLISH_1A priedas_InoStartas en_2.html
│   └── ... (22 HTML files)
├── 📂 Output_folder/                   # Analysis results
│   ├── *.csv files                    # CSV format analysis
│   ├── *.json files                   # JSON format analysis
│   ├── *.txt files                    # Text summary
│   └── *.xlsx files                   # Excel format analysis
├── 📂 Updated_excel_workbooks/         # Updated Excel files
├── 📂 final_project/                   # Main system code
│   ├── run_complete_system.py         # Main entry point
│   ├── unified_processor.py           # Core processing classes
│   └── requirements.txt               # Dependencies
├── excel_processor.py                  # Legacy processor
├── final_json_from_outputFolder_to_xlsx_filling.py  # Excel updater
├── requirements.txt                    # Main dependencies
└── README.md                          # This file
```

## 🚀 Features

### Core Functionality
- **Multi-worksheet Support**: Processes Excel files with multiple worksheets
- **AI-Powered Analysis**: Uses OpenAI GPT models for intelligent data matching
- **Format Preservation**: Maintains Excel formatting, borders, alignment, and merged cells
- **Multiple Output Formats**: Generates CSV, JSON, Excel, and text outputs
- **Batch Processing**: Handles multiple files automatically
- **Error Handling**: Robust error handling with detailed logging

### Advanced Features
- **Contextual Data Matching**: AI analyzes table structure and matches data contextually
- **Cell Reference Mapping**: Provides Excel-style cell references (A1, B2, etc.)
- **Encoding Detection**: Handles multiple text encodings automatically
- **Retry Mechanism**: Built-in retry logic for API calls
- **Modular Design**: Clean, extensible architecture

## 🛠️ Installation

### Prerequisites
- Python 3.8 or higher
- OpenAI API key

### Setup

1. **Clone or download the project**
   ```bash
   git clone <repository-url>
   cd final_project
   ```

2. **Create virtual environment (recommended)**
   ```bash
   python -m venv formfillingvenv
   # Windows
   formfillingvenv\Scripts\activate
   # Linux/Mac
   source formfillingvenv/bin/activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Set up OpenAI API key**
   
   **Option A: Environment variable**
   ```bash
   # Windows
   set OPENAI_API_KEY=your-api-key-here
   
   # Linux/Mac
   export OPENAI_API_KEY=your-api-key-here
   ```
   
   **Option B: .env file**
   Create a `.env` file in the project root:
   ```
   OPENAI_API_KEY=your-api-key-here
   ```

## 📋 Dependencies

The system requires the following Python packages:

```
# Excel Processing
openpyxl==3.1.5          # Excel file manipulation
pandas==2.2.3            # Data analysis
xlsxwriter==3.2.0        # Excel writing

# AI Integration
openai==1.51.0           # OpenAI API client

# Utilities
python-dotenv==1.0.1     # Environment variable management
typing_extensions==4.12.2 # Type hints support
```

## 🎮 Usage

### Quick Start

1. **Place your Excel files** in the `Input_Folder/` directory
2. **Add corresponding text data files** in the `DATA_SOURCES/` directory
3. **Run the complete system**:
   ```bash
   python final_project/excel_processor.py
   ```

### Individual Components

#### 1. Excel to HTML Conversion
```python
from excel_processor import ExcelToHTMLConverter

converter = ExcelToHTMLConverter()
html_files = converter.process_all_excel_files("Input_Folder", "html_outputs")
```

#### 2. HTML Analysis with AI
```python
from excel_processor import HTMLAnalyzer

analyzer = HTMLAnalyzer()
success = analyzer.process_all_files()
```

#### 3. Excel File Updating
```python
from final_json_from_outputFolder_to_xlsx_filling import update_excel_from_json

update_excel_from_json("Input_Folder", "Output_folder")
```

### File Naming Convention

The system uses a specific naming convention for file matching:

- **Excel files**: `workbook_name.xlsx`
- **Text data files**: `workbook_name_worksheet_name.txt`
- **HTML files**: `workbook_name_worksheet_name.html`
- **Analysis outputs**: `workbook_name_worksheet_name.{csv,json,txt,xlsx}`

## 📊 Output Formats

### 1. CSV Files
Structured data with columns:
- `row`: Row number
- `column`: Column letter
- `cell_reference`: Excel cell reference (A1, B2, etc.)
- `value`: Data to fill
- `context`: Explanation of data placement
- `source_file`: Source file name

### 2. JSON Files
```json
[
  {
    "cell_reference": "B7",
    "value": "Patent application drafting and preparation...",
    "context": "Phase I attorney service (item 1.1) supplier / offer details"
  }
]
```

### 3. Excel Files
Analysis results in Excel format with auto-adjusted column widths and proper formatting.

### 4. Text Files
Human-readable summaries with:
- Source file information
- Number of cells filled
- Detailed cell mappings
- Generated file list

## 🔧 Configuration

### OpenAI Model Settings
The system uses `gpt-4o-mini` by default with the following parameters:
- `temperature`: 0.1 (for consistent results)
- `max_completion_tokens`: 16000
- `top_p`: 0.9
- `frequency_penalty`: 0.0
- `presence_penalty`: 0.0

### Retry Mechanism
- **Max retries**: 3 attempts
- **Error handling**: Graceful fallback to raw response saving

## 🐛 Troubleshooting

### Common Issues

1. **OpenAI API Key Not Found**
   ```
   Error: Please set your OPENAI_API_KEY environment variable
   ```
   **Solution**: Set your OpenAI API key as described in the Installation section.

2. **No Excel Files Found**
   ```
   Error: Input folder 'Input_Folder' not found
   ```
   **Solution**: Ensure the `Input_Folder` directory exists and contains `.xlsx` files.

3. **JSON Parsing Errors**
   ```
   Error parsing JSON response
   ```
   **Solution**: The system will save raw responses as fallback. Check your OpenAI API key and internet connection.

4. **Encoding Issues**
   ```
   UnicodeDecodeError
   ```
   **Solution**: The system automatically tries multiple encodings. Ensure your text files are properly encoded.

### Debug Mode
Enable detailed logging by modifying the print statements in the code or adding logging configuration.

## 📈 Performance

- **Processing Speed**: ~2-5 seconds per worksheet (depending on complexity)
- **Memory Usage**: Moderate (loads one Excel file at a time)
- **API Costs**: Varies based on OpenAI usage (typically $0.01-0.10 per file)

## 🔒 Security

- **API Key**: Store securely, never commit to version control
- **Data Privacy**: Text data is sent to OpenAI for analysis
- **File Access**: System only reads specified input directories

## 🤝 Contributing

The system is designed with modularity in mind. Key areas for extension:

1. **New AI Models**: Add support for other AI providers
2. **Additional Formats**: Support for other file formats
3. **Enhanced Analysis**: More sophisticated data matching algorithms
4. **UI Interface**: Web or desktop interface for easier use

## 📝 License

This project is designed for internal use. Please ensure compliance with OpenAI's usage policies and any applicable data protection regulations.

## 🆘 Support

For issues or questions:
1. Check the troubleshooting section above
2. Review the console output for detailed error messages
3. Ensure all dependencies are properly installed
4. Verify your OpenAI API key is valid and has sufficient credits

## 🔄 Workflow Summary

1. **Input**: Excel files with empty cells + corresponding text data files
2. **Step 1**: Convert Excel worksheets to HTML with preserved formatting
3. **Step 2**: Use AI to analyze HTML tables and match text data to cells
4. **Step 3**: Generate analysis results in multiple formats (CSV, JSON, Excel, TXT)
5. **Step 4**: Update original Excel files with filled data
6. **Output**: Updated Excel files with intelligently filled cells

The system is particularly effective for:
- Financial planning documents
- Project proposal forms
- Research data compilation
- Administrative form filling
- Data migration tasks

---

*This system represents a sophisticated approach to automated Excel processing, combining traditional data manipulation with modern AI capabilities to create an efficient, reliable solution for complex spreadsheet management tasks.*
