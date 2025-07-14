# 🏢📋 Address Extractor Pro

A modern, user-friendly desktop application for extracting addresses from PDF files and saving them to Excel spreadsheets. Built with Python and Tkinter, featuring intelligent printing capabilities with Adobe integration.

## 📋 Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Usage Guide](#usage-guide)
- [Printing Modes](#printing-modes)
- [Technical Documentation](#technical-documentation)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

## ✨ Features

### Core Functionality
- **📄 PDF Text Extraction**: Automatically extracts addresses from PDF files using pattern matching
- **📊 Excel Integration**: Saves extracted addresses to Excel spreadsheets with auto-formatting
- **🔍 Batch Processing**: Process multiple PDF files in a single operation
- **🧹 Data Management**: Clear and manage Excel spreadsheet data

### Smart Printing System
- **🤖 Intelligent Mode Selection**: Automatically chooses optimal printing method
- **🔄 Background Printing**: Silent, non-intrusive printing when Adobe is available
- **👁️ Visible Printing**: Automated visible printing with keyboard simulation
- **🛡️ Fallback Protection**: Automatic fallback from background to visible mode on errors
- **⏹️ Process Control**: Start/stop printing operations with real-time feedback

### Modern User Interface
- **🎨 Modern Design**: Clean, card-based interface with professional styling
- **📱 Responsive Layout**: Resizable window with adaptive components
- **📈 Progress Tracking**: Real-time progress indicators and status logging
- **🎯 User-Friendly**: Intuitive workflow with clear instructions

## 🛠️ Requirements

### System Requirements
- **Operating System**: Windows 7/8/10/11
- **Python**: 3.7 or higher
- **Memory**: 2GB RAM minimum
- **Storage**: 50MB free space

### Software Dependencies
```
tkinter (built-in with Python)
PyPDF2>=3.0.0
openpyxl>=3.1.0
```

### Optional (for enhanced printing)
- **Adobe Reader DC** or **Adobe Acrobat** (any version)
- **XWCSmartPrint** or compatible printer driver with stapling support

## 🚀 Installation

### Method 1: Clone Repository
```bash
git clone https://github.com/AS30081974/letterfolder.git
cd letterfolder
pip install -r requirements.txt
```

### Method 2: Download and Setup
1. Download the ZIP file from GitHub
2. Extract to your desired location
3. Install dependencies:
```bash
pip install PyPDF2 openpyxl
```

### Method 3: Virtual Environment (Recommended)
```bash
git clone https://github.com/AS30081974/letterfolder.git
cd letterfolder
python -m venv address_extractor_env
address_extractor_env\Scripts\activate  # Windows
pip install -r requirements.txt
```

## 🏃‍♂️ Quick Start

1. **Launch the Application**
   ```bash
   python pdf_address_extractor_gui.py
   ```

2. **Select PDF Folder**
   - Click "📂 Browse" next to "PDF Folder"
   - Choose the folder containing your PDF files

3. **Set Excel Output**
   - Specify the Excel file name (default: `addresses.xlsx`)
   - File will be created in the same directory if it doesn't exist

4. **Extract Addresses**
   - Click "🔍 Extract Addresses"
   - Monitor progress in the status log

5. **Print Documents** (Optional)
   - Click "🖨️ Print PDFs"
   - Follow the setup instructions in the confirmation dialog

## 📖 Usage Guide

### Address Extraction Process

The application extracts addresses using pattern matching:

1. **Pattern Detection**: Searches for text between `uk_team_gbmailgps@lilly.com` and `Dear`
2. **Text Processing**: Cleans and formats the extracted text
3. **Excel Output**: Saves results with auto-formatted columns

#### Example PDF Structure
```
... other content ...
uk_team_gbmailgps@lilly.com

[Company Information]
John Smith
123 Main Street
Anytown, AN 12345
United Kingdom

Dear [Recipient]
... rest of content ...
```

#### Extracted Result
The application will extract:
```
John Smith
123 Main Street
Anytown, AN 12345
United Kingdom
```

### Excel Output Format

| PDF File | Extracted Address |
|----------|-------------------|
| document1.pdf | John Smith<br>123 Main Street<br>Anytown, AN 12345 |
| document2.pdf | Jane Doe<br>456 Oak Avenue<br>Another City, AC 67890 |

### File Management

- **Existing Files**: Appends new data to existing Excel files
- **Column Sizing**: Automatically adjusts column widths for readability
- **Headers**: Adds appropriate headers if the file is new or empty

## 🖨️ Printing Modes

The application features an intelligent dual-mode printing system:

### Background Mode (Preferred)
**When Available**: Adobe Reader or Acrobat is installed
- ✅ **Silent Operation**: No windows open during printing
- ✅ **Continue Working**: Use your computer normally while printing
- ✅ **Fast Processing**: 2-3 seconds per file
- ✅ **Command Line**: Uses Adobe's `/t` parameter for direct printing

### Visible Mode (Fallback)
**When Used**: Adobe not available or background mode fails
- 👁️ **Visual Feedback**: See each PDF open and print
- ⌨️ **Keyboard Automation**: Automated Ctrl+P and Enter key presses
- ⏱️ **Longer Processing**: 10-15 seconds per file
- ⚠️ **User Restriction**: Don't use keyboard/mouse during operation

### Automatic Fallback System
- **Smart Detection**: Automatically tries background mode first
- **Error Handling**: Falls back to visible mode if background fails
- **Detailed Logging**: Reports why fallback occurred
- **Seamless Operation**: No user intervention required

### Printer Setup (XWCSmartPrint)
For optimal results with stapling:

1. Open **Settings** → **Printers & Scanners**
2. Select **XWCSmartPrint**
3. Click **Print Properties**
4. Under **Presets**, select **"1 Staple, 2-Sided"**
5. Click **Apply** and **OK**

## 🔧 Technical Documentation

### Architecture Overview

```
Address Extractor Pro
├── GUI Layer (tkinter/ttk)
│   ├── Modern Styling System
│   ├── Card-based Layout
│   └── Progress Tracking
├── Core Processing
│   ├── PDF Text Extraction (PyPDF2)
│   ├── Pattern Matching Engine
│   └── Excel Integration (openpyxl)
├── Printing System
│   ├── Adobe Detection
│   ├── Background Printing
│   ├── Visible Printing with Automation
│   └── Fallback Management
└── Utilities
    ├── Windows API Integration
    ├── Process Management
    └── Error Handling
```

### Key Classes and Methods

#### `PDFAddressExtractorGUI`
Main application class handling UI and orchestration.

**Key Methods:**
- `extract_text_from_pdfs()`: Core extraction logic
- `print_pdfs()`: Main printing coordinator
- `print_single_pdf_background()`: Background printing implementation
- `print_single_pdf_visible()`: Visible printing with automation
- `find_adobe_reader()`: Adobe installation detection

#### Adobe Integration
```python
# Background printing command
[adobe_path, '/t', pdf_path]

# Common Adobe paths checked:
- C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe
- C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe
# ... and registry lookups
```

#### Windows API Usage
```python
# Keyboard automation for visible printing
VK_CONTROL = 0x11  # Ctrl key
VK_P = 0x50       # P key
VK_RETURN = 0x0D  # Enter key
VK_MENU = 0x12    # Alt key
VK_F4 = 0x73      # F4 key
```

### File Structure
```
letterfolder/
├── pdf_address_extractor_gui.py  # Main application
├── README.md                     # This documentation
├── requirements.txt              # Python dependencies
├── .gitignore                   # Git exclusions
└── [Local files - not in repo]
    ├── *.pdf                    # PDF documents
    ├── *.xlsx                   # Excel spreadsheets
    └── app_icon.ico             # Optional app icon
```

### Threading Model
- **Main Thread**: GUI operations and user interaction
- **Worker Threads**: Address extraction and printing operations
- **Thread Safety**: Proper synchronization for UI updates

### Error Handling Strategy
1. **Graceful Degradation**: Application continues functioning when non-critical features fail
2. **Detailed Logging**: Comprehensive error reporting in the status log
3. **User Feedback**: Clear error messages and recovery suggestions
4. **Automatic Recovery**: Fallback mechanisms for printing failures

## 🐛 Troubleshooting

### Common Issues and Solutions

#### PDF Extraction Issues

**Problem**: No addresses extracted from PDFs
- **Cause**: PDF doesn't contain the expected pattern markers
- **Solution**: Verify PDFs contain `uk_team_gbmailgps@lilly.com` and `Dear` markers
- **Alternative**: Check if text is in images (not extractable)

**Problem**: Partial address extraction
- **Cause**: Unexpected formatting in PDF
- **Solution**: Review the pattern matching logic for edge cases

#### Printing Issues

**Problem**: Background printing fails immediately
- **Symptoms**: Immediately falls back to visible mode
- **Causes**:
  - Adobe not properly installed
  - Printer not configured
  - PDF file permissions
- **Solutions**:
  - Reinstall Adobe Reader/Acrobat
  - Check printer setup
  - Verify file permissions

**Problem**: Visible printing doesn't work
- **Symptoms**: Windows open but don't print
- **Causes**:
  - Default PDF handler not set
  - Keyboard automation blocked
- **Solutions**:
  - Set Adobe as default PDF handler
  - Run as administrator
  - Disable keyboard blocking software

**Problem**: Printing stops unexpectedly
- **Cause**: User interruption or system resource issues
- **Solution**: Use the "⏹️ Stop Printing" button and restart

#### Excel Issues

**Problem**: Excel file not created/updated
- **Cause**: File permissions or path issues
- **Solution**: 
  - Check write permissions
  - Close Excel if file is open
  - Verify folder exists

**Problem**: Formatting issues in Excel
- **Cause**: Large address text or special characters
- **Solution**: Manually adjust column widths if needed

### Performance Optimization

#### For Large Batches
- **Close unnecessary applications** during processing
- **Ensure sufficient disk space** for Excel files
- **Use background printing mode** when possible
- **Process in smaller batches** if memory issues occur

#### For Slow Systems
- **Increase delays** between print operations
- **Close other programs** during printing
- **Use SSD storage** for better performance

### Adobe Detection Issues

If Adobe is installed but not detected:

1. **Check Installation Paths**: Verify Adobe is in standard locations
2. **Registry Check**: Ensure proper registry entries exist
3. **Permissions**: Run application as administrator
4. **Reinstall Adobe**: Fresh installation may resolve issues

### Getting Help

1. **Check Status Log**: Detailed error information is logged
2. **Enable Debug Mode**: Modify logging level for more details
3. **System Information**: Note Windows version, Python version, Adobe version
4. **Reproduce Steps**: Document exact steps that cause issues

## 🤝 Contributing

### Development Setup

1. **Fork the Repository**
2. **Create Virtual Environment**
   ```bash
   python -m venv dev_env
   dev_env\Scripts\activate
   ```
3. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   pip install -r dev-requirements.txt  # If available
   ```

### Code Style Guidelines

- **Python**: Follow PEP 8 style guidelines
- **Comments**: Document complex logic and Windows API calls
- **Error Handling**: Include comprehensive exception handling
- **Logging**: Use descriptive log messages with appropriate icons

### Testing

#### Manual Testing Checklist
- [ ] Address extraction with various PDF formats
- [ ] Excel file creation and updates
- [ ] Background printing (with Adobe)
- [ ] Visible printing (without Adobe)
- [ ] Fallback functionality
- [ ] UI responsiveness during operations
- [ ] Error handling scenarios

#### Test Environment Setup
- Windows VM with/without Adobe
- Various PDF formats and structures
- Different printer configurations
- Permission-restricted environments

### Submitting Changes

1. **Create Feature Branch**: `git checkout -b feature/description`
2. **Make Changes**: Implement and test thoroughly
3. **Update Documentation**: Update README if needed
4. **Submit Pull Request**: Include detailed description of changes

## 📄 License

This project is licensed under the MIT License - see below for details:

```
MIT License

Copyright (c) 2025 Address Extractor Pro

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## 📞 Support

### Documentation
- **README**: Comprehensive usage and technical documentation
- **Code Comments**: Inline documentation for complex functions
- **Status Log**: Real-time operation feedback in the application

### Community
- **GitHub Issues**: Report bugs and request features
- **Discussions**: Share tips and best practices
- **Wiki**: Extended documentation and examples

### Enterprise Support
For enterprise deployments and custom integrations, contact the development team through GitHub issues with the "enterprise" label.

---

**Built with ❤️ for efficient document processing**

*Address Extractor Pro - Making address extraction and printing effortless*
