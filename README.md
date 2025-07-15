# 📄 Document Processor Pro

A modern, professional desktop application for processing PDF documents and extracting structured data to Excel spreadsheets. Built with Python and Tkinter, featuring intelligent printing capabilities with Adobe integration, custom branding, and standalone Windows executable support.💼 Document Processor Pro

A modern, user-friendly desktop application for processing PDF documents and extracting structured data to Excel spreadsheets. Built with Python and Tkinter, featuring intelligent printing capabilities with Adobe integration.

## 📋 Table of Contents

- [Features](#features)
- [Installation & Quick Start](#installation--quick-start)
- [Windows Executable](#windows-executable)
- [Custom Icon Features](#custom-icon-features)
- [Usage Guide](#usage-guide)
- [Printing Modes](#printing-modes)
- [Technical Documentation](#technical-documentation)
- [Build & Distribution](#build--distribution)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

## ✨ Features

### Core Functionality
- **📄 PDF Text Extraction**: Automatically extracts structured data from PDF files using pattern matching
- **📊 Excel Integration**: Saves extracted data to Excel spreadsheets with auto-formatting
- **🔍 Batch Processing**: Process multiple PDF files in a single operation
- **🧹 Data Management**: Clear and manage Excel spreadsheet data

### Professional Branding & Distribution
- **🎨 Custom Icon**: Professional branding with custom app icon (app_icon.png)
- **📦 Windows Executable**: Standalone .exe file (36MB) with embedded dependencies
- **🏢 Professional UI**: Modern card-based interface with Document Processor Pro branding
- **⚡ No Installation Required**: Portable executable runs without Python installation

### Smart Printing System
- **🤖 Intelligent Mode Selection**: Automatically chooses optimal printing method
- **🔄 Background Printing**: Silent, non-intrusive printing when Adobe is available
- **👁️ Visible Printing**: Automated visible printing with keyboard simulation
- **🛡️ Fallback Protection**: Automatic fallback from background to visible mode on errors
- **⏹️ Process Control**: Start/stop printing operations with real-time feedback

### Modern User Interface
- **� Professional Design**: Clean, card-based interface with custom icon branding
- **📱 Responsive Layout**: Resizable window with adaptive components
- **📈 Progress Tracking**: Real-time progress indicators and status logging
- **🖼️ Custom Branding**: Replaces default tkinter icon with professional app icon

## 🛠️ System Requirements

### For Windows Executable (Recommended)
- **Operating System**: Windows 7/8/10/11 (64-bit)
- **Memory**: 1GB RAM minimum
- **Storage**: 50MB free space
- **Dependencies**: None (all included in executable)

### For Python Source Code
- **Operating System**: Windows 7/8/10/11
- **Python**: 3.7 or higher
- **Memory**: 2GB RAM minimum
- **Storage**: 50MB free space

### Software Dependencies (Python version only)
```
tkinter (built-in with Python)
PyPDF2>=3.0.0
openpyxl>=3.1.0
Pillow>=9.0.0 (for custom icon support)
pyinstaller>=5.0.0 (for building executable)
```

### Optional (for enhanced printing)
- **Adobe Reader DC** or **Adobe Acrobat** (any version)
- **XWCSmartPrint** or compatible printer driver with stapling support

## 🚀 Installation & Quick Start

### Method 1: Ready-to-Use Package (Recommended for End Users)
**Zero Installation Required!**

1. **Download**: [DocumentProcessorPro.zip](./DocumentProcessorPro.zip) (~37MB)
2. **Extract**: Right-click ZIP → "Extract All" to any location
3. **Run**: Double-click `DocumentProcessorPro.exe`
4. **Start Processing**: Select PDF folder → Process Documents

**What You Get:**
- ✅ Standalone executable (no Python needed)
- ✅ Custom professional icon
- ✅ Quick start guide included
- ✅ Sample files for testing
- ✅ Works on any Windows 7+ computer

### Method 2: Python Source Code (For Developers)
1. **Clone Repository**
   ```bash
   git clone https://github.com/AS30081974/letterfolder.git
   cd letterfolder
   pip install -r requirements.txt
   ```

2. **Run Application**
   ```bash
   python pdf_address_extractor_gui.py
   ```

### Method 3: Build Your Own Executable
1. **Setup Environment**
   ```bash
   git clone https://github.com/AS30081974/letterfolder.git
   cd letterfolder
   pip install -r requirements.txt
   pip install pyinstaller
   ```

2. **Build & Package**
   ```bash
   python simple_build.py           # Build executable
   create_distribution.bat          # Create user-friendly package
   ```

3. **Distribute**
   - Share `DocumentProcessorPro.zip` with end users
   - They just extract and run - no installation needed!

## � Windows Executable

### Features
- **Standalone**: No Python installation required
- **Portable**: Copy and run on any Windows computer
- **Custom Icon**: Professional branding with your app icon
- **Size**: Approximately 36MB with all dependencies
- **Performance**: Optimized for fast startup and operation

### What's Included
- Python 3.11 runtime
- All required libraries (tkinter, PyPDF2, openpyxl, Pillow)
- Custom icon assets (app_icon.png embedded)
- Windows-specific optimizations

### Distribution
- Share the `DocumentProcessorPro.exe` file
- No additional files needed
- Works on any Windows 7+ system
- Can be run from USB drives or network locations

## 🎨 Custom Icon Features

### Professional Branding
The application now features custom icon support that replaces the default tkinter "Tk" icon with your professional branding:

- **📄 app_icon.png**: Your custom application icon
- **🖼️ Multiple Formats**: Supports PNG input, converts to ICO for Windows
- **📐 Multiple Sizes**: Automatically creates 16x16, 32x32, and 48x48 versions
- **🪟 Windows Integration**: Appears in title bar, taskbar, and Alt+Tab

### Icon Display Locations
✅ Window title bar  
✅ Windows taskbar  
✅ Alt+Tab application switcher  
✅ System tray (when minimized)  
✅ Windows Explorer (for .exe file)  

### Implementation Details
```python
# Automatic icon loading with fallback protection
try:
    pil_image = Image.open("app_icon.png")
    pil_image_32 = pil_image.resize((32, 32), Image.Resampling.LANCZOS)
    icon_photo = ImageTk.PhotoImage(pil_image_32)
    root.iconphoto(True, icon_photo)
    
    # Windows ICO conversion for better compatibility
    pil_image.save("temp_icon.ico", format='ICO', sizes=[(16, 16), (32, 32), (48, 48)])
    root.iconbitmap("temp_icon.ico")
except Exception:
    # Graceful fallback - removes default tkinter icon
    root.iconbitmap(default="")
```

## 📖 Usage Guide

### Document Processing Workflow

The application processes documents using intelligent pattern matching:

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

| PDF File | Extracted Data |
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

#### `DocumentProcessorGUI`
Main application class handling UI, icon management, and orchestration.

**Key Methods:**
- `__init__()`: Initialize UI with custom icon loading
- `extract_text_from_pdfs()`: Core extraction logic
- `print_pdfs()`: Main printing coordinator
- `print_single_pdf_background()`: Background printing implementation
- `print_single_pdf_visible()`: Visible printing with automation
- `find_adobe_reader()`: Adobe installation detection

#### Custom Icon System
```python
# Icon loading with PIL/Pillow
script_dir = os.path.dirname(os.path.abspath(__file__))
icon_path = os.path.join(script_dir, "app_icon.png")

if os.path.exists(icon_path):
    pil_image = Image.open(icon_path)
    pil_image_32 = pil_image.resize((32, 32), Image.Resampling.LANCZOS)
    self.icon_photo = ImageTk.PhotoImage(pil_image_32)
    self.root.iconphoto(True, self.icon_photo)
    
    # Windows ICO conversion
    ico_path = os.path.join(script_dir, "temp_icon.ico")
    pil_image.save(ico_path, format='ICO', sizes=[(16, 16), (32, 32), (48, 48)])
    self.root.iconbitmap(ico_path)
    os.remove(ico_path)  # Cleanup
```

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
├── app_icon.png                  # Custom application icon
├── README.md                     # This documentation
├── requirements.txt              # Python dependencies
├── simple_build.py               # Build script for executable
├── build_app.py                  # Advanced build script
├── run_document_processor.bat    # Windows batch file
├── test_icon.py                  # Icon testing utility
├── .gitignore                   # Git exclusions
├── dist/                        # Built executables
│   └── DocumentProcessorPro.exe # Windows executable (36MB)
├── build/                       # Build artifacts
└── [Local files - not in repo]
    ├── *.pdf                    # PDF documents
    ├── *.xlsx                   # Excel spreadsheets
    └── extracted_data.xlsx      # Default output file
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
5. **Icon Fallback**: Removes default tkinter icon if custom icon loading fails

## 🔨 Build & Distribution

### Building Windows Executable

The application includes build scripts to create standalone Windows executables:

#### Simple Build Process
```bash
# Install build dependencies
pip install pyinstaller pillow

# Run the build script
python simple_build.py
```

#### Build Output
- **Location**: `dist/DocumentProcessorPro.exe`
- **Size**: ~36MB (includes Python runtime and all dependencies)
- **Features**: Custom icon, professional branding, no external dependencies

#### Build Configuration
The build process:
1. **Icon Conversion**: Converts `app_icon.png` to ICO format for Windows
2. **Dependency Bundling**: Includes all required Python packages
3. **Optimization**: Creates single-file executable for easy distribution
4. **Asset Embedding**: Includes custom icon and resources

#### Advanced Build Options
```bash
# Custom build with specific options
python build_app.py

# Manual PyInstaller command
pyinstaller --onefile --windowed --icon=app_icon.ico --name=DocumentProcessorPro pdf_address_extractor_gui.py
```

#### Build Scripts Included
- **`simple_build.py`**: Basic build with automatic configuration
- **`build_app.py`**: Advanced build with detailed logging and options
- **`run_document_processor.bat`**: Windows batch file for Python version

### Distribution Best Practices

#### For End Users
1. **Distribute**: Only the `DocumentProcessorPro.exe` file
2. **Requirements**: Windows 7+ (64-bit recommended)
3. **Size**: Plan for ~36MB download/storage
4. **Installation**: No installation required - just run the executable

#### For Developers
1. **Source Code**: Clone the full repository
2. **Dependencies**: Install via `pip install -r requirements.txt`
3. **Icon Assets**: Ensure `app_icon.png` is present for custom branding
4. **Build Environment**: Windows recommended for building Windows executables

### Version Management
- **Executable Versioning**: Update version info in build scripts
- **Icon Updates**: Replace `app_icon.png` and rebuild
- **Dependency Updates**: Update `requirements.txt` and rebuild

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

#### Custom Icon Issues

**Problem**: Custom icon not displaying
- **Symptoms**: Default tkinter "Tk" icon still showing
- **Causes**:
  - `app_icon.png` file missing or corrupted
  - PIL/Pillow not installed
  - Insufficient file permissions
- **Solutions**:
  - Verify `app_icon.png` exists in application directory
  - Install Pillow: `pip install Pillow`
  - Check file permissions
  - Use test script: `python test_icon.py`

**Problem**: Icon appears blurry or distorted
- **Cause**: Low resolution source image or scaling issues
- **Solution**: 
  - Use high-resolution PNG (256x256 or higher)
  - Ensure square aspect ratio
  - Test with different icon sizes

#### Executable Issues

**Problem**: Executable won't run
- **Symptoms**: Nothing happens when double-clicking .exe
- **Causes**:
  - Windows Defender blocking
  - Missing Visual C++ redistributables
  - Corrupted download
- **Solutions**:
  - Add exception in Windows Defender
  - Install Microsoft Visual C++ Redistributable
  - Re-download and verify file integrity

**Problem**: Executable is very slow to start
- **Cause**: Windows Defender real-time scanning
- **Solution**: Add executable folder to Defender exclusions

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
   pip install pyinstaller  # For building executables
   ```

### Code Style Guidelines

- **Python**: Follow PEP 8 style guidelines
- **Comments**: Document complex logic and Windows API calls
- **Error Handling**: Include comprehensive exception handling
- **Logging**: Use descriptive log messages with appropriate icons
- **Icon Assets**: Include proper attribution for custom icons

### Testing

#### Manual Testing Checklist
- [ ] Custom icon displays correctly in all contexts
- [ ] Address extraction with various PDF formats
- [ ] Excel file creation and updates
- [ ] Background printing (with Adobe)
- [ ] Visible printing (without Adobe)
- [ ] Fallback functionality
- [ ] UI responsiveness during operations
- [ ] Error handling scenarios
- [ ] Executable builds successfully
- [ ] Executable runs on clean Windows system

#### Test Environment Setup
- Windows VM with/without Adobe
- Various PDF formats and structures
- Different printer configurations
- Permission-restricted environments
- Systems without Python installed (for executable testing)

### Building and Testing Executables

1. **Test Icon Loading**
   ```bash
   python test_icon.py
   ```

2. **Build Executable**
   ```bash
   python simple_build.py
   ```

3. **Test Executable**
   - Run on development machine
   - Test on clean Windows VM
   - Verify icon appears correctly
   - Test all functionality

### Icon Design Guidelines

- **Format**: PNG with transparency support
- **Size**: 256x256 pixels minimum
- **Aspect Ratio**: Square (1:1)
- **Content**: Clear, professional design
- **Colors**: High contrast for visibility at small sizes

### Submitting Changes

1. **Create Feature Branch**: `git checkout -b feature/description`
2. **Make Changes**: Implement and test thoroughly
3. **Update Documentation**: Update README if needed
4. **Submit Pull Request**: Include detailed description of changes

## 📄 License

This project is licensed under the MIT License - see below for details:

```
MIT License

Copyright (c) 2025 Document Processor Pro

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
- **README.md**: Comprehensive usage and technical documentation
- **QUICK_REFERENCE.md**: Quick start guide and common commands
- **Code Comments**: Inline documentation for complex functions
- **Status Log**: Real-time operation feedback in the application

### Files and Resources
- **Windows Executable**: `dist/DocumentProcessorPro.exe` - Standalone application
- **Custom Icon**: `app_icon.png` - Professional branding asset
- **Build Scripts**: `simple_build.py`, `build_app.py` - For creating executables
- **Test Utilities**: `test_icon.py` - For verifying icon functionality

### Community
- **GitHub Issues**: Report bugs and request features
- **Discussions**: Share tips and best practices
- **Wiki**: Extended documentation and examples

### Enterprise Support
For enterprise deployments and custom integrations, contact the development team through GitHub issues with the "enterprise" label.

---

**Built with ❤️ for efficient document processing**

*Document Processor Pro - Professional document processing with custom branding*

## 📊 Version History

### v2.0 - Professional Edition (Current)
- ✅ Custom icon support with professional branding
- ✅ Standalone Windows executable (36MB)
- ✅ Modern UI with card-based design
- ✅ Enhanced error handling and fallback systems
- ✅ Build automation and distribution tools
- ✅ Comprehensive documentation updates

### v1.0 - Initial Release
- ✅ PDF text extraction and Excel output
- ✅ Intelligent printing with Adobe integration
- ✅ Modern tkinter GUI
- ✅ Background and visible printing modes

## 🎯 Roadmap

### Planned Features
- 🔄 Multi-language support
- 📱 Configuration file for custom patterns
- 🔍 Advanced PDF text recognition
- 📊 Enhanced Excel formatting options
- 🔐 Digital signature support
- 🌐 Web-based version

### Community Requests
- ⚙️ Settings dialog for configuration
- 📈 Processing statistics and reporting
- 🔔 Email notifications for batch completion
- 📁 Network folder support

*Suggestions and contributions welcome via GitHub Issues*
