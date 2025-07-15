# Document Processor Pro - Quick Reference

## 🚀 Quick Start Guide

### For End Users (Windows Executable)
1. Double-click `DocumentProcessorPro.exe`
2. Select PDF folder
3. Choose Excel output file
4. Click "Process Documents"

### For Developers (Python Source)
1. Install Python 3.7+
2. Run: `pip install -r requirements.txt`
3. Run: `python pdf_address_extractor_gui.py`

## 📦 File Overview

| File | Purpose | Required |
|------|---------|----------|
| `DocumentProcessorPro.exe` | Main Windows executable | ✅ End users |
| `pdf_address_extractor_gui.py` | Main Python application | ✅ Developers |
| `app_icon.png` | Custom application icon | ✅ For branding |
| `requirements.txt` | Python dependencies | ✅ Developers |
| `simple_build.py` | Build executable script | 🔧 Building |
| `run_document_processor.bat` | Windows launcher | 🔧 Convenience |

## 🎨 Custom Icon Features

### What You Get
- ✅ Professional window branding
- ✅ Custom taskbar icon
- ✅ Alt+Tab integration
- ✅ No more tkinter "Tk" icon

### How It Works
- Loads `app_icon.png` automatically
- Converts to multiple sizes (16x16, 32x32, 48x48)
- Fallback to clean icon removal if loading fails
- Windows ICO conversion for best compatibility

## 🖨️ Printing Modes

| Mode | When Used | Speed | User Impact |
|------|-----------|-------|-------------|
| Background | Adobe installed | Fast (2-3s/file) | No disruption |
| Visible | No Adobe/fallback | Slow (10-15s/file) | Don't use PC |

## 🔧 Common Commands

### Development
```bash
# Run application
python pdf_address_extractor_gui.py

# Test icon loading
python test_icon.py

# Build executable
python simple_build.py

# Install dependencies
pip install -r requirements.txt
```

### Troubleshooting
```bash
# Check Python version
python --version

# Check if Pillow is installed
python -c "import PIL; print(PIL.__version__)"

# Test icon file exists
python -c "import os; print(os.path.exists('app_icon.png'))"
```

## 📋 Checklist for New Installations

### End User Setup
- [ ] Download/extract files
- [ ] Verify `DocumentProcessorPro.exe` is present
- [ ] Double-click to run
- [ ] Check custom icon appears

### Developer Setup
- [ ] Clone repository
- [ ] Install Python 3.7+
- [ ] Run `pip install -r requirements.txt`
- [ ] Verify `app_icon.png` exists
- [ ] Run `python pdf_address_extractor_gui.py`
- [ ] Confirm custom icon loads

### Building Executable
- [ ] Install PyInstaller: `pip install pyinstaller`
- [ ] Ensure `app_icon.png` is present
- [ ] Run `python simple_build.py`
- [ ] Check `dist/DocumentProcessorPro.exe` created
- [ ] Test executable on clean system

## 🎯 Key Features Summary

### Document Processing
- Extracts text between email and "Dear" markers
- Saves to Excel with auto-formatting
- Batch processes multiple PDFs
- Real-time progress tracking

### Professional Branding
- Custom app icon (replaces tkinter default)
- Modern card-based UI design
- Professional window titles
- Consistent visual branding

### Smart Printing
- Automatic Adobe detection
- Background printing when possible
- Visible printing fallback
- Print job management

### Distribution Ready
- Standalone 36MB executable
- No Python installation required
- Portable - runs from any location
- Custom icon embedded in .exe

## 🆘 Quick Fixes

### Icon Not Showing
1. Check `app_icon.png` exists in same folder
2. Install Pillow: `pip install Pillow`
3. Run as administrator if permissions issue

### Executable Won't Run
1. Add Windows Defender exception
2. Install Visual C++ Redistributable
3. Run from Command Prompt to see errors

### Printing Issues
1. Install Adobe Reader for background mode
2. Set Adobe as default PDF handler
3. Check printer is configured correctly

---

*For detailed documentation, see README.md*
