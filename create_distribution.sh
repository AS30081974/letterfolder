#!/bin/bash
# Document Processor Pro - Distribution Package Creator
# This script creates a user-friendly distribution package

echo "📦 Creating Document Processor Pro Distribution Package..."

# Create distribution directory
DIST_DIR="DocumentProcessorPro_Distribution"
mkdir -p "$DIST_DIR"

# Copy essential files for end users
echo "📁 Copying essential files..."

# Main executable (if it exists)
if [ -f "dist/DocumentProcessorPro.exe" ]; then
    cp "dist/DocumentProcessorPro.exe" "$DIST_DIR/"
    echo "✅ Added DocumentProcessorPro.exe"
else
    echo "⚠️ Warning: DocumentProcessorPro.exe not found in dist/"
    echo "   Run 'python simple_build.py' first to build the executable"
fi

# User documentation
cp "QUICK_START.txt" "$DIST_DIR/" 2>/dev/null && echo "✅ Added QUICK_START.txt"
cp "DOWNLOAD_GUIDE.md" "$DIST_DIR/" 2>/dev/null && echo "✅ Added DOWNLOAD_GUIDE.md"
cp "EASY_LAUNCH.bat" "$DIST_DIR/" 2>/dev/null && echo "✅ Added EASY_LAUNCH.bat"

# Sample files
if [ -d "sample_files" ]; then
    cp -r "sample_files" "$DIST_DIR/"
    echo "✅ Added sample_files directory"
fi

# Create a sample PDF if none exists
if [ ! -f "$DIST_DIR/sample_files/sample_letter.pdf" ]; then
    mkdir -p "$DIST_DIR/sample_files"
    echo "Creating sample PDF placeholder..."
    # Note: Would need a PDF creation tool here
fi

# Create ZIP package
echo "🗜️ Creating distribution ZIP..."
if command -v zip >/dev/null 2>&1; then
    zip -r "DocumentProcessorPro_v2.0.zip" "$DIST_DIR/"
    echo "✅ Created DocumentProcessorPro_v2.0.zip"
else
    echo "⚠️ ZIP utility not found. Package created in folder: $DIST_DIR"
fi

echo ""
echo "📊 Distribution Package Summary:"
echo "================================"
ls -la "$DIST_DIR/"

echo ""
echo "📦 Package Ready!"
echo "Share 'DocumentProcessorPro_v2.0.zip' with end users"
echo "They just need to extract and run DocumentProcessorPro.exe"
echo ""
echo "📋 End User Instructions:"
echo "1. Extract the ZIP file"
echo "2. Double-click DocumentProcessorPro.exe"
echo "3. Follow the Quick Start guide"
