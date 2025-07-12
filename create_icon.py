#!/usr/bin/env python3
"""
Script to create a simple application icon for Address Extractor Pro
This version creates a basic icon without requiring PIL
"""

import tkinter as tk
from tkinter import Canvas
import os

def create_simple_icon():
    """Create a simple icon using tkinter Canvas"""
    
    # Create a hidden tkinter window for drawing
    root = tk.Tk()
    root.withdraw()  # Hide the window
    
    # Create a canvas
    size = 64
    canvas = Canvas(root, width=size, height=size, bg='#2563eb')
    
    # Draw a document icon
    margin = 8
    doc_width = size - (margin * 2)
    doc_height = size - (margin * 2)
    
    # Document body
    canvas.create_rectangle(margin, margin + 8, margin + doc_width - 8, margin + doc_height, 
                           fill='white', outline='#e2e8f0', width=1)
    
    # Folded corner
    canvas.create_polygon(margin + doc_width - 16, margin + 8,
                         margin + doc_width - 8, margin + 16,
                         margin + doc_width - 8, margin + 8,
                         fill='#f1f5f9')
    
    # Address lines
    line_y = margin + 20
    for i in range(4):
        line_width = doc_width - 20 if i % 2 == 0 else doc_width - 28
        canvas.create_line(margin + 6, line_y, margin + line_width, line_y, 
                          fill='#64748b', width=2)
        line_y += 6
    
    # Excel symbol in corner
    excel_size = 12
    excel_x = margin + doc_width - excel_size - 4
    excel_y = margin + doc_height - excel_size - 4
    canvas.create_rectangle(excel_x, excel_y, excel_x + excel_size, excel_y + excel_size, 
                           fill='#059669', outline='white')
    
    # Grid lines in excel symbol
    for i in range(1, 3):
        x = excel_x + (excel_size // 3) * i
        canvas.create_line(x, excel_y, x, excel_y + excel_size, fill='white', width=1)
        y = excel_y + (excel_size // 3) * i
        canvas.create_line(excel_x, y, excel_x + excel_size, y, fill='white', width=1)
    
    canvas.pack()
    canvas.update()
    
    try:
        # Try to save as PostScript and convert (if available)
        ps_file = os.path.join(os.path.dirname(__file__), "temp_icon.ps")
        canvas.postscript(file=ps_file)
        print(f"✅ Created basic icon template: {ps_file}")
        print("Note: For a proper .ico file, consider using an online converter or image editing software.")
    except Exception as e:
        print(f"Could not save icon: {e}")
    
    root.destroy()

def main():
    print("Creating simple application icon...")
    print("This creates a basic icon template. For best results, use professional icon creation tools.")
    
    create_simple_icon()
    
    print("\n📋 Icon Implementation Notes:")
    print("1. The app now uses a building + clipboard emoji (🏢📋) as the logo")
    print("2. This represents address extraction and data organization")
    print("3. The window title has been updated to 'Address Extractor Pro'")
    print("4. For a custom .ico file, you can:")
    print("   - Use online converters like convertio.co or favicon.io")
    print("   - Use professional tools like GIMP, Photoshop, or IconEdit")
    print("   - Download icons from icon libraries like flaticon.com")

if __name__ == "__main__":
    main()
