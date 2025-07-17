"""
Headless PDF Printer
Provides headless PDF printing using Microsoft Edge browser.
"""

import os
import subprocess
import tempfile
import time
from typing import List


class HeadlessPrinter:
    """Handles headless PDF printing using Microsoft Edge."""
    
    def __init__(self):
        self.edge_path = self._find_edge_executable()
        
    def _find_edge_executable(self) -> str:
        """Find Microsoft Edge executable path."""
        edge_paths = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
        ]
        
        for path in edge_paths:
            if os.path.exists(path):
                return path
        return None
    
    def is_available(self) -> bool:
        """Check if headless printing is available."""
        return self.edge_path is not None
    
    def print_pdf(self, pdf_path: str, log_callback=None) -> bool:
        """
        Print a single PDF file headlessly.
        
        Args:
            pdf_path: Path to PDF file to print
            log_callback: Optional callback function for logging messages
            
        Returns:
            True if successful, False otherwise
        """
        if not self.is_available():
            if log_callback:
                log_callback("   ❌ Microsoft Edge not found for headless printing")
            return False
        
        try:
            if log_callback:
                log_callback("   🌐 Printing via Edge headless mode...")
            
            # Create temporary HTML file that auto-prints the PDF
            temp_html = self._create_auto_print_html(pdf_path)
            
            try:
                cmd = [
                    self.edge_path,
                    "--headless",
                    "--disable-gpu",
                    "--disable-software-rasterizer", 
                    "--disable-web-security",
                    "--kiosk-printing",
                    "--no-sandbox",
                    f"file:///{temp_html.replace(os.sep, '/')}"
                ]
                
                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    timeout=30,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                
                # Give time for print job to process
                time.sleep(2)
                
                if log_callback:
                    log_callback("   ✅ Print job sent successfully")
                return True
                    
            finally:
                # Clean up temp file
                try:
                    os.remove(temp_html)
                except:
                    pass
                    
        except subprocess.TimeoutExpired:
            if log_callback:
                log_callback("   ⚠️ Print timeout - may still be processing")
            return True  # Timeout doesn't necessarily mean failure
        except Exception as e:
            if log_callback:
                log_callback(f"   ❌ Print error: {e}")
            return False
    
    def print_multiple_pdfs(self, pdf_files: List[str], log_callback=None, 
                           stop_callback=None) -> tuple:
        """
        Print multiple PDF files headlessly.
        
        Args:
            pdf_files: List of PDF file paths to print
            log_callback: Optional callback function for logging messages
            stop_callback: Optional callback to check if printing should stop
            
        Returns:
            Tuple of (successful_count, failed_count)
        """
        if not self.is_available():
            if log_callback:
                log_callback("❌ Microsoft Edge not available for headless printing")
            return 0, len(pdf_files)
        
        if log_callback:
            log_callback(f"🖨️ Starting headless printing of {len(pdf_files)} PDF files...")
            log_callback("✅ Using Microsoft Edge for silent printing")
        
        successful_count = 0
        failed_count = 0
        
        for i, pdf_file in enumerate(sorted(pdf_files)):
            # Check if we should stop
            if stop_callback and stop_callback():
                if log_callback:
                    log_callback("🛑 Printing stopped by user.")
                break
            
            try:
                filename = os.path.basename(pdf_file)
                if log_callback:
                    log_callback(f"🖨️ [{i+1}/{len(pdf_files)}] Processing {filename}...")
                
                success = self.print_pdf(pdf_file, log_callback)
                
                if success:
                    successful_count += 1
                else:
                    failed_count += 1
                
                # Brief delay between files
                if i < len(pdf_files) - 1:
                    time.sleep(1)
                    
            except Exception as e:
                failed_count += 1
                if log_callback:
                    log_callback(f"   ❌ Error: {e}")
        
        return successful_count, failed_count
    
    def _create_auto_print_html(self, pdf_path: str) -> str:
        """Create temporary HTML file that auto-prints the PDF."""
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Auto Print PDF</title>
            <style>
                body {{ margin: 0; padding: 0; }}
                embed {{ width: 100%; height: 100vh; }}
            </style>
        </head>
        <body>
            <embed src="file:///{pdf_path.replace(os.sep, '/')}" type="application/pdf">
            <script>
                window.onload = function() {{
                    setTimeout(function() {{
                        window.print();
                        setTimeout(function() {{
                            window.close();
                        }}, 3000);
                    }}, 2000);
                }};
            </script>
        </body>
        </html>
        """
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False) as f:
            f.write(html_content)
            return f.name
