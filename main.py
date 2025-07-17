"""
Document Processor Pro
A modern application for extracting addresses from PDF files and headless printing.

Features:
- Extract addresses from PDF files containing "uk_team_gbmailgps@lilly.com"
- Export addresses to Excel format
- Headless PDF printing using Microsoft Edge
- Modern GUI with progress tracking
"""

import sys
import os
from gui.controller import AppController


def main():
    """Main application entry point."""
    try:
        # Create and run the application
        app = AppController()
        app.run()
        
    except KeyboardInterrupt:
        print("\nApplication interrupted by user.")
        sys.exit(0)
    except Exception as e:
        print(f"Application error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
