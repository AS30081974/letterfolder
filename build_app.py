# -*- coding: utf-8 -*-
"""
Build script to create a Windows executable for Document Processor Pro
This will create a standalone .exe file with the custom icon embedded
"""

from PyInstaller import __main__ as pyi_main
import os
import sys
import shutil

def build_executable():
    """Build the Windows executable with custom icon"""
    
    # Get the current directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Define paths
    script_path = os.path.join(current_dir, "pdf_address_extractor_gui.py")
    icon_path = os.path.join(current_dir, "app_icon.png")
    
    # Check if files exist
    if not os.path.exists(script_path):
        print("❌ Error: pdf_address_extractor_gui.py not found!")
        return False
        
    if not os.path.exists(icon_path):
        print("❌ Error: app_icon.png not found!")
        return False
    
    # Convert PNG to ICO for PyInstaller (if needed)
    try:
        from PIL import Image
        ico_path = os.path.join(current_dir, "app_icon.ico")
        
        # Convert PNG to ICO
        img = Image.open(icon_path)
        img.save(ico_path, format='ICO', sizes=[(16, 16), (32, 32), (48, 48), (64, 64)])
        print(f"✅ Created ICO file: {ico_path}")
        
    except Exception as e:
        print(f"⚠️ Warning: Could not create ICO file: {e}")
        ico_path = None
    
    # Prepare PyInstaller arguments
    args = [
        '--onefile',                    # Create a single executable file
        '--windowed',                   # Hide console window (GUI app)
        '--name=DocumentProcessorPro',  # Name of the executable
        '--clean',                      # Clean PyInstaller cache
        '--noconfirm',                  # Replace output directory without asking
    ]
    
    # Add icon if available
    if ico_path and os.path.exists(ico_path):
        args.extend(['--icon', ico_path])
    
    # Add data files to include
    args.extend([
        '--add-data', f'{icon_path};.',  # Include PNG icon in the bundle
    ])
    
    # Add the main script
    args.append(script_path)
    
    print("🔨 Building Windows executable...")
    print(f"📁 Script: {script_path}")
    print(f"🎨 Icon: {icon_path}")
    print(f"⚙️ PyInstaller args: {' '.join(args)}")
    
    try:
        # Run PyInstaller
        pyi_main.run(args)
        
        # Check if build was successful
        exe_path = os.path.join(current_dir, "dist", "DocumentProcessorPro.exe")
        if os.path.exists(exe_path):
            print(f"✅ Build successful!")
            print(f"📦 Executable created: {exe_path}")
            print(f"📏 File size: {os.path.getsize(exe_path) / (1024*1024):.1f} MB")
            return True
        else:
            print("❌ Build failed: Executable not found")
            return False
            
    except Exception as e:
        print(f"❌ Build error: {e}")
        return False
    
    finally:
        # Clean up temporary ICO file
        if ico_path and os.path.exists(ico_path):
            try:
                os.remove(ico_path)
                print("🧹 Cleaned up temporary ICO file")
            except:
                pass

def install_pyinstaller():
    """Install PyInstaller if not available"""
    try:
        import PyInstaller
        print("✅ PyInstaller already installed")
        return True
    except ImportError:
        print("📦 Installing PyInstaller...")
        import subprocess
        try:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])
            print("✅ PyInstaller installed successfully")
            return True
        except Exception as e:
            print(f"❌ Failed to install PyInstaller: {e}")
            return False

if __name__ == "__main__":
    print("🚀 Document Processor Pro - Build Script")
    print("=" * 50)
    
    # Install PyInstaller if needed
    if not install_pyinstaller():
        sys.exit(1)
    
    # Build the executable
    if build_executable():
        print("\n🎉 Build completed successfully!")
        print("📍 You can find the executable in the 'dist' folder")
        print("💡 The executable includes the custom icon and all dependencies")
    else:
        print("\n❌ Build failed!")
        sys.exit(1)
