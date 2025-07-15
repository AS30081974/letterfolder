"""
Simple build script for Document Processor Pro
Creates a Windows executable with custom icon
"""
import os
import subprocess
import sys

def simple_build():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Paths
    script_path = "pdf_address_extractor_gui.py"
    icon_path = "app_icon.png"
    
    print("🔨 Building Document Processor Pro...")
    
    # Create ICO file from PNG
    try:
        from PIL import Image
        png_path = os.path.join(current_dir, icon_path)
        ico_path = os.path.join(current_dir, "app_icon.ico")
        
        if os.path.exists(png_path):
            img = Image.open(png_path)
            img.save(ico_path, format='ICO', sizes=[(32, 32), (48, 48)])
            print("✅ Created ICO file for Windows")
        else:
            print("⚠️ PNG icon not found")
            ico_path = None
    except Exception as e:
        print(f"⚠️ Could not create ICO: {e}")
        ico_path = None
    
    # Build command
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        "--name=DocumentProcessorPro",
        "--clean",
        "--noconfirm"
    ]
    
    if ico_path and os.path.exists(ico_path):
        cmd.extend(["--icon", ico_path])
    
    cmd.append(script_path)
    
    print(f"📦 Command: {' '.join(cmd)}")
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("✅ Build successful!")
        
        # Check if exe exists
        exe_path = os.path.join(current_dir, "dist", "DocumentProcessorPro.exe")
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"📦 Executable: {exe_path}")
            print(f"📏 Size: {size_mb:.1f} MB")
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Build failed: {e}")
        print(f"Output: {e.stdout}")
        print(f"Error: {e.stderr}")
    
    # Clean up
    if ico_path and os.path.exists(ico_path):
        try:
            os.remove(ico_path)
        except:
            pass

if __name__ == "__main__":
    simple_build()
