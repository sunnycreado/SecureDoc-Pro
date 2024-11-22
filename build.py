import PyInstaller.__main__
import os
import shutil
import sys

def build_exe():
    """Build a single executable for SecureDoc Pro"""
    
    print("Starting build process for SecureDoc Pro...")
    
    # Clean previous builds and dist
    for dir_to_clean in ['build', 'dist', '__pycache__']:
        if os.path.exists(dir_to_clean):
            shutil.rmtree(dir_to_clean)
    
    # Remove spec file if exists
    if os.path.exists('SecureDoc_Pro.spec'):
        os.remove('SecureDoc_Pro.spec')
    
    # PyInstaller configuration
    PyInstaller.__main__.run([
        'launcher.py',
        '--name=SecureDoc_Pro',
        '--onefile',
        '--console',
        '--icon=app/static/img/icon.ico',
        '--add-data=app/templates;app/templates',
        '--add-data=app/static;app/static',
        '--hidden-import=pythoncom',
        '--hidden-import=win32com.client',
        '--hidden-import=flask',
        '--hidden-import=werkzeug',
        '--hidden-import=waitress',
        '--clean',
        '--noconfirm',  # Remove output directory without confirmation
        '--distpath=.',  # Put the exe in current directory
        '--workpath=temp_build'  # Temporary build files
    ])
    
    # Clean up temporary build directory
    if os.path.exists('temp_build'):
        shutil.rmtree('temp_build')
    
    # Create uploads directory next to exe
    os.makedirs('uploads', exist_ok=True)
    
    print("\nBuild completed successfully!")
    print("\nExecutable created: SecureDoc_Pro.exe")
    print("Note: Closing the console window will terminate the application.")

if __name__ == "__main__":
    try:
        build_exe()
    except Exception as e:
        print(f"Build failed: {str(e)}")
        sys.exit(1) 