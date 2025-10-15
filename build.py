#!/usr/bin/env python3
"""
Cross-platform build script for Excel Handler
Works on Windows, Linux, and Mac
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

# Set UTF-8 encoding for Windows console
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

def print_header(message):
    """Print a formatted header."""
    print("\n" + "=" * 60)
    print(f"  {message}")
    print("=" * 60 + "\n")

def print_success(message):
    """Print a success message."""
    print(f"[OK] {message}")

def print_error(message):
    """Print an error message."""
    print(f"[ERROR] {message}")

def print_info(message):
    """Print an info message."""
    print(f"  {message}")

def check_python_version():
    """Check if Python version is adequate."""
    if sys.version_info < (3, 7):
        print_error("Python 3.7 or higher is required")
        print_info(f"Current version: {sys.version}")
        return False
    print_success(f"Python version: {sys.version_info.major}.{sys.version_info.minor}")
    return True

def install_dependencies():
    """Install required dependencies."""
    print_info("Installing dependencies...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print_success("Dependencies installed")
        return True
    except subprocess.CalledProcessError as e:
        print_error(f"Failed to install dependencies: {e}")
        print_info("\nPlease install dependencies manually:")
        print_info("  1. Create/activate a virtual environment")
        print_info("  2. Run: pip install -r requirements.txt")
        print_info("  3. Run this script again")
        return False

def check_dependencies():
    """Check if required packages are installed."""
    required = ["pandas", "openpyxl", "PyInstaller"]
    missing = []
    
    for package in required:
        try:
            __import__(package.lower().replace("-", "_"))
            print_success(f"{package} is installed")
        except ImportError:
            missing.append(package)
            print_error(f"{package} is not installed")
    
    if missing:
        print_info("Installing missing dependencies...")
        return install_dependencies()
    
    return True

def clean_build_files():
    """Remove previous build artifacts."""
    print_info("Cleaning previous build files...")
    
    dirs_to_remove = ["build", "dist", "__pycache__"]
    files_to_remove = ["*.spec"]
    
    for dir_name in dirs_to_remove:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print_info(f"  Removed {dir_name}/")
    
    for pattern in files_to_remove:
        for file in Path(".").glob(pattern):
            file.unlink()
            print_info(f"  Removed {file}")
    
    print_success("Build files cleaned")

def build_executable():
    """Build the executable using PyInstaller."""
    print_info("Building executable...")
    
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--console",
        "src/main.py",
        "--name", "Excel_Handler"
    ]
    
    try:
        subprocess.check_call(cmd)
        return True
    except subprocess.CalledProcessError as e:
        print_error(f"Build failed: {e}")
        return False

def main():
    """Main build process."""
    print_header("Excel Handler - Build Script")
    
    # Check Python version
    if not check_python_version():
        sys.exit(1)
    
    # Check/install dependencies
    print_info("\nChecking dependencies...")
    if not check_dependencies():
        print_error("Failed to install dependencies")
        sys.exit(1)
    
    # Clean previous builds
    print_info("\nCleaning previous builds...")
    clean_build_files()
    
    # Build executable
    print_header("Building Executable")
    if not build_executable():
        print_error("Build failed!")
        sys.exit(1)
    
    # Success message
    print_header("Build Successful!")
    
    # Determine executable name based on OS
    if sys.platform == "win32":
        exe_name = "Excel_Handler.exe"
    else:
        exe_name = "Excel_Handler"
    
    exe_path = os.path.join("dist", exe_name)
    
    if os.path.exists(exe_path):
        print_success(f"Executable created: {exe_path}")
        print_info(f"File size: {os.path.getsize(exe_path) / (1024*1024):.2f} MB")
        
        print("\nTo run the application:")
        if sys.platform == "win32":
            print_info(f"  dist\\{exe_name}")
        else:
            print_info(f"  ./dist/{exe_name}")
    else:
        print_error(f"Executable not found at {exe_path}")
        sys.exit(1)
    
    print()

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nBuild cancelled by user.")
        sys.exit(1)
    except Exception as e:
        print_error(f"Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
