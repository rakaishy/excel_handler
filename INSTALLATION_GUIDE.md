# Excel Handler - Installation Guide for Non-Technical Users

## Simple Installation (Recommended)

### Step 1: Install Python

**Option A: Microsoft Store (Easiest)**
1. Open Microsoft Store on Windows
2. Search for "Python"
3. Click "Python 3.12" (or latest version)
4. Click "Get" or "Install"
5. Wait for installation to complete

**Option B: Python.org**
1. Go to https://www.python.org/downloads/
2. Click the yellow "Download Python" button
3. Run the installer
4. ⚠️ **IMPORTANT**: Check "Add Python to PATH"
5. Click "Install Now"

### Step 2: Extract Excel Handler

1. You should have received `excel_handler.zip`
2. Right-click the ZIP file
3. Click "Extract All..."
4. Choose a folder (e.g., `C:\ExcelHandler`)
5. Click "Extract"

### Step 3: Run the Application

1. Open the extracted folder
2. **Double-click `RUN_EXCEL_HANDLER.bat`**
3. The first time, it will install dependencies (takes 1-2 minutes)
4. After that, it starts immediately!

## That's It!

From now on, just double-click `RUN_EXCEL_HANDLER.bat` to run the application.

## Troubleshooting

### "Python is not installed"

Python is not installed or not in PATH. Try:
1. Reinstall Python from Microsoft Store (easiest)
2. OR reinstall from python.org and check "Add Python to PATH"

### "Failed to install dependencies"

Your network might block pip. Try:
1. Open Command Prompt as Administrator
2. Navigate to the folder: `cd C:\ExcelHandler`
3. Run: `python -m pip install -r requirements.txt`

### Antivirus Blocking

If your antivirus blocks the batch file:
1. Open Command Prompt
2. Navigate to the folder: `cd C:\ExcelHandler`
3. Run: `python src\main.py`

## Why This Approach?

✅ **No SmartScreen warnings** - Python scripts are not blocked
✅ **No security policy issues** - Source code is visible
✅ **Easy to update** - Just download new version
✅ **Transparent** - Users can see the code
✅ **Works everywhere** - No platform restrictions

## For IT Departments

This application:
- Uses only standard Python libraries (pandas, openpyxl)
- No network requests (runs offline)
- No system modifications
- Source code is fully visible
- No elevated privileges required
