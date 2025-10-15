# Excel Handler

A user-friendly Python application for validating and processing Excel files, designed for non-technical users.

## Features

- **File Browser**: Easy file selection with graphical file picker
- **Column Validation**: Automatically validates Excel file format
- **Data Type Validation**: Validates data types (string, numeric, date)
- **Required Field Checking**: Ensures required fields are not empty
- **Data Extraction**: Extracts data into a structured object for processing
- **Error Handling**: Clear error messages for validation failures
- **User-Friendly**: Simple console interface with step-by-step prompts

## Installation

### For Non-Technical Users

1. Go to the [Releases page](../../releases)
2. Download `Excel_Handler.exe` from the latest release
3. Double-click the executable to run the application

**No installation required!** Just download and run.

## Building from Source

### Automated Build (Recommended)

This project uses GitHub Actions to automatically build Windows executables:

1. Push a version tag: `git tag v1.0.0 && git push origin v1.0.0`
2. GitHub Actions builds the Windows `.exe` automatically
3. Download from the Releases page

See `GITHUB_ACTIONS_GUIDE.md` for detailed instructions.

### Manual Build

If you want to build locally:

1. Make sure you have Python 3.7+ installed
2. Run the build script:

   ```bash
   python build.py
   ```

   Or manually:

   ```bash
   pip install -r requirements.txt
   pyinstaller --onefile --console src/main.py --name Excel_Handler
   ```

3. The executable will be in the `dist` folder

**Note:** Build on Windows to get a Windows executable, or use GitHub Actions to build automatically.

## Expected Excel Format

The application validates Excel files against the **Optix Input Template** format with these required columns:

| Column Name | Data Type | Required |
|------------|-----------|----------|
| Customer | String | Yes |
| Customer RO | String | Yes |
| Customer Pre Work PN | String | Yes |
| Customer Pre Work SN | String | Yes |
| Description | String | Yes |
| Customer Supplier Code | String | Yes |
| Priority - PDR/TDR | Numeric | Yes |
| Supplier Name | String | Yes |
| Last Quote ID | String | No (Optional) |
| RO Create Date | Date | Yes |
| RO Close Date | Date | Yes |

**Template file**: `templates/Optix_input_template.xlsx`

### Validation Rules

- **Column Names**: Must match exactly (case-sensitive)
- **Required Fields**: Cannot be empty/null
- **Data Types**:
  - String: Any text value
  - Numeric: Numbers only (integers or decimals)
  - Date: Valid date format (Excel date or text date)

## Usage

1. Run `Excel_Handler.exe`
2. Choose option 1 to browse for your Excel file (recommended for non-technical users)
3. The application will:
   - Read the Excel file
   - Validate the columns match the expected format
   - Extract data into an object if validation passes
   - Display error messages if validation fails
4. Follow the on-screen prompts

## Requirements

- Linux or Windows operating system
- Excel files must be in `.xlsx` or `.xls` format
- For building: Python 3.7+ with pandas, openpyxl, PyInstaller

## Troubleshooting

### EOFError when running executable

If you encounter an `EOFError` when running the executable, this has been fixed in the latest version. To rebuild:

```bash
python build.py
```

See `BUGFIX_EOFERROR.md` for details.

## License

This project is licensed under the MIT License.
