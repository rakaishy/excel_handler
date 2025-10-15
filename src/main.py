import os
import sys
import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

# Expected column format for Optix PDR TDR - Digitization Priority
EXPECTED_COLUMNS = [
    'Customer',
    'Customer RO',
    'Customer Pre Work PN',
    'Customer Pre Work SN',
    'Description',
    'Customer Supplier Code',
    'Priority - PDR/TDR',
    'Supplier Name',
    'Last Quote ID',
    'RO Create Date',
    'RO Close Date',
]

# Column data type expectations and validation rules
COLUMN_VALIDATIONS = {
    'Customer': {'type': 'string', 'required': True},
    'Customer RO': {'type': 'string', 'required': True},
    'Customer Pre Work PN': {'type': 'string', 'required': True},
    'Customer Pre Work SN': {'type': 'string', 'required': True},
    'Description': {'type': 'string', 'required': True},
    'Customer Supplier Code': {'type': 'string', 'required': True},
    'Priority - PDR/TDR': {'type': 'numeric', 'required': True},
    'Supplier Name': {'type': 'string', 'required': True},
    'Last Quote ID': {'type': 'string', 'required': False},  # Can be empty
    'RO Create Date': {'type': 'date', 'required': True},
    'RO Close Date': {'type': 'date', 'required': True},
}

class ExcelData:
    """Class to hold the extracted Excel data."""
    def __init__(self, dataframe, file_path):
        self.dataframe = dataframe
        self.file_path = file_path
        self.rows = len(dataframe)
        self.columns = list(dataframe.columns)
        
    def __repr__(self):
        return f"ExcelData(rows={self.rows}, columns={self.columns})"
    
    def get_data(self):
        """Return the dataframe."""
        return self.dataframe
    
    def to_dict(self):
        """Convert dataframe to dictionary."""
        return self.dataframe.to_dict('records')
    
    def summary(self):
        """Print a summary of the data."""
        print(f"\n=== Data Summary ===")
        print(f"File: {self.file_path}")
        print(f"Total rows: {self.rows}")
        print(f"Columns: {', '.join(self.columns)}")
        print(f"\nFirst 5 rows:")
        print(self.dataframe.head())

def select_file_dialog():
    """Open a file dialog to select an Excel file."""
    try:
        # Create a hidden root window
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        
        # Open file dialog
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        
        root.destroy()
        return file_path
    except Exception as e:
        print(f"Error opening file dialog: {str(e)}")
        return None

def validate_columns(df, expected_columns):
    """
    Validate that the dataframe has the expected columns.
    
    Args:
        df: pandas DataFrame
        expected_columns: list of expected column names
        
    Returns:
        tuple: (is_valid, missing_columns, extra_columns)
    """
    if not expected_columns:
        # If no expected columns defined, accept any columns
        return True, [], []
    
    actual_columns = set(df.columns)
    expected_set = set(expected_columns)
    
    missing_columns = expected_set - actual_columns
    extra_columns = actual_columns - expected_set
    
    is_valid = len(missing_columns) == 0
    
    return is_valid, list(missing_columns), list(extra_columns)

def validate_data_types(df, column_validations):
    """
    Validate data types and required fields in the dataframe.
    
    Args:
        df: pandas DataFrame
        column_validations: dict of column validation rules
        
    Returns:
        tuple: (is_valid, errors_list)
    """
    errors = []
    
    for col_name, rules in column_validations.items():
        if col_name not in df.columns:
            continue
            
        col_data = df[col_name]
        expected_type = rules.get('type', 'string')
        is_required = rules.get('required', True)
        
        # Check for required fields
        if is_required:
            null_count = col_data.isna().sum()
            if null_count > 0:
                errors.append(f"Column '{col_name}' has {null_count} empty/null values (required field)")
        
        # Check data types (only for non-null values)
        non_null_data = col_data.dropna()
        if len(non_null_data) > 0:
            if expected_type == 'numeric':
                # Check if values can be converted to numeric
                try:
                    pd.to_numeric(non_null_data, errors='coerce')
                    non_numeric = pd.to_numeric(non_null_data, errors='coerce').isna().sum()
                    if non_numeric > 0:
                        errors.append(f"Column '{col_name}' has {non_numeric} non-numeric values")
                except:
                    errors.append(f"Column '{col_name}' should contain numeric values")
                    
            elif expected_type == 'date':
                # Check if values can be converted to datetime
                try:
                    pd.to_datetime(non_null_data, errors='coerce')
                    non_date = pd.to_datetime(non_null_data, errors='coerce').isna().sum()
                    if non_date > 0:
                        errors.append(f"Column '{col_name}' has {non_date} invalid date values")
                except:
                    errors.append(f"Column '{col_name}' should contain date values")
    
    is_valid = len(errors) == 0
    return is_valid, errors

def read_and_validate_excel(file_path, expected_columns=None, column_validations=None):
    """
    Read and validate an Excel file.
    
    Args:
        file_path: path to the Excel file
        expected_columns: list of expected column names (optional)
        column_validations: dict of column validation rules (optional)
        
    Returns:
        ExcelData object if successful, None if validation fails
    """
    try:
        # Read the Excel file
        print(f"\nReading file: {file_path}")
        df = pd.read_excel(file_path)
        
        print(f"✓ File loaded successfully!")
        print(f"  Found {len(df)} rows and {len(df.columns)} columns")
        print(f"  Columns: {', '.join(df.columns.astype(str))}")
        
        # Validate columns if expected columns are provided
        if expected_columns:
            is_valid, missing, extra = validate_columns(df, expected_columns)
            
            if not is_valid:
                print("\n✗ ERROR: Column validation failed!")
                print(f"  Expected columns: {', '.join(expected_columns)}")
                print(f"  Missing columns: {', '.join(missing) if missing else 'None'}")
                if extra:
                    print(f"  Extra columns found: {', '.join(extra)}")
                return None
            else:
                print("✓ Column validation passed!")
                if extra:
                    print(f"  Note: Extra columns found: {', '.join(extra)}")
        
        # Validate data types and format if validation rules are provided
        if column_validations:
            print("\nValidating data types and format...")
            is_valid, errors = validate_data_types(df, column_validations)
            
            if not is_valid:
                print("\n✗ ERROR: Data validation failed!")
                for error in errors:
                    print(f"  - {error}")
                return None
            else:
                print("✓ Data type validation passed!")
        
        # Create and return ExcelData object
        excel_data = ExcelData(df, file_path)
        return excel_data
        
    except FileNotFoundError:
        print(f"\n✗ ERROR: File not found: {file_path}")
        return None
    except Exception as e:
        print(f"\n✗ ERROR: Failed to read Excel file")
        print(f"  Details: {str(e)}")
        return None

def main():
    """Main function to run the Excel Handler application."""
    print("=" * 60)
    print("          EXCEL HANDLER - Data Validation Tool")
    print("=" * 60)
    print("\nWelcome! This tool will help you validate and process Excel files.\n")
    
    # Show expected format if defined
    if EXPECTED_COLUMNS:
        print("Expected Excel format:")
        print(f"  Required columns: {', '.join(EXPECTED_COLUMNS)}")
        print()
    
    while True:
        try:
            print("\nHow would you like to select your Excel file?")
            print("1. Browse for file (recommended)")
            print("2. Enter file path manually")
            print("3. Exit")
            
            try:
                choice = input("\nSelect option (1-3): ").strip()
            except EOFError:
                print("\n\nInput stream closed. Exiting...")
                break
            
            if choice == '3':
                print("\nThank you for using Excel Handler. Goodbye!")
                break
            
            file_path = None
            
            if choice == '1':
                print("\nOpening file browser...")
                file_path = select_file_dialog()
                if not file_path:
                    print("No file selected.")
                    continue
            elif choice == '2':
                try:
                    file_path = input("\nEnter the full path to your Excel file: ").strip('"').strip("'")
                except EOFError:
                    print("\n\nInput stream closed. Exiting...")
                    break
            else:
                print("Invalid option. Please try again.")
                continue
            
            # Validate file exists
            if not os.path.exists(file_path):
                print(f"\n✗ ERROR: File not found: {file_path}")
                continue
            
            # Validate file extension
            if not (file_path.lower().endswith('.xlsx') or file_path.lower().endswith('.xls')):
                print("\n✗ ERROR: Please provide a valid Excel file (.xlsx or .xls)")
                continue
            
            # Read and validate the Excel file
            excel_data = read_and_validate_excel(file_path, EXPECTED_COLUMNS, COLUMN_VALIDATIONS)
            
            if excel_data:
                print("\n" + "=" * 60)
                print("SUCCESS! Data extracted and validated.")
                print("=" * 60)
                excel_data.summary()
                
                # Here you can process the data further
                # Example: data_dict = excel_data.to_dict()
                # Example: df = excel_data.get_data()
                
                print("\n✓ Data is ready for processing!")
                
                # Ask if user wants to process another file
                try:
                    another = input("\nProcess another file? (y/n): ").strip().lower()
                    if another != 'y':
                        print("\nThank you for using Excel Handler. Goodbye!")
                        break
                except EOFError:
                    print("\n\nInput stream closed. Exiting...")
                    break
            else:
                print("\nPlease correct the errors and try again.")
                try:
                    retry = input("Try again? (y/n): ").strip().lower()
                    if retry != 'y':
                        break
                except EOFError:
                    print("\n\nInput stream closed. Exiting...")
                    break
                    
        except KeyboardInterrupt:
            print("\n\nOperation cancelled by user.")
            break
        except EOFError:
            print("\n\nInput stream closed. Exiting...")
            break
        except Exception as e:
            print(f"\n✗ An unexpected error occurred: {str(e)}")
            import traceback
            traceback.print_exc()
            try:
                retry = input("Try again? (y/n): ").strip().lower()
                if retry != 'y':
                    break
            except (EOFError, KeyboardInterrupt):
                break
    
    try:
        input("\nPress Enter to exit...")
    except (EOFError, KeyboardInterrupt):
        pass

if __name__ == "__main__":
    main()
