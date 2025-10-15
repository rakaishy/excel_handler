# Configuration Example for Excel Handler
# This shows the current expected format for Optix Input Template

# Define the expected columns in your Excel file
# Leave empty [] to accept any columns
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

# Example configurations:

# For a simple contact list:
# EXPECTED_COLUMNS = ['Name', 'Email', 'Phone']

# For an inventory system:
# EXPECTED_COLUMNS = ['SKU', 'Product Name', 'Quantity', 'Price', 'Category']

# For employee data:
# EXPECTED_COLUMNS = ['Employee ID', 'First Name', 'Last Name', 'Department', 'Hire Date']

# To accept any columns (no validation):
# EXPECTED_COLUMNS = []
