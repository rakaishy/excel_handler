"""
Test script to validate the Excel Handler with the template file.
"""
import sys
sys.path.insert(0, 'src')

from main import read_and_validate_excel, EXPECTED_COLUMNS, COLUMN_VALIDATIONS

print("=" * 60)
print("Testing Excel Handler Validation")
print("=" * 60)

print("\nExpected columns:")
for i, col in enumerate(EXPECTED_COLUMNS, 1):
    print(f"  {i}. {col}")

print("\nColumn validation rules:")
for col, rules in COLUMN_VALIDATIONS.items():
    required = "Required" if rules['required'] else "Optional"
    print(f"  - {col}: {rules['type']} ({required})")

print("\n" + "=" * 60)
print("Testing with template file...")
print("=" * 60)

template_file = 'templates/Optix_input_template.xlsx'

excel_data = read_and_validate_excel(template_file, EXPECTED_COLUMNS, COLUMN_VALIDATIONS)

if excel_data:
    print("\n✓ VALIDATION PASSED!")
    excel_data.summary()
    
    print("\n" + "=" * 60)
    print("Data extraction successful!")
    print("=" * 60)
    print(f"Data object type: {type(excel_data)}")
    print(f"Access methods available:")
    print(f"  - excel_data.get_data() -> DataFrame")
    print(f"  - excel_data.to_dict() -> List of dictionaries")
    print(f"  - excel_data.summary() -> Print summary")
    
    # Show sample data
    print("\n" + "=" * 60)
    print("Sample data (first row):")
    print("=" * 60)
    data_dict = excel_data.to_dict()
    if data_dict:
        first_row = data_dict[0]
        for key, value in first_row.items():
            print(f"  {key}: {value}")
else:
    print("\n✗ VALIDATION FAILED!")
    print("The template file does not match the expected format.")
