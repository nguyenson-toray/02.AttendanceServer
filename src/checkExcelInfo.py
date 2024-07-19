import openpyxl
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives import hashes
from cryptography.exceptions import InvalidSignature
def check_excel_properties(filename):
  """
  This function checks basic properties of an Excel file using openpyxl.

  Args:
      filename (str): Path to the Excel file.

  Returns:
      dict: A dictionary containing creation and modification details
             or None if the file cannot be opened.
  """
  try:
    wb = openpyxl.load_workbook(filename=filename, read_only=True)
    properties = wb.properties
    return {
        "created": properties.created,
        "modified": properties.modified,
        "creator": properties.creator,
        "last_modified_by": properties.lastModifiedBy
    }
  except FileNotFoundError:
    print(f"Error: File '{filename}' not found.")
    return None
  except Exception as e:
    print(f"Error opening '{filename}': {e}")
    return None

# Example usage
filename =r"D:\Programming\01.AttendanceApp\02.Server\Toray's employees information All in one.xlsx"
properties = check_excel_properties(filename)

if properties:
  print("File Properties:")
  for key, value in properties.items():
    print(f"{key}: {value}")
else:
  print("Failed to retrieve properties.")

