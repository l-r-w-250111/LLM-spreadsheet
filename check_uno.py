
import sys
import os

# LibreOfficeのprogramディレクトリのパス
libreoffice_program_path = r"C:\Program Files\LibreOffice\program"

# sys.pathにLibreOfficeのprogramディレクトリを追加
if libreoffice_program_path not in sys.path:
    sys.path.append(libreoffice_program_path)

try:
    import uno
    print("Successfully imported uno module.")
except ImportError:
    print("Failed to import uno module. Please ensure LibreOffice is installed and the 'program' directory is correctly added to sys.path.")
    print(f"Current sys.path: {sys.path}")

# uno.pyの場所も確認
try:
    import uno
    uno_path = os.path.dirname(uno.__file__)
    print(f"uno.py is located at: {uno_path}")
except AttributeError:
    print("Could not determine location of uno.py (uno module might be a built-in or not a file-based module).")
except Exception as e:
    print(f"An unexpected error occurred while checking uno.py location: {e}")
