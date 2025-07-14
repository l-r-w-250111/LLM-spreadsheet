import sys

print(f"Python Version: {sys.version}")
print("\n--- Python Executable ---")
print(sys.executable)

print("\n--- sys.path (Module Search Paths) ---")
for path in sys.path:
    print(path)

print("\n--- Attempting to import google.generativeai ---")
try:
    # google.generativeai の場所を特定
    import google.generativeai as genai
    print(f"Successfully imported google.generativeai")
    print(f"Location: {genai.__file__}")
    # バージョン情報を表示
    if hasattr(genai, '__version__'):
        print(f"Version: {genai.__version__}")
    else:
        print("Version: __version__ attribute not found.")

except ImportError as e:
    print(f"ImportError: {e}")
except AttributeError as e:
    print(f"AttributeError: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
