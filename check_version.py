
try:
    import google.generativeai as genai
    print(f"Version: {genai.__version__}")
    print(f"Location: {genai.__file__}")
except ImportError:
    print("google.generativeai is not installed.")
except AttributeError:
    # Fallback for very old versions that might not have __version__
    import google.generativeai
    print("Could not determine version. The library is likely very old.")
    print(f"Location: {google.generativeai.__file__}")
