import ollama

try:
    # Ollamaサーバーのデフォルトアドレスに接続を試みる
    response = ollama.chat(model='gemma3:latest', messages=[{'role': 'user', 'content': 'Hello'}])
    print("Successfully connected to Ollama and received a response.")
    print(f"Response: {response['message']['content']}")
except Exception as e:
    print(f"Failed to connect to Ollama or get a response: {e}")
    print("Please ensure Ollama server is running and 'gemma3:latest' model is downloaded.")
