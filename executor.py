
def execute_code(code_string):
    """
    文字列として渡されたPythonコードを実行する。

    Args:
        code_string (str): 実行するPythonコード。

    Returns:
        str: 実行中に発生したエラーメッセージ。成功した場合はNone。
    """
    try:
        # セキュリティのため、組み込み関数やグローバル変数を制限する
        # この例では簡単化のため、制限は緩やかにしています
        # 実際のアプリケーションでは、より厳格な制限が必要です
        exec(code_string, {"__builtins__": __builtins__}, {})
        return None
    except Exception as e:
        print(f"コードの実行中にエラーが発生しました: {e}")
        return str(e)
