
import os
import google.generativeai as genai

# 注意: このスクリプトを実行する前に、環境変数に 'GOOGLE_API_KEY' を設定してください。
# 例: os.environ['GOOGLE_API_KEY'] = 'YOUR_API_KEY'
genai.configure(api_key=os.environ["GOOGLE_API_KEY"])

# LLMモデルの設定
GENERATION_CONFIG = {
    "temperature": 0.7,
    "top_p": 1,
    "top_k": 1,
    "max_output_tokens": 2048,
}
SAFETY_SETTINGS = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
]

# --- プロンプトテンプレート ---

GENERATOR_PROMPT_TEMPLATE = """
あなたは、ユーザーの指示をPythonのUNO (Universal Network Objects) APIを使ったLibreOffice Calc操作コードに変換するエキスパートです。

# 厳格なルール
- **絶対に**新しいドキュメントを作成してはいけません。`desktop.loadComponentFromURL`の使用は固く禁止します。
- **必ず** `desktop.getCurrentComponent()` を使って、既に開かれているドキュメントを取得してください。
- **必ず** `model.getCurrentController().getActiveSheet()` を使って、現在アクティブなシートを取得してください。

# コード生成の注意点
生成するコードは、LibreOfficeに付属のPython環境で実行可能な、単独のスクリプトにしてください。
コードの冒頭で、以下の様にしてデスクトップオブジェクトを取得する必要があります。
import uno
local_context = uno.getComponentContext()
resolver = local_context.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local_context)
context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
model = desktop.getCurrentComponent() # 既存のドキュメントを取得

生成するコードは、```python ```で囲んでください。

# 指示
{instruction}

# 過去の試行と評価 (フィードバック)
{feedback_history}

# あなたが生成するべきPythonコード:
"""

EVALUATOR_PROMPT_TEMPLATE = """
あなたは、Pythonコードの実行結果を評価する、非常に厳格なエキスパートです。
以下の「指示」「実行前の状態」「実行後の状態」「実行されたコード」を分析し、指示が達成されたかを評価してください。

評価の最重要ポイントは、「指示にない副作用(Side Effect)がないか」です。
例えば、不要なドキュメントやシートが作成されていないか、意図しないセルが変更されていないかなどを厳しくチェックしてください。

評価は以下の形式で、簡潔に日本語で記述してください。
- 成功/失敗: [成功または失敗]
- 理由: [成功または失敗した理由を、実行前後の状態変化と副作用に言及しながら具体的に記述する]
- 改善案: [失敗した場合の、具体的なコードの改善案]
- 復元コード: [失敗した場合に、状態を実行前に戻すためのPythonコードを ```python ``` で囲んで記述。復元が不要な場合は、コードブロック内に`pass`とだけ記述してください。]

# 指示
{instruction}

# 実行前のアプリの状態
{pre_run_state}

# 実行されたコード
```python
{code}
```

# 実行後のアプリの状態
{post_run_state}

# あなたの評価:
"""

def invoke_llm(prompt):
    """
    指定されたプロンプトを使用してLLMを呼び出し、応答を返す。

    Args:
        prompt (str): LLMに送信するプロンプト。

    Returns:
        str: LLMからの応答テキスト。
    """
    model = genai.GenerativeModel(
        model_name="gemini-1.5-flash",
        generation_config=GENERATION_CONFIG,
        safety_settings=SAFETY_SETTINGS,
    )
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        print(f"LLMの呼び出し中にエラーが発生しました: {e}")
        return None
