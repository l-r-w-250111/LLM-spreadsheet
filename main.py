import re
import sys
from llm_wrapper import invoke_llm, GENERATOR_PROMPT_TEMPLATE, EVALUATOR_PROMPT_TEMPLATE
from executor import execute_code
from state_extractor import get_calc_state
from libreoffice_manager import check_libreoffice_connection, stop_libreoffice

def extract_python_code(response_text, block_name=""):
    """
    LLMの応答から、指定された名前のPythonコードブロックを抽出する。
    block_nameが空の場合、最初のコードブロックを返す。
    """
    if block_name:
        pattern = re.compile(f"{re.escape(block_name)}.*?```python\n(.*?)\n```", re.DOTALL)
    else:
        pattern = re.compile(r"```python\n(.*?)\n```", re.DOTALL)
    
    match = pattern.search(response_text)
    if match:
        return match.group(1).strip()
    return ""

def extract_cell_references(instruction):
    """
    指示からセル参照（例: A1, Sheet1.B2）を抽出する。
    """
    # セル参照の正規表現パターン
    # 例: A1, B2, Sheet1.C3, 'Sheet Name'.D4
    pattern = re.compile(r"(?:[A-Za-z_][A-Za-z0-9_\s]*\.)?[A-Z]+\d+")
    matches = pattern.findall(instruction)
    # 重複を排除し、リストとして返す
    return list(set(matches))

def is_successful(evaluation_result):
    """
    評価結果の文字列を厳密に解釈し、成功したかどうかを判定する。
    最初に見つかった「- 成功/失敗:」行に「成功」が含まれ、かつ「失敗」が含まれていない場合のみTrueを返す。
    """
    for line in evaluation_result.splitlines():
        clean_line = line.strip()
        match = re.search(r"^- 成功/失敗:\s*\[?(成功|失敗)\]?", clean_line)
        if match:
            status = match.group(1)
            return status == "成功"
    # 「- 成功/失敗:」の行が見つからなかった場合は失敗とみなす
    return False

def main():
    """
    メインの自己改善ループを実行する。
    """
    # --- 初期設定 ---
    instruction = input("どのような操作をしますか？（例: A1セルに'Hello'と入力）: ")
    if not instruction.strip():
        print("指示が入力されなかったため、処理を終了します。")
        return

    # 指示からセル参照を抽出
    target_cells = extract_cell_references(instruction)
    if not target_cells:
        print("指示から変更対象のセルを特定できませんでした。A1セルをデフォルトとして追跡します。")
        target_cells = ["A1"]

    state_queries = {
        "cell_values": target_cells,
        "sheet_count": True,
        "active_sheet_name": True,
        "document_count": True,
    }
    max_iterations = 3
    current_iteration = 0

    feedback_history = "なし"
    final_code = ""

    print(f"--- 初期指示 ---\n{instruction}\n")

    # LibreOfficeの接続確認
    if not check_libreoffice_connection():
        return

    lo_process = None # process object is no longer used for management
    try:
        while current_iteration < max_iterations:
            current_iteration += 1
            print(f"--- イテレーション {current_iteration}/{max_iterations} ---")

            # 1. 実行前の状態確認
            print("1. 実行前の状態を確認中...")
            pre_run_state_dict = get_calc_state(state_queries)
            pre_run_state_str = "\n".join([f"{k}: {v}" for k, v in pre_run_state_dict.items()])
            print(f"実行前の状態:\n---\n{pre_run_state_str}\n---")

            # 2. コード生成
            print("2. コードを生成中...")
            generator_prompt = GENERATOR_PROMPT_TEMPLATE.format(
                instruction=instruction,
                feedback_history=feedback_history
            )
            generated_text = invoke_llm(generator_prompt)
            if not generated_text:
                print("コード生成に失敗しました。処理を中断します。")
                break

            code_to_execute = extract_python_code(generated_text)
            if not code_to_execute:
                print("応答からPythonコードを抽出できませんでした。")
                feedback_history += f"\n試行{current_iteration}: コードブロックが生成されませんでした。\n"
                continue

            print(f"生成されたコード:\n---\n{code_to_execute}\n---")

            # 3. コード実行
            print("3. コードを実行中...")
            execution_error = execute_code(code_to_execute)

            # 4. 実行後の状態確認
            print("4. 実行後の状態を確認中...")
            post_run_state_dict = get_calc_state(state_queries)
            post_run_state_str = "\n".join([f"{k}: {v}" for k, v in post_run_state_dict.items()])
            print(f"実行後の状態:\n---\n{post_run_state_str}\n---")

            # 5. 評価
            print("5. 実行結果を評価中...")
            evaluator_prompt = EVALUATOR_PROMPT_TEMPLATE.format(
                instruction=instruction,
                pre_run_state=pre_run_state_str,
                code=code_to_execute,
                post_run_state=post_run_state_str + (f"\n実行時エラー: {execution_error}" if execution_error else ""),
                target_cells=", ".join(target_cells)
            )
            evaluation_result = invoke_llm(evaluator_prompt)
            if not evaluation_result:
                print("評価の生成に失敗しました。")
                feedback_history += f"\n試行{current_iteration}の評価取得に失敗しました。\n"
                continue

            print(f"評価結果:\n---\n{evaluation_result}\n---")

            # デバッグ用: is_successfulの戻り値を確認
            success_check = is_successful(evaluation_result)
            print(f"is_successful(evaluation_result) の戻り値: {success_check}")

            # 6. フィードバックの更新と成功判定
            feedback_history += f"\n# 試行 {current_iteration}:\n実行前状態:\n{pre_run_state_str}\nコード:\n{code_to_execute}\n実行後状態:\n{post_run_state_str}\n評価:\n{evaluation_result}\n"

            if success_check:
                print("\n--- タスク成功！ ---")
                final_code = code_to_execute
                break
            else:
                print("\n--- 失敗。状態を復元します。 ---")
                revert_code = extract_python_code(evaluation_result, "- 復元コード:")
                if revert_code and revert_code != "pass":
                    print(f"復元コードを実行中...\n---\n{revert_code}\n---")
                    try:
                        revert_error = execute_code(revert_code)
                        if revert_error:
                            print(f"状態の復元中にエラーが発生しました: {revert_error}")
                            print("処理を中断します。")
                            break
                    except Exception as e:
                        print(f"復元コードの実行中に予期せぬエラーが発生しました: {e}")
                        print("評価LLMが不正なコードを生成した可能性があるため、復元をスキップします。")
                else:
                    print("復元コードが不要、または見つかりませんでした。")

            # 7. リトライ延長の確認
            if current_iteration == max_iterations:
                print("\n--- 最大試行回数に達しました ---")
                try:
                    if sys.stdout.isatty():
                        user_input = input("さらに3回試行しますか？ (y/n): ")
                        if user_input.lower() == 'y':
                            max_iterations += 3
                        else:
                            print("処理を終了します。")
                            break
                    else:
                        print("非対話環境のため、処理を終了します。")
                        break
                except (EOFError, KeyboardInterrupt):
                    print("\n処理が中断されました。終了します。")
                    break

    finally:
        stop_libreoffice(lo_process)

    print("\n--- 処理完了 ---")
    if final_code:
        print(f"最終的に成功したコード:\n{final_code}")
    else:
        print("タスクは指定された試行回数内に成功しませんでした。")

if __name__ == "__main__":
    main()
