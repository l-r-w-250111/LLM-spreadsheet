
import uno

def get_calc_state(queries):
    """
    実行中のLibreOffice Calcインスタンスに接続し、指定された複数の情報を取得する。

    Args:
        queries (dict): 取得したい情報のクエリ。
                        例: {"cell_value": "A1", "active_sheet_name": True, "sheet_count": True}

    Returns:
        dict: クエリに対する結果のキーと値のペア。
    """
    results = {}
    try:
        # UNOコンポーネントの取得
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context)
        context = resolver.resolve(
            "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        desktop = context.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", context)
        
        model = desktop.getCurrentComponent()
        if not hasattr(model, "Sheets"):
            return {"error": "アクティブなドキュメントがCalcのスプレッドシートではありません。"}

        # クエリに基づいて情報を収集
        if queries.get("cell_values"):
            for cell_address in queries["cell_values"]:
                try:
                    if "." in cell_address:
                        sheet_name, cell = cell_address.split(".", 1)
                        sheet = model.Sheets.getByName(sheet_name)
                    else:
                        sheet = model.getCurrentController().getActiveSheet()
                        cell = cell_address
                    
                    cell_obj = sheet.getCellRangeByName(cell)
                    try:
                        value = cell_obj.getString()
                        if value == "":
                            raise ValueError("Empty string, try other types")
                    except Exception:
                        try:
                            value = cell_obj.getValue()
                        except Exception:
                            value = cell_obj.getFormula()
                    results[f"cell_value_{cell_address}"] = f"セル {cell_address} の値: {value}"
                except Exception as e:
                    results[f"cell_value_{cell_address}"] = f"セル {cell_address} の値の取得に失敗: {e}"

        if queries.get("active_sheet_name"):
            try:
                sheet = model.getCurrentController().getActiveSheet()
                results["active_sheet_name"] = f"アクティブシート名: {sheet.getName()}"
            except Exception as e:
                results["active_sheet_name"] = f"アクティブシート名の取得に失敗: {e}"

        if queries.get("sheet_count"):
            try:
                count = model.Sheets.getCount()
                results["sheet_count"] = f"シートの総数: {count}"
            except Exception as e:
                results["sheet_count"] = f"シート数の取得に失敗: {e}"

        if queries.get("sheet_names"):
            try:
                names = model.Sheets.getElementNames()
                results["sheet_names"] = f"全シート名: {list(names)}"
            except Exception as e:
                results["sheet_names"] = f"全シート名の取得に失敗: {e}"

        if queries.get("document_count"):
            try:
                components = desktop.getComponents()
                # com.sun.star.sheet.SpreadsheetDocument をサポートするコンポーネントを数える
                doc_count = sum(1 for comp in components if hasattr(comp, "Sheets"))
                results["document_count"] = f"Calcドキュメントの総数: {doc_count}"
            except Exception as e:
                results["document_count"] = f"ドキュメント数の取得に失敗: {e}"

    except Exception as e:
        results["error"] = f"Calcの状態取得中にエラーが発生しました: {e}"
    
    return results
