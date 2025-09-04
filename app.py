from flask import Flask, request, jsonify, send_from_directory, url_for
import pandas as pd
import os
import uuid
import json

app = Flask(__name__)

# 儲存 Excel 輸出的目錄
OUTPUT_DIR = "outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)


def export_to_excel(schema: dict, data: dict, output_path: str) -> str:
    """
    根據 schema 和 data 產生 Excel 檔案
    """
    # 找出需要展平的表格名稱 (例如 "損益表")
    table_name = None
    for key, val in schema.items():
        if isinstance(val, list):
            table_name = key
            break
    if not table_name:
        raise ValueError("❌ schema 沒有包含表格 (list) 結構")

    # 固定欄位
    fixed_cols = [k for k, v in schema.items() if not isinstance(v, list)]
    # 表格欄位
    table_cols = list(schema[table_name][0].keys())
    columns = fixed_cols + table_cols

    # 展平成 rows
    rows = []
    for item in data[table_name]:
        row = {}
        for col in fixed_cols:
            row[col] = data.get(col, "")
        for col in table_cols:
            row[col] = item.get(col, "")
        rows.append(row)

    df = pd.DataFrame(rows, columns=columns)
    df.to_excel(output_path, index=False)
    return output_path


@app.route("/export", methods=["POST"])
def export_api():
    try:
        payload = request.get_json()
        schema = payload.get("schema")
        data = payload.get("data")

        # 如果 schema/data 是字串 → 轉回 dict
        if isinstance(schema, str):
            schema = json.loads(schema)
        if isinstance(data, str):
            data = json.loads(data)

        filename = f"{uuid.uuid4().hex}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, filename)
        export_to_excel(schema, data, output_path)

        download_url = url_for("download_file", filename=filename, _external=True)
        return jsonify({"download_url": download_url})

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/download/<filename>", methods=["GET"])
def download_file(filename):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))  # Zeabur 會傳入 $PORT
    app.run(host="0.0.0.0", port=port)
