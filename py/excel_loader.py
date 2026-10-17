"""Excel helpers executed inside Pyodide (Python compiled to WebAssembly)."""
import io
import sys
import json
import pandas

def list_sheets(path):
    with pandas.ExcelFile(path, engine="openpyxl") as xls:
        return json.dumps(xls.sheet_names)

def read_sheet(path, sheet, header_row=0):
    """Read one sheet and return JSON {columns, data, dtypes, shape} for Handsontable."""
    header = None if header_row is None or header_row < 0 else int(header_row)
    df = pandas.read_excel(path, sheet_name=sheet, header=header, engine="openpyxl")
    df.columns = [str(c) for c in df.columns]
    payload = json.loads(df.to_json(orient="split", date_format="iso", default_handler=str))

    return json.dumps({
        "columns": payload["columns"],
        "data": payload["data"],
        "dtypes": {c: str(t) for c, t in df.dtypes.items()},
        "shape": list(df.shape),
    })

def describe(columns_json, data_json):
    df = pandas.DataFrame(json.loads(data_json), columns=json.loads(columns_json))

    for col in df.columns:
        converted = pandas.to_numeric(df[col], errors="coerce")
        if converted.notna().sum() == df[col].notna().sum():
            df[col] = converted

    numeric = df.select_dtypes(include="number")

    if numeric.empty:
        return json.dumps({"columns": [], "data": []})

    stats = numeric.describe().round(4).reset_index().rename(columns={"index": "stat"})
    payload = json.loads(stats.to_json(orient="split"))

    return json.dumps({"columns": payload["columns"], "data": payload["data"]})

def to_xlsx_bytes(columns_json, data_json, sheet_name="Sheet1"):
    df = pandas.DataFrame(json.loads(data_json), columns=json.loads(columns_json))
    buf = io.BytesIO()

    with pandas.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name[:31] or "Sheet1", index=False)

    return buf.getvalue()
