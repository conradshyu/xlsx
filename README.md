# Excel → pandas (Pyodide / WebAssembly) → Handsontable

Reads `.xlsx` files entirely in the browser: Python + pandas run in WebAssembly via
[Pyodide](https://pyodide.org), and the resulting DataFrame is shown in
[Handsontable](https://handsontable.com). UI styled with [Semantic UI](https://semantic-ui.com).

## Run

Pyodide must be served over HTTP (not `file://`):

```bash
python3 -m http.server 8000
# open http://localhost:8000
```

## Features

- Upload any `.xlsx` (or click **Load sample.xlsx**), choose a sheet and header row
- `pandas.read_excel()` (openpyxl engine) parses the workbook in WASM
- Editable Handsontable grid with sorting, filters, dtype shown in headers
- `df.describe()` summary statistics of the (edited) grid
- **Save as .xlsx**: grid data → pandas → `to_excel()` → browser download

## Layout

- `index.html` – page (Semantic UI, Handsontable, Pyodide from CDN)
- `js/app.js` – JS ↔ Python bridge and grid rendering
- `py/excel_loader.py` – pandas code executed inside Pyodide
- `data/sample.xlsx` – sample workbook (Sales, Employees sheets)

Handsontable uses the `non-commercial-and-evaluation` license key; replace for commercial use.
