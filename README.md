# web_xlsx — Excel loader powered by WebAssembly + Handsontable

This example loads an Excel workbook (`.xlsx`, `.xlsm`, or legacy `.xls`)
selected in the browser, parses it **entirely client-side** using a
WebAssembly module compiled from Rust (via the
[`calamine`](https://github.com/tafia/calamine) crate + `wasm-bindgen`), and
renders/edits the data with [Handsontable](https://handsontable.com/).

No file is ever uploaded to a server — parsing happens in-process inside the
browser's WebAssembly runtime.

## Project layout

```
web_xlsx/
├── wasm-xlsx/                 Rust crate compiled to WebAssembly
│   ├── Cargo.toml
│   └── src/lib.rs             parse_xlsx() / parse_xls() -> JSON string
├── public/                    Static site served to the browser
│   ├── index.html
│   ├── style.css
│   ├── app.js                 wires <input type=file> -> wasm -> Handsontable
│   └── vendor/wasm-xlsx/      output of `wasm-pack build` (generated)
├── server.js                  tiny zero-dependency static file server
└── package.json
```

## How it works

1. `wasm-xlsx/src/lib.rs` exposes `parse_xlsx(bytes: &[u8]) -> String` and
   `parse_xls(bytes: &[u8]) -> String` via `wasm_bindgen`. Each reads the raw
   file bytes with `calamine`, walks every worksheet, and serializes the
   result as JSON: `{ "sheets": [{ "name": ..., "rows": [[...]] }] }`.
2. `public/app.js` loads the compiled `.wasm` module with `init()`, reads the
   user-selected file as an `ArrayBuffer`, hands the bytes to `parse_xlsx`,
   and parses the returned JSON.
3. The parsed rows are fed into a `Handsontable` instance for viewing and
   editing, with a sheet selector for multi-sheet workbooks and a "Export to
   CSV" button for saving edits back out.

## Rebuilding the WebAssembly module

Requires Rust + `wasm-pack` (already built once into `public/vendor/wasm-xlsx`):

```bash
rustup target add wasm32-unknown-unknown    # first time only
cargo install wasm-pack                     # first time only
npm run build:wasm
```

## Running the demo

```bash
npm run serve
# then open http://localhost:8080
```

(WebAssembly + ES module imports require an HTTP server — opening
`index.html` directly via `file://` will not work in most browsers.)

## Notes

* Handsontable is loaded from a CDN in `index.html` under its free
  non-commercial evaluation license key (`non-commercial-and-evaluation`).
  Swap in a commercial license key for production use.
* `parse_xlsx` handles `.xlsx`/`.xlsm`; `parse_xls` handles legacy binary
  `.xls` files. `app.js` picks the right one based on the file extension.
* Numeric cells are rendered as plain numbers (no scientific notation for
  whole numbers); dates come through using calamine's default formatting.
