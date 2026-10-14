// app.js — glue code between the file input, the WebAssembly xlsx parser,
// and the Handsontable grid.

import init, { parse_xlsx, parse_xls } from "./vendor/wasm-xlsx/wasm_xlsx.js";

const fileInput = document.getElementById("fileInput");
const sheetSelect = document.getElementById("sheetSelect");
const exportBtn = document.getElementById("exportBtn");
const statusEl = document.getElementById("status");
const hotContainer = document.getElementById("hotContainer");

let hotInstance = null;
let workbookSheets = []; // [{ name, rows }]
let wasmReady = false;

function setStatus(message, kind = "") {
    statusEl.textContent = message;
    statusEl.className = `status ${kind}`.trim();
}

async function bootWasm() {
    setStatus("Loading WebAssembly module…");
    await init(); // fetches & instantiates wasm_xlsx_bg.wasm
    wasmReady = true;
    setStatus("Ready. Choose an Excel file to begin.");
}

/** Turns a ragged 2D array into a rectangular one Handsontable can render. */
function normalizeRows(rows) {
    const maxCols = rows.reduce((max, row) => Math.max(max, row.length), 0);
    return rows.map((row) => {
        const padded = row.slice();
        while (padded.length < maxCols) padded.push("");
        return padded;
    });
}

function renderSheet(index) {
    const sheet = workbookSheets[index];
    if (!sheet) return;

    const data = normalizeRows(sheet.rows).length
        ? normalizeRows(sheet.rows)
        : [[""]];

    if (hotInstance) {
        hotInstance.destroy();
    }

    hotInstance = new Handsontable(hotContainer, {
        data,
        rowHeaders: true,
        colHeaders: true,
        width: "100%",
        height: 520,
        licenseKey: "non-commercial-and-evaluation",
        contextMenu: true,
        manualColumnResize: true,
        manualRowResize: true,
        filters: true,
        dropdownMenu: true,
        stretchH: "all",
    });
}

function populateSheetSelector() {
    sheetSelect.innerHTML = "";
    workbookSheets.forEach((sheet, i) => {
        const opt = document.createElement("option");
        opt.value = String(i);
        opt.textContent = `${sheet.name} (${sheet.rows.length} rows)`;
        sheetSelect.appendChild(opt);
    });
    sheetSelect.disabled = workbookSheets.length === 0;
    exportBtn.disabled = workbookSheets.length === 0;
}

async function handleFile(file) {
    if (!wasmReady) {
        setStatus("WebAssembly module is still loading, please wait…", "error");
        return;
    }

    setStatus(`Parsing "${file.name}" in WebAssembly…`);

    try {
        const buffer = await file.arrayBuffer();
        const bytes = new Uint8Array(buffer);

        const isLegacyXls = /\.xls$/i.test(file.name) && !/\.xlsx$|\.xlsm$/i.test(file.name);
        const jsonStr = isLegacyXls ? parse_xls(bytes) : parse_xlsx(bytes);
        const workbook = JSON.parse(jsonStr);

        workbookSheets = workbook.sheets;
        populateSheetSelector();
        renderSheet(0);

        const totalRows = workbookSheets.reduce((sum, s) => sum + s.rows.length, 0);
        setStatus(
            `Loaded "${file.name}" — ${workbookSheets.length} sheet(s), ${totalRows} row(s) total.`,
            "ok"
        );
    } catch (err) {
        console.error(err);
        setStatus(`Failed to parse file: ${err.message ?? err}`, "error");
    }
}

function exportCurrentSheetToCsv() {
    if (!hotInstance) return;
    const data = hotInstance.getData();
    const csv = data
        .map((row) =>
            row
                .map((cell) => {
                    const value = cell === null || cell === undefined ? "" : String(cell);
                    return /[",\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
                })
                .join(",")
        )
        .join("\n");

    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const sheetName = workbookSheets[sheetSelect.value]?.name ?? "sheet";
    a.href = url;
    a.download = `${sheetName}.csv`;
    a.click();
    URL.revokeObjectURL(url);
}

fileInput.addEventListener("change", (e) => {
    const file = e.target.files?.[0];
    if (file) handleFile(file);
});

sheetSelect.addEventListener("change", () => {
    renderSheet(Number(sheetSelect.value));
});

exportBtn.addEventListener("click", exportCurrentSheetToCsv);

bootWasm();
