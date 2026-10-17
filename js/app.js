const XLSX_PATH = "/tmp/upload.xlsx";
const HOT_LICENSE = "non-commercial-and-evaluation";

const $ = (id) => document.getElementById(id);
let pyodide, py, hot, statsHot;
let current = { name: null, columns: [], sheet: null };

function setStatus(kind, header, text, loading = false) {
  $("status").className = `ui icon ${kind} message`;
  $("status").querySelector("i").className = loading ? "notched circle loading icon" : {
    success: "check circle icon", error: "exclamation triangle icon", info: "info circle icon",
  }[kind];
  $("status").querySelector(".header").textContent = header;
  $("status-text").textContent = text;
}

function enable(on, ids = ["file-input", "btn-sample", "btn-reload", "btn-describe", "btn-download", "sheet-select"]) {
  ids.forEach((id) => ($(id).disabled = !on));
}

async function initPython() {
  pyodide = await loadPyodide();
  setStatus("info", "Loading Python packages…", "pandas, openpyxl", true);
  await pyodide.loadPackage(["pandas", "micropip"]);
  const micropip = pyodide.pyimport("micropip");
  await micropip.install("openpyxl");

  const src = await (await fetch("py/excel_loader.py")).text();
  pyodide.FS.mkdirTree("/home/pyodide/app");
  pyodide.FS.writeFile("/home/pyodide/app/excel_loader.py", src);
  pyodide.runPython("import sys; sys.path.insert(0, '/home/pyodide/app')");
  py = pyodide.pyimport("excel_loader");

  const ver = pyodide.runPython("import pandas; pandas.__version__");
  setStatus("success", "Python is ready", `Pyodide ${pyodide.version} · pandas ${ver}. Choose an .xlsx file or load the sample.`);
  enable(true, ["file-input", "btn-sample"]);
}

async function openWorkbook(bytes, name) {
  try {
    setStatus("info", "Reading workbook…", name, true);
    pyodide.FS.writeFile(XLSX_PATH, new Uint8Array(bytes));
    const sheets = JSON.parse(py.list_sheets(XLSX_PATH));
    $("sheet-select").innerHTML = sheets.map((s) => `<option>${escapeHtml(s)}</option>`).join("");
    current.name = name;
    enable(true);
    await loadSheet();
  } catch (err) {
    showError(err);
  }
}

async function loadSheet() {
  try {
    const sheet = $("sheet-select").value;
    const headerRow = parseInt($("header-row").value, 10);
    const t0 = performance.now();
    const res = JSON.parse(py.read_sheet(XLSX_PATH, sheet, Number.isNaN(headerRow) ? 0 : headerRow));
    const ms = (performance.now() - t0).toFixed(0);
    current.columns = res.columns;
    current.sheet = sheet;
    renderGrid(res);
    $("stats-segment").style.display = "none";
    $("data-info").textContent = `${current.name} › ${sheet} — ${res.shape[0]} rows × ${res.shape[1]} columns`;
    setStatus("success", "Loaded with pandas.read_excel()", `Parsed in ${ms} ms inside WebAssembly.`);
  } catch (err) {
    showError(err);
  }
}

function columnType(dtype) {
  if (/int|float/.test(dtype)) return { type: "numeric" };
  if (/bool/.test(dtype)) return { type: "checkbox" };
  return { type: "text" };
}

function renderGrid({ columns, data, dtypes }) {
  if (hot) hot.destroy();
  hot = new Handsontable($("grid"), {
    data,
    colHeaders: columns.map((c) => `${c}<br><small>${dtypes[c]}</small>`),
    columns: columns.map((c) => columnType(dtypes[c])),
    rowHeaders: true,
    height: 480,
    width: "100%",
    stretchH: "all",
    filters: true,
    dropdownMenu: true,
    columnSorting: true,
    manualColumnResize: true,
    contextMenu: true,
    licenseKey: HOT_LICENSE,
  });
}

function currentData() {
  return JSON.stringify(hot.getSourceData().map((row) => (Array.isArray(row) ? row : Object.values(row))));
}

function describe() {
  try {
    const res = JSON.parse(py.describe(JSON.stringify(current.columns), currentData()));
    $("stats-segment").style.display = "";
    if (statsHot) statsHot.destroy();
    if (!res.columns.length) {
      $("stats-grid").innerHTML = '<div class="ui warning message">No numeric columns.</div>';
      statsHot = null;
      return;
    }
    $("stats-grid").innerHTML = "";
    statsHot = new Handsontable($("stats-grid"), {
      data: res.data, colHeaders: res.columns, readOnly: true, stretchH: "all",
      width: "100%", height: "auto", licenseKey: HOT_LICENSE,
    });
  } catch (err) {
    showError(err);
  }
}

function download() {
  try {
    const proxy = py.to_xlsx_bytes(JSON.stringify(current.columns), currentData(), current.sheet || "Sheet1");
    const bytes = proxy.toJs();
    proxy.destroy();
    const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const a = Object.assign(document.createElement("a"), {
      href: URL.createObjectURL(blob),
      download: (current.name || "data.xlsx").replace(/\.xls[xm]?$/i, "") + "-edited.xlsx",
    });
    a.click();
    URL.revokeObjectURL(a.href);
  } catch (err) {
    showError(err);
  }
}

function showError(err) {
  console.error(err);
  setStatus("error", "Error", String(err.message || err).split("\n").slice(-2).join(" "));
}

function escapeHtml(s) {
  return s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

$("file-input").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (file) await openWorkbook(await file.arrayBuffer(), file.name);
});
$("btn-sample").addEventListener("click", async () => {
  const resp = await fetch("data/sample.xlsx");
  await openWorkbook(await resp.arrayBuffer(), "sample.xlsx");
});
$("sheet-select").addEventListener("change", loadSheet);
$("btn-reload").addEventListener("click", loadSheet);
$("btn-describe").addEventListener("click", describe);
$("btn-download").addEventListener("click", download);

initPython().catch(showError);
