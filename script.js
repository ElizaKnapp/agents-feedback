// --- State ---
let allRows = [];
let filteredRows = [];
let identifierCols = [];
let sheetNames = [];
let currentSheet = null;

let sortCol = null;
let sortAsc = true;

const el = (id) => document.getElementById(id);

const fileInput = el("fileInput");
const sheetSelect = el("sheetSelect");
const filtersDiv = el("filters");
const clearFiltersBtn = el("clearFilters");
const rowList = el("rowList");
const chat = el("chat");
const meta = el("meta");
const sortSelect = el("sortSelect");
const sortDirBtn = el("sortDirBtn");

// --- Helpers ---
function isBlank(v) {
  return v === null || v === undefined || (typeof v === "string" && v.trim() === "");
}

function detectIdentifierCols(headers) {
  // Treat anything not Qk/Rk as "identifier"
  return headers.filter((h) => !/^Q\d+$/i.test(h) && !/^R\d+$/i.test(h));
}

function rowToMessages(row) {
  // Dynamically consume Q1/R1, Q2/R2, ...
  const msgs = [];
  let k = 1;
  while (true) {
    const q = row[`Q${k}`];
    const r = row[`R${k}`];

    // stop if both missing/blank AND there is no higher pair
    if (isBlank(q) && isBlank(r)) {
      // lookahead: if next pair also blank, break; otherwise skip gap
      const qn = row[`Q${k + 1}`];
      const rn = row[`R${k + 1}`];
      if (isBlank(qn) && isBlank(rn)) break;
      k++;
      continue;
    }

    if (!isBlank(q)) msgs.push({ role: "user", text: String(q) });
    if (!isBlank(r)) msgs.push({ role: "assistant", text: String(r) });
    k++;
  }
  return msgs;
}

function getRowTitle(row) {
  // pick a nice display title
  if ("Number" in row) return `#${row["Number"]}`;
  if ("ID" in row) return `ID ${row["ID"]}`;
  return "Row";
}

function stableValue(v) {
  if (v === null || v === undefined) return "";
  return String(v);
}

// --- Filtering UI ---
let activeFilters = {}; // { colName: Set(values) }

function buildFilters() {
  filtersDiv.innerHTML = "";
  activeFilters = {};

  identifierCols.forEach((col) => {
    // get unique values
    const vals = Array.from(
      new Set(allRows.map((r) => stableValue(r[col])).filter((v) => v !== ""))
    ).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

    if (vals.length === 0) return;

    const wrap = document.createElement("div");
    wrap.className = "filter";

    const label = document.createElement("label");
    label.textContent = col;

    const select = document.createElement("select");
    select.multiple = true;
    select.size = Math.min(8, Math.max(3, vals.length));

    vals.forEach((v) => {
      const opt = document.createElement("option");
      opt.value = v;
      opt.textContent = v;
      opt.selected = true; // default all selected
      select.appendChild(opt);
    });

    activeFilters[col] = new Set(vals);

    select.addEventListener("change", () => {
      const chosen = Array.from(select.selectedOptions).map((o) => o.value);
      activeFilters[col] = new Set(chosen);
      applyAll();
    });

    wrap.appendChild(label);
    wrap.appendChild(select);
    filtersDiv.appendChild(wrap);
  });

  clearFiltersBtn.disabled = identifierCols.length === 0;
  clearFiltersBtn.onclick = () => {
    // reset: select all in every filter
    filtersDiv.querySelectorAll("select").forEach((sel) => {
      Array.from(sel.options).forEach((o) => (o.selected = true));
    });
    identifierCols.forEach((c) => {
      const vals = Array.from(
        new Set(allRows.map((r) => stableValue(r[c])).filter((v) => v !== ""))
      );
      activeFilters[c] = new Set(vals);
    });
    applyAll();
  };
}

function applyFilters(rows) {
  return rows.filter((r) => {
    for (const col of identifierCols) {
      const v = stableValue(r[col]);
      const allowed = activeFilters[col];
      if (allowed && allowed.size > 0) {
        // if column has a value, it must be in allowed; if blank, allow
        if (v !== "" && !allowed.has(v)) return false;
      }
    }
    return true;
  });
}

function applySort(rows) {
  if (!sortCol) return rows;
  const out = [...rows];
  out.sort((a, b) => {
    const av = stableValue(a[sortCol]);
    const bv = stableValue(b[sortCol]);
    const cmp = av.localeCompare(bv, undefined, { numeric: true });
    return sortAsc ? cmp : -cmp;
  });
  return out;
}

// --- Row list + chat rendering ---
let activeIndex = null;

function renderRowList(rows) {
  rowList.innerHTML = "";
  if (rows.length === 0) {
    rowList.innerHTML = `<div class="hint">No rows match filters.</div>`;
    return;
  }

  rows.forEach((row, i) => {
    const item = document.createElement("div");
    item.className = "rowitem" + (i === activeIndex ? " active" : "");

    const title = document.createElement("div");
    title.className = "rowtitle";

    const app = row["App"] ? ` — ${row["App"]}` : "";
    title.textContent = `${getRowTitle(row)}${app}`;

    const sub = document.createElement("div");
    sub.className = "rowsub";
    const p = row["Priority"] !== undefined ? `Priority: ${row["Priority"]}` : "";
    const exp = row["Expected"] !== undefined ? `Expected: ${row["Expected"]}` : "";
    sub.textContent = [p, exp].filter(Boolean).join(" • ");

    item.appendChild(title);
    item.appendChild(sub);

    item.onclick = () => {
      activeIndex = i;
      renderRowList(rows);
      renderChat(row);
    };

    rowList.appendChild(item);
  });

  // autoselect first row if none active
  if (activeIndex === null && rows.length > 0) {
    activeIndex = 0;
    renderRowList(rows);
    renderChat(rows[0]);
  }
}

function renderChat(row) {
  chat.innerHTML = "";
  meta.innerHTML = "";

  // badges for identifier columns
  identifierCols.forEach((c) => {
    const v = stableValue(row[c]);
    if (v === "") return;
    const b = document.createElement("span");
    b.className = "badge";
    b.textContent = `${c}: ${v}`;
    meta.appendChild(b);
  });

  const messages = rowToMessages(row);
  if (messages.length === 0) {
    chat.innerHTML = `<div class="hint">No Q/R pairs found for this row.</div>`;
    return;
  }

  messages.forEach((m) => {
    const bubble = document.createElement("div");
    bubble.className = `bubble ${m.role}`;

    const role = document.createElement("div");
    role.className = "role";
    role.textContent = m.role === "user" ? "User" : "Assistant";

    const text = document.createElement("div");
    text.textContent = m.text;

    bubble.appendChild(role);
    bubble.appendChild(text);
    chat.appendChild(bubble);
  });

  // scroll to bottom like chatgpt
  chat.scrollTop = chat.scrollHeight;
}

// --- Workbook loading ---
fileInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type: "array" });

  sheetNames = wb.SheetNames;
  currentSheet = sheetNames[0];

  // Populate sheet dropdown
  sheetSelect.innerHTML = "";
  sheetNames.forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    sheetSelect.appendChild(opt);
  });
  sheetSelect.disabled = false;

  sheetSelect.onchange = () => {
    currentSheet = sheetSelect.value;
    loadSheet(wb, currentSheet);
  };

  loadSheet(wb, currentSheet);
});

function loadSheet(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(ws, { defval: "" }); // keep blanks as ""

  allRows = json;
  const headers = json.length ? Object.keys(json[0]) : [];
  identifierCols = detectIdentifierCols(headers);

  // Setup sorting dropdown (identifier columns only)
  sortSelect.innerHTML = "";
  identifierCols.forEach((c) => {
    const opt = document.createElement("option");
    opt.value = c;
    opt.textContent = `Sort: ${c}`;
    sortSelect.appendChild(opt);
  });

  sortCol = identifierCols[0] || null;
  sortSelect.disabled = !sortCol;
  sortDirBtn.disabled = !sortCol;

  sortSelect.onchange = () => {
    sortCol = sortSelect.value;
    applyAll(true);
  };
  sortDirBtn.onclick = () => {
    sortAsc = !sortAsc;
    sortDirBtn.textContent = sortAsc ? "↑" : "↓";
    applyAll(true);
  };

  buildFilters();
  activeIndex = null;
  applyAll();
}

function applyAll(keepSelection = false) {
  const afterFilter = applyFilters(allRows);
  const afterSort = applySort(afterFilter);

  filteredRows = afterSort;

  if (!keepSelection) activeIndex = null;
  renderRowList(filteredRows);
}
