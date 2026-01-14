// --- State ---
let allRows = [];
let filteredRows = [];
let identifierCols = [];
let sheetNames = [];
let currentSheet = null;

const el = (id) => document.getElementById(id);

const fileInput = el("fileInput");
const sheetSelect = el("sheetSelect");
const filtersDiv = el("filters");
const clearFiltersBtn = el("clearFilters");
const rowList = el("rowList");
const chat = el("chat");
const meta = el("meta");

// --- Helpers ---
function isBlank(v) {
  return v === null || v === undefined || (typeof v === "string" && v.trim() === "");
}

function detectIdentifierCols(headers) {
  // Exclude Q/R/Expected columns - these are chat content, not metadata
  return headers.filter((h) => {
    const normalized = h.trim();
    return !/^Q\d+$/i.test(normalized) && 
           !/^R\d+$/i.test(normalized) && 
           !/^Expected$/i.test(normalized);
  });
}

function rowToMessages(row) {
  // Dynamically consume Q1/R1/Expected, Q2/R2/Expected, ...
  // Expected appears after each R, so pattern is: Q1, R1, Expected, Q2, R2, Expected, ...
  // SheetJS handles duplicate column names as: Expected, Expected_1, Expected_2, etc.
  // So Expected corresponds to R1, Expected_1 to R2, Expected_2 to R3, etc.
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

    if (!isBlank(q)) msgs.push({ role: "user", text: String(q), expected: null });
    
    // For responses, find the corresponding Expected value
    let expectedValue = null;
    if (!isBlank(r)) {
      // Pattern: Expected for R1, Expected_1 for R2, Expected_2 for R3, etc.
      // Also try Expected{k} pattern (Expected1, Expected2) as fallback
      if (k === 1) {
        // First Expected column (no suffix)
        const expectedFirst = row[`Expected`];
        if (!isBlank(expectedFirst)) {
          expectedValue = String(expectedFirst).trim();
        }
      } else {
        // Subsequent Expected columns: Expected_1, Expected_2, etc.
        const expectedUnderscore = row[`Expected_${k - 1}`];
        if (!isBlank(expectedUnderscore)) {
          expectedValue = String(expectedUnderscore).trim();
        }
      }
      
      // Fallback: try Expected{k} pattern
      if (expectedValue === null) {
        const expectedK = row[`Expected${k}`];
        if (!isBlank(expectedK)) {
          expectedValue = String(expectedK).trim();
        }
      }
      
      msgs.push({ 
        role: "assistant", 
        text: String(r),
        expected: expectedValue
      });
    }
    k++;
  }
  return msgs;
}

function getRowTitle(row) {
  // pick a nice display title - prefer Id, then Chat Id, then row number
  if (!isBlank(row["Id"])) return `ID: ${row["Id"]}`;
  if (!isBlank(row["Chat Id"])) return `Chat: ${row["Chat Id"]}`;
  if (!isBlank(row["ID"])) return `ID: ${row["ID"]}`;
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

  // Filter out Note column from filters
  const filterableCols = identifierCols.filter((col) => col.toLowerCase() !== "note");

  filterableCols.forEach((col) => {
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

  clearFiltersBtn.disabled = filterableCols.length === 0;
  clearFiltersBtn.onclick = () => {
    // reset: select all in every filter
    filtersDiv.querySelectorAll("select").forEach((sel) => {
      Array.from(sel.options).forEach((o) => (o.selected = true));
    });
    filterableCols.forEach((c) => {
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
    // Filter out Note column from filtering logic
    const filterableCols = identifierCols.filter((col) => col.toLowerCase() !== "note");
    for (const col of filterableCols) {
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
    title.textContent = getRowTitle(row);

    const sub = document.createElement("div");
    sub.className = "rowsub";
    const parts = [];
    if (!isBlank(row["Priority"])) parts.push(`Priority: ${row["Priority"]}`);
    if (!isBlank(row["Project Name"])) parts.push(row["Project Name"]);
    if (!isBlank(row["User Email"])) parts.push(row["User Email"]);
    sub.textContent = parts.length > 0 ? parts.join(" • ") : "No metadata";

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

  // Display metadata columns nicely formatted
  const metadataCols = identifierCols;
  metadataCols.forEach((c) => {
    const v = stableValue(row[c]);
    if (v === "") return;
    const b = document.createElement("div");
    b.className = "meta-item";
    const label = document.createElement("span");
    label.className = "meta-label";
    label.textContent = `${c}:`;
    const value = document.createElement("span");
    value.className = "meta-value";
    value.textContent = v;
    b.appendChild(label);
    b.appendChild(value);
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
    
    // Add color coding for assistant responses based on Expected value
    if (m.role === "assistant") {
      if (m.expected !== null && m.expected !== undefined && String(m.expected).trim() !== "") {
        const expectedLower = String(m.expected).toLowerCase().trim();
        if (expectedLower === "yes") {
          bubble.classList.add("expected-yes");
        } else if (expectedLower === "no") {
          bubble.classList.add("expected-no");
        } else {
          bubble.classList.add("expected-unknown");
        }
      } else {
        bubble.classList.add("expected-unknown");
      }
    }

    const role = document.createElement("div");
    role.className = "role";
    role.textContent = m.role === "user" ? "User" : "Assistant";
    
    // Add Expected indicator for assistant messages
    if (m.role === "assistant" && m.expected !== null && m.expected !== undefined && String(m.expected).trim() !== "") {
      const expectedLabel = document.createElement("span");
      expectedLabel.className = "expected-label";
      const expectedLower = String(m.expected).toLowerCase().trim();
      if (expectedLower === "yes") {
        expectedLabel.textContent = "✓ Expected";
        expectedLabel.classList.add("expected-yes-label");
      } else if (expectedLower === "no") {
        expectedLabel.textContent = "✗ Unexpected";
        expectedLabel.classList.add("expected-no-label");
      } else {
        expectedLabel.textContent = `Expected: ${m.expected}`;
        expectedLabel.classList.add("expected-unknown-label");
      }
      role.appendChild(expectedLabel);
    }

    const text = document.createElement("div");
    text.className = "message-text";
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

  buildFilters();
  activeIndex = null;
  applyAll();
}

function applyAll(keepSelection = false) {
  const afterFilter = applyFilters(allRows);
  filteredRows = afterFilter;

  if (!keepSelection) activeIndex = null;
  renderRowList(filteredRows);
}
