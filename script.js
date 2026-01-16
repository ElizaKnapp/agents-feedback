// --- State ---
let allRows = [];
let filteredRows = [];
let identifierCols = [];
let sheetNames = [];
let currentSheet = null;
let nextId = 1; // Track next ID for new chats

// Environment configuration state
let envConfig = {
  location: null,
  userEmail: null,
  apiKey: null,
  baseUrl: null
};
let availableProjects = []; // Array of {project_id, name}
let availableChats = []; // Array of {chat_id, name}
let filteredChats = []; // Filtered chats based on search

const el = (id) => document.getElementById(id);

const fileInput = el("fileInput");
const sheetSelect = el("sheetSelect");
const filtersDiv = el("filters");
const rowList = el("rowList");
const chat = el("chat");
const meta = el("meta");
const configEnvForm = el("configEnvForm");
const addChatForm = el("addChatForm");
const downloadBtn = el("downloadBtn");

// --- Helpers ---
function isBlank(v) {
  return v === null || v === undefined || (typeof v === "string" && v.trim() === "");
}

function detectIdentifierCols(headers) {
  // Exclude Q/R/N/Expected columns - these are chat content, not metadata
  return headers.filter((h) => {
    const normalized = h.trim();
    return !/^Q\d+$/i.test(normalized) && 
           !/^R\d+$/i.test(normalized) && 
           !/^N\d+$/i.test(normalized) &&
           !/^Expected$/i.test(normalized) &&
           !/^Expected_\d+$/i.test(normalized) &&
           !/^Expected\d+$/i.test(normalized);
  });
}

function rowToMessages(row) {
  // Dynamically consume Q1/R1/N1/Expected, Q2/R2/N2/Expected, ...
  // Pattern is: Q1, R1, N1, Expected, Q2, R2, N2, Expected, ...
  // SheetJS handles duplicate column names as: Expected, Expected_1, Expected_2, etc.
  // So Expected corresponds to R1, Expected_1 to R2, Expected_2 to R3, etc.
  const msgs = [];
  let k = 1;
  while (true) {
    const q = row[`Q${k}`];
    const r = row[`R${k}`];
    const n = row[`N${k}`];

    // stop if both Q and R missing/blank AND there is no higher pair
    if (isBlank(q) && isBlank(r)) {
      // lookahead: if next pair also blank, break; otherwise skip gap
      const qn = row[`Q${k + 1}`];
      const rn = row[`R${k + 1}`];
      if (isBlank(qn) && isBlank(rn)) break;
      k++;
      continue;
    }

    if (!isBlank(q)) msgs.push({ role: "user", text: String(q), expected: null, note: null });
    
    // For responses, find the corresponding Note and Expected value
    let expectedValue = null;
    let noteValue = null;
    if (!isBlank(r)) {
      // Get the note (N1, N2, etc.)
      if (!isBlank(n)) {
        noteValue = String(n).trim();
      }
      
      // Pattern: Expected for R1, Expected_1 for R2, Expected_2 for R3, etc.
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
        expected: expectedValue,
        note: noteValue,
        exchangeIndex: k - 1 // Track which exchange this is (0-indexed)
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

// Format date values from Excel (which come as numbers) to readable text
function formatDateValue(value, colName) {
  if (isBlank(value)) return "";
  
  // Check if column name suggests it's a date
  const isDateColumn = /date/i.test(colName);
  
  if (!isDateColumn) return stableValue(value);
  
  // Excel dates are numbers representing days since 1900-01-01
  // Check if it's a number that could be a date (between reasonable bounds)
  const num = Number(value);
  if (!isNaN(num) && num > 0 && num < 1000000) {
    try {
      // Excel epoch: January 1, 1900 is day 1
      // But Excel incorrectly treats 1900 as a leap year, so we use Dec 30, 1899 as epoch
      const excelEpoch = new Date(1899, 11, 30); // Dec 30, 1899
      const date = new Date(excelEpoch.getTime() + num * 24 * 60 * 60 * 1000);
      
      // Check if the date is reasonable (between 1900 and 2100)
      if (date.getFullYear() >= 1900 && date.getFullYear() <= 2100 && !isNaN(date.getTime())) {
        return date.toLocaleDateString('en-US', { 
          year: 'numeric', 
          month: 'long', 
          day: 'numeric' 
        });
      }
    } catch (e) {
      // If date conversion fails, fall through to return as string
    }
  }
  
  // If not a valid date number, return as string
  return stableValue(value);
}

// --- Filtering UI ---
let activeFilters = {}; // { colName: Set(values) }

const FILTER_COLUMNS = ["Priority", "Location Run", "User Email", "Project Name"];

function buildFilters() {
  filtersDiv.innerHTML = "";
  activeFilters = {};

  FILTER_COLUMNS.forEach((col) => {
    // Check if column exists in the data
    if (!allRows.length || !(col in allRows[0])) return;

    // Get unique values from all rows
    const vals = Array.from(
      new Set(allRows.map((r) => stableValue(r[col])).filter((v) => v !== ""))
    ).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

    if (vals.length === 0) return;

    const wrap = document.createElement("div");
    wrap.className = "filter";

    const label = document.createElement("label");
    label.textContent = col;

    const select = document.createElement("select");
    select.multiple = false; // Single select dropdown
    select.size = 1;
    
    // Add "All" option
    const allOpt = document.createElement("option");
    allOpt.value = "";
    allOpt.textContent = "All";
    allOpt.selected = true;
    select.appendChild(allOpt);

    vals.forEach((v) => {
      const opt = document.createElement("option");
      opt.value = v;
      opt.textContent = v;
      select.appendChild(opt);
    });

    activeFilters[col] = null; // null means "all selected"

    select.addEventListener("change", () => {
      const selectedValue = select.value;
      activeFilters[col] = selectedValue === "" ? null : selectedValue;
      applyAll();
    });

    wrap.appendChild(label);
    wrap.appendChild(select);
    filtersDiv.appendChild(wrap);
  });
}

function applyFilters(rows) {
  return rows.filter((r) => {
    for (const col of FILTER_COLUMNS) {
      const filterValue = activeFilters[col];
      if (filterValue === null) continue; // No filter applied
      
      const rowValue = stableValue(r[col]);
      if (rowValue !== filterValue) return false;
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

let currentRow = null; // Store reference to currently displayed row

function renderChat(row) {
  chat.innerHTML = "";
  meta.innerHTML = "";
  currentRow = row; // Store reference to current row

  // Display metadata columns nicely formatted, excluding Expected columns
  const metadataCols = identifierCols.filter((c) => {
    const normalized = c.trim();
    // Exclude Expected columns from metadata display
    return !/^Expected$/i.test(normalized) &&
           !/^Expected_\d+$/i.test(normalized) &&
           !/^Expected\d+$/i.test(normalized);
  });
  
  metadataCols.forEach((c) => {
    const b = document.createElement("div");
    b.className = "meta-item";
    const label = document.createElement("span");
    label.className = "meta-label";
    label.textContent = `${c}:`;
    
    // Create editable input field
    const value = document.createElement("input");
    value.type = "text";
    value.className = "meta-value editable-meta";
    
    // Special handling for Date column - show formatted date but store raw value
    if (/date/i.test(c)) {
      const rawValue = row[c] || "";
      const displayValue = formatDateValue(rawValue, c);
      if (displayValue && displayValue !== stableValue(rawValue)) {
        // If it's a formatted date, show it but store the original
        value.value = displayValue;
        value.dataset.originalValue = rawValue;
      } else {
        value.value = rawValue;
      }
    } else {
      value.value = row[c] || "";
    }
    
    value.placeholder = `Enter ${c}...`;
    
    // Store column name for updates
    value.dataset.column = c;
    
    // Update row on blur
    value.addEventListener("blur", () => {
      updateMetadataInRow(c, value.value);
      // Update download button
      downloadBtn.disabled = false;
    });
    
    // Handle Enter key
    value.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        value.blur();
      }
    });
    
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
    
    // Add Expected indicator for assistant messages (editable)
    if (m.role === "assistant") {
      const expectedSelect = document.createElement("select");
      expectedSelect.className = "expected-select";
      expectedSelect.dataset.exchangeIndex = m.exchangeIndex;
      
      const options = [
        { value: "", text: "Not Set" },
        { value: "Yes", text: "✓ Expected" },
        { value: "No", text: "✗ Unexpected" }
      ];
      
      const currentExpected = m.expected !== null && m.expected !== undefined ? String(m.expected).trim() : "";
      options.forEach(opt => {
        const option = document.createElement("option");
        option.value = opt.value;
        option.textContent = opt.text;
        if (currentExpected.toLowerCase() === opt.value.toLowerCase()) {
          option.selected = true;
        }
        expectedSelect.appendChild(option);
      });
      
      // Apply styling based on current value
      if (currentExpected.toLowerCase() === "yes") {
        expectedSelect.classList.add("expected-yes-label");
      } else if (currentExpected.toLowerCase() === "no") {
        expectedSelect.classList.add("expected-no-label");
      } else {
        expectedSelect.classList.add("expected-unknown-label");
      }
      
      expectedSelect.addEventListener("change", () => {
        updateExpectedInRow(row, m.exchangeIndex, expectedSelect.value);
        // Update styling
        expectedSelect.className = "expected-select";
        if (expectedSelect.value.toLowerCase() === "yes") {
          expectedSelect.classList.add("expected-yes-label");
        } else if (expectedSelect.value.toLowerCase() === "no") {
          expectedSelect.classList.add("expected-no-label");
        } else {
          expectedSelect.classList.add("expected-unknown-label");
        }
        // Update bubble color
        const bubbleEl = expectedSelect.closest(".bubble");
        bubbleEl.className = "bubble assistant";
        if (expectedSelect.value.toLowerCase() === "yes") {
          bubbleEl.classList.add("expected-yes");
        } else if (expectedSelect.value.toLowerCase() === "no") {
          bubbleEl.classList.add("expected-no");
        } else {
          bubbleEl.classList.add("expected-unknown");
        }
        // Note container colors are handled by bubble class CSS
      });
      
      role.appendChild(expectedSelect);
    }

    const text = document.createElement("div");
    text.className = "message-text";
    
    // Render markdown for assistant messages, plain text for user messages
    if (m.role === "assistant") {
      if (typeof marked !== "undefined" && marked.parse) {
        try {
          text.innerHTML = marked.parse(m.text);
        } catch (e) {
          // Fallback to plain text if markdown parsing fails
          text.textContent = m.text;
        }
      } else {
        // Fallback if marked is not loaded
        text.textContent = m.text;
      }
    } else {
      text.textContent = m.text;
    }

    bubble.appendChild(role);
    bubble.appendChild(text);

    // Add note for assistant messages (always editable)
    if (m.role === "assistant") {
      const noteContainer = document.createElement("div");
      noteContainer.className = "note-container";
      
      const noteLabel = document.createElement("span");
      noteLabel.className = "note-label";
      noteLabel.textContent = "Note:";
      
      const noteText = document.createElement("textarea");
      noteText.className = "note-text editable-note";
      noteText.value = m.note !== null && m.note !== undefined ? String(m.note).trim() : "";
      noteText.placeholder = "Add a note...";
      noteText.rows = 2;
      
      // Store row reference and exchange index for editing
      noteText.dataset.exchangeIndex = m.exchangeIndex;
      
      noteText.addEventListener("blur", () => {
        updateNoteInRow(row, m.exchangeIndex, noteText.value);
      });
      
      noteContainer.appendChild(noteLabel);
      noteContainer.appendChild(noteText);
      bubble.appendChild(noteContainer);
    }

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

  // Find max ID to set nextId
  if (allRows.length > 0) {
    const ids = allRows
      .map(r => {
        const id = r["Id"] || r["ID"];
        return id ? parseInt(id) : 0;
      })
      .filter(id => !isNaN(id));
    nextId = ids.length > 0 ? Math.max(...ids) + 1 : 1;
  } else {
    nextId = 1;
  }

  buildFilters();
  activeIndex = null;
  applyAll();
  downloadBtn.disabled = false;
}

function applyAll(keepSelection = false) {
  const afterFilter = applyFilters(allRows);
  filteredRows = afterFilter;

  if (!keepSelection) activeIndex = null;
  renderRowList(filteredRows);
}

// --- Update functions for editing ---
function updateNoteInRow(row, exchangeIndex, noteValue) {
  // exchangeIndex is 0-based, but columns are 1-based (N1, N2, etc.)
  const columnName = `N${exchangeIndex + 1}`;
  row[columnName] = noteValue;
  // Trigger download button update
  downloadBtn.disabled = false;
}

function updateExpectedInRow(row, exchangeIndex, expectedValue) {
  // exchangeIndex is 0-based
  let columnName;
  if (exchangeIndex === 0) {
    columnName = "Expected";
  } else {
    columnName = `Expected_${exchangeIndex}`;
  }
  row[columnName] = expectedValue;
  // Trigger download button update
  downloadBtn.disabled = false;
}

function updateMetadataInRow(columnName, value) {
  if (!currentRow) return;
  
  // For Date column, if user entered a formatted date, try to convert it back
  if (/date/i.test(columnName)) {
    // Try to parse the formatted date back to a value we can store
    const parsedDate = parseDateString(value);
    if (parsedDate) {
      // Convert to ISO string for storage
      const isoDate = parsedDate.toISOString().split('T')[0];
      currentRow[columnName] = isoDate;
      return;
    }
  }
  
  // For other columns, store as-is
  currentRow[columnName] = value;
}

// --- API Integration ---
function getHeaders() {
  return {
    'User': envConfig.userEmail,
    'Api-key': envConfig.apiKey,
    'Content-Type': 'application/json'
  };
}

async function fetchProjects() {
  const url = `${envConfig.baseUrl}/component/get-projects-for-user`;
  const headers = getHeaders();

  try {
    const response = await fetch(url, { headers });
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const data = await response.json();
    return data.projects || [];
  } catch (error) {
    console.error("Error fetching projects:", error);
    throw error;
  }
}

async function fetchChats(projectId) {
  const url = `${envConfig.baseUrl}/component/get-chats?project_id=${projectId}`;
  const headers = getHeaders();

  try {
    const response = await fetch(url, { headers });
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const data = await response.json();
    return data.chats || [];
  } catch (error) {
    console.error("Error fetching chats:", error);
    throw error;
  }
}

async function fetchChatExchanges(projectId, chatId) {
  const url = `${envConfig.baseUrl}/component/get-exchanges-for-chat?project_id=${projectId}&chat_id=${chatId}`;
  const headers = getHeaders();

  try {
    const response = await fetch(url, { headers });
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const data = await response.json();
    return data;
  } catch (error) {
    console.error("Error fetching exchanges:", error);
    throw error;
  }
}

function exchangesToRow(exchanges, location, userEmail, projectId, chatId, projectName = "") {
  // Sort exchanges by created_at (earliest to latest)
  const sortedExchanges = [...exchanges].sort((a, b) => {
    const timeA = parseInt(a.created_at) || 0;
    const timeB = parseInt(b.created_at) || 0;
    return timeA - timeB;
  });

  // Create new row - store date as ISO string for easy conversion
  const today = new Date();
  const dateStr = today.toISOString().split('T')[0]; // YYYY-MM-DD format
  
  const newRow = {
    "Id": nextId++,
    "Priority": "None",
    "Project Name": projectName, // Auto-set from selected project
    "Project Description": "",
    "Date": dateStr,
    "Location Run": location,
    "User Email": userEmail,
    "Project Id": projectId,
    "Chat Id": chatId,
    "Note": ""
  };

  // Add Q/R/N/Expected pairs
  sortedExchanges.forEach((exchange, index) => {
    const k = index + 1;
    newRow[`Q${k}`] = exchange.query || "";
    newRow[`R${k}`] = exchange.response || "";
    newRow[`N${k}`] = ""; // Notes start empty
    if (k === 1) {
      newRow["Expected"] = "";
    } else {
      newRow[`Expected_${k - 1}`] = "";
    }
  });

  return newRow;
}

// --- Download Excel ---
function downloadExcel() {
  if (allRows.length === 0) {
    alert("No data to download");
    return;
  }

  // Create workbook
  const wb = XLSX.utils.book_new();
  
  // Get all unique column names from all rows
  const allColumns = new Set();
  allRows.forEach(row => {
    Object.keys(row).forEach(key => allColumns.add(key));
  });
  
  // Ensure standard columns are first
  const standardCols = ["Id", "Priority", "Project Name", "Project Description", "Date", 
                         "Location Run", "User Email", "Project Id", "Chat Id", "Note"];
  const otherCols = Array.from(allColumns).filter(col => !standardCols.includes(col));
  
  // Sort Q/R/N/Expected columns properly
  const qCols = otherCols.filter(c => /^Q\d+$/i.test(c)).sort((a, b) => {
    const numA = parseInt(a.match(/\d+/)[0]);
    const numB = parseInt(b.match(/\d+/)[0]);
    return numA - numB;
  });
  const rCols = otherCols.filter(c => /^R\d+$/i.test(c)).sort((a, b) => {
    const numA = parseInt(a.match(/\d+/)[0]);
    const numB = parseInt(b.match(/\d+/)[0]);
    return numA - numB;
  });
  const nCols = otherCols.filter(c => /^N\d+$/i.test(c)).sort((a, b) => {
    const numA = parseInt(a.match(/\d+/)[0]);
    const numB = parseInt(b.match(/\d+/)[0]);
    return numA - numB;
  });
  const expectedCols = otherCols.filter(c => /^Expected/i.test(c)).sort((a, b) => {
    if (a === "Expected") return -1;
    if (b === "Expected") return 1;
    const numA = parseInt(a.match(/\d+/)?.[0]) || 0;
    const numB = parseInt(b.match(/\d+/)?.[0]) || 0;
    return numA - numB;
  });
  
  // Interleave Q/R/N/Expected columns
  const maxPairs = Math.max(qCols.length, rCols.length, nCols.length);
  const interleavedCols = [];
  for (let i = 0; i < maxPairs; i++) {
    if (qCols[i]) interleavedCols.push(qCols[i]);
    if (rCols[i]) interleavedCols.push(rCols[i]);
    if (nCols[i]) interleavedCols.push(nCols[i]);
    if (i === 0 && expectedCols.find(c => c === "Expected")) {
      interleavedCols.push("Expected");
    } else if (expectedCols.find(c => c === `Expected_${i}`)) {
      interleavedCols.push(`Expected_${i}`);
    }
  }
  
  const orderedColumns = [...standardCols, ...interleavedCols];
  
  // Convert rows to array of arrays, converting dates to Excel serial numbers
  const wsData = [orderedColumns]; // Header row
  const dateColIndex = orderedColumns.indexOf("Date");
  
  allRows.forEach(row => {
    const rowData = orderedColumns.map((col, colIndex) => {
      const value = row[col] || "";
      
      // Convert Date column to Excel date serial number
      if (col === "Date" && value) {
        return convertToExcelDate(value);
      }
      
      return value;
    });
    wsData.push(rowData);
  });
  
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  
  // Format Date column cells as dates (dd-mmm-yy format)
  if (dateColIndex >= 0) {
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let row = 1; row <= range.e.r; row++) { // Skip header row
      const cellAddress = XLSX.utils.encode_cell({ r: row, c: dateColIndex });
      if (ws[cellAddress] && ws[cellAddress].v !== null && ws[cellAddress].v !== "") {
        // Ensure it's a number (Excel date serial)
        if (typeof ws[cellAddress].v === 'number') {
          ws[cellAddress].z = 'dd-mmm-yy'; // Excel date format code
          ws[cellAddress].t = 'n'; // Number type (Excel dates are numbers)
        }
      }
    }
    
    // Set column width for Date column
    if (!ws['!cols']) ws['!cols'] = [];
    ws['!cols'][dateColIndex] = { wch: 12 };
  }
  
  XLSX.utils.book_append_sheet(wb, ws, currentSheet || "Sheet1");
  
  // Download
  XLSX.writeFile(wb, "chat-data.xlsx");
}

// Convert date string or number to Excel date serial number
function convertToExcelDate(value) {
  if (!value) return "";
  
  // If it's already a number (Excel serial), return it
  const num = Number(value);
  if (!isNaN(num) && num > 0 && num < 1000000) {
    // Check if it's already an Excel date serial number
    const testDate = new Date(1899, 11, 30);
    testDate.setDate(testDate.getDate() + num);
    if (testDate.getFullYear() >= 1900 && testDate.getFullYear() <= 2100) {
      return num;
    }
  }
  
  // Try to parse as date string
  let date;
  if (typeof value === 'string') {
    // Try various date formats
    date = new Date(value);
    if (isNaN(date.getTime())) {
      // Try parsing formats like "January 15, 2024"
      date = parseDateString(value);
    }
  } else {
    date = new Date(value);
  }
  
  if (isNaN(date.getTime())) {
    return value; // Return original if can't parse
  }
  
  // Convert to Excel serial number
  // Excel epoch: Dec 30, 1899 (Excel incorrectly treats 1900 as leap year)
  const excelEpoch = new Date(1899, 11, 30);
  const diffTime = date.getTime() - excelEpoch.getTime();
  const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
  
  return diffDays;
}

// Parse date strings like "January 15, 2024" or "2024-01-15"
function parseDateString(dateStr) {
  if (!dateStr || typeof dateStr !== 'string') return null;
  
  // Try ISO format first (YYYY-MM-DD)
  const isoMatch = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    return new Date(parseInt(isoMatch[1]), parseInt(isoMatch[2]) - 1, parseInt(isoMatch[3]));
  }
  
  // Try month name format "January 15, 2024"
  const monthNameMatch = dateStr.match(/(\w+)\s+(\d+),\s+(\d+)/);
  if (monthNameMatch) {
    const monthNames = ["january", "february", "march", "april", "may", "june",
      "july", "august", "september", "october", "november", "december"];
    const month = monthNames.indexOf(monthNameMatch[1].toLowerCase());
    if (month >= 0) {
      return new Date(parseInt(monthNameMatch[3]), month, parseInt(monthNameMatch[2]));
    }
  }
  
  // Try numeric formats
  const numericMatch = dateStr.match(/(\d+)[\/-](\d+)[\/-](\d+)/);
  if (numericMatch) {
    const parts = numericMatch.slice(1).map(Number);
    if (parts[0] > 1000) {
      // YYYY-MM-DD or YYYY/MM/DD
      return new Date(parts[0], parts[1] - 1, parts[2]);
    } else if (parts[2] > 1000) {
      // MM/DD/YYYY
      return new Date(parts[2], parts[0] - 1, parts[1]);
    }
  }
  
  // Fallback to default parsing
  const parsed = new Date(dateStr);
  return isNaN(parsed.getTime()) ? null : parsed;
}

// --- Configure Environment Form ---
configEnvForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  
  const location = el("location").value;
  const userEmail = el("userEmail").value;
  const apiKey = el("apiKey").value;
  
  const configBtn = el("configEnvBtn");
  configBtn.disabled = true;
  configBtn.textContent = "Loading...";
  
  try {
    // Set environment config
    envConfig.location = location;
    envConfig.userEmail = userEmail;
    envConfig.apiKey = apiKey;
    envConfig.baseUrl = location === "Development" 
      ? "https://dev-api.ikigailabs.io" 
      : "https://api.ikigailabs.io";
    
    // Fetch projects
    const projects = await fetchProjects();
    availableProjects = projects
      .filter(p => p.project_id && p.name)
      .map(p => ({ project_id: p.project_id, name: p.name }));
    
    // Populate project dropdown with names
    const projectSelect = el("projectId");
    projectSelect.innerHTML = '<option value="">Select a project...</option>';
    availableProjects.forEach(project => {
      const option = document.createElement("option");
      option.value = project.project_id; // Store ID as value
      option.textContent = project.name; // Display name
      projectSelect.appendChild(option);
    });
    projectSelect.disabled = false;
    
    // Reset chat dropdown and search
    const chatSelect = el("chatId");
    const chatSearch = el("chatSearch");
    chatSelect.innerHTML = '<option value="">Select a project first</option>';
    chatSelect.disabled = true;
    chatSearch.disabled = true;
    chatSearch.value = "";
    availableChats = [];
    filteredChats = [];
    
    alert(`Environment configured! Found ${availableProjects.length} project(s).`);
  } catch (error) {
    alert(`Error configuring environment: ${error.message}`);
  } finally {
    configBtn.disabled = false;
    configBtn.textContent = "Configure";
  }
});

// --- Project selection handler ---
el("projectId").addEventListener("change", async (e) => {
  const projectId = e.target.value;
  const chatSelect = el("chatId");
  const chatSearch = el("chatSearch");
  const addChatBtn = el("addChatBtn");
  
  if (!projectId) {
    chatSelect.innerHTML = '<option value="">Select a project first</option>';
    chatSelect.disabled = true;
    chatSearch.disabled = true;
    chatSearch.value = "";
    addChatBtn.disabled = true;
    availableChats = [];
    filteredChats = [];
    return;
  }
  
  chatSelect.disabled = true;
  chatSearch.disabled = true;
  chatSelect.innerHTML = '<option value="">Loading chats...</option>';
  chatSearch.value = "";
  
  try {
    const chats = await fetchChats(projectId);
    availableChats = chats
      .filter(c => c.chat_id && c.name)
      .map(c => ({ chat_id: c.chat_id, name: c.name }));
    filteredChats = [...availableChats];
    
    populateChatDropdown();
    chatSearch.disabled = false;
    
    if (availableChats.length > 0) {
      addChatBtn.disabled = false;
    }
  } catch (error) {
    alert(`Error fetching chats: ${error.message}`);
    chatSelect.innerHTML = '<option value="">Error loading chats</option>';
    chatSelect.disabled = true;
    chatSearch.disabled = true;
    addChatBtn.disabled = true;
    availableChats = [];
    filteredChats = [];
  }
});

// --- Populate chat dropdown helper ---
function populateChatDropdown() {
  const chatSelect = el("chatId");
  chatSelect.innerHTML = '<option value="">Select a chat...</option>';
  filteredChats.forEach(chat => {
    const option = document.createElement("option");
    option.value = chat.chat_id; // Store ID as value
    option.textContent = chat.name; // Display name
    chatSelect.appendChild(option);
  });
  chatSelect.disabled = false;
}

// --- Chat search handler ---
el("chatSearch").addEventListener("input", (e) => {
  const searchTerm = e.target.value.toLowerCase().trim();
  
  if (!searchTerm) {
    filteredChats = [...availableChats];
  } else {
    filteredChats = availableChats.filter(chat => 
      chat.name.toLowerCase().includes(searchTerm)
    );
  }
  
  populateChatDropdown();
  
  // Clear chat selection if current selection is not in filtered results
  const chatSelect = el("chatId");
  const selectedChatId = chatSelect.value;
  if (selectedChatId && !filteredChats.find(c => c.chat_id === selectedChatId)) {
    chatSelect.value = "";
    el("addChatBtn").disabled = true;
  }
});

// --- Chat selection handler ---
el("chatId").addEventListener("change", (e) => {
  const addChatBtn = el("addChatBtn");
  addChatBtn.disabled = !e.target.value;
});

// --- Add Chat Form submission ---
addChatForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  
  const projectId = el("projectId").value;
  const chatId = el("chatId").value;
  
  if (!envConfig.baseUrl || !envConfig.userEmail || !envConfig.apiKey) {
    alert("Please configure environment first!");
    return;
  }
  
  const addChatBtn = el("addChatBtn");
  addChatBtn.disabled = true;
  addChatBtn.textContent = "Loading...";
  
  try {
    const data = await fetchChatExchanges(projectId, chatId);
    
    // Find project name from availableProjects
    const selectedProject = availableProjects.find(p => p.project_id === projectId);
    const projectName = selectedProject ? selectedProject.name : "";
    
    const newRow = exchangesToRow(data.exchanges || [], envConfig.location, envConfig.userEmail, projectId, chatId, projectName);
    
    // Add to allRows
    allRows.push(newRow);
    
    // Refresh filters and row list
    buildFilters();
    applyAll();
    
    // Select the new row
    activeIndex = filteredRows.length - 1;
    renderRowList(filteredRows);
    renderChat(newRow);
    
    // Clear chat selection and search (keep project selected)
    el("chatId").value = "";
    el("chatSearch").value = "";
    filteredChats = [...availableChats];
    populateChatDropdown();
    addChatBtn.disabled = true;
    
    // Enable download button
    downloadBtn.disabled = false;
    
    alert("Chat added successfully!");
  } catch (error) {
    alert(`Error adding chat: ${error.message}`);
  } finally {
    addChatBtn.disabled = false;
    addChatBtn.textContent = "Add Chat";
  }
});

// --- Download button ---
downloadBtn.addEventListener("click", downloadExcel);

// --- Collapsible forms ---
const configEnvToggle = el("configEnvToggle");
const configEnvFormContainer = el("configEnvFormContainer");
let configEnvFormExpanded = true;

configEnvToggle.addEventListener("click", () => {
  configEnvFormExpanded = !configEnvFormExpanded;
  if (configEnvFormExpanded) {
    configEnvFormContainer.style.display = "block";
    configEnvToggle.textContent = "−";
  } else {
    configEnvFormContainer.style.display = "none";
    configEnvToggle.textContent = "+";
  }
});

const addChatToggle = el("addChatToggle");
const addChatFormContainer = el("addChatFormContainer");
let addChatFormExpanded = true;

addChatToggle.addEventListener("click", () => {
  addChatFormExpanded = !addChatFormExpanded;
  if (addChatFormExpanded) {
    addChatFormContainer.style.display = "block";
    addChatToggle.textContent = "−";
  } else {
    addChatFormContainer.style.display = "none";
    addChatToggle.textContent = "+";
  }
});
