/* =============================================
   MEAT DISTRIBUTION LIST — script.js
   All data lives in localStorage as JSON array.
   ============================================= */

// ── STORAGE KEY ──────────────────────────────
const STORAGE_KEY = 'meatDistList_v1';

// ── STATE ─────────────────────────────────────
// `rows` is an array of objects: { id, name, status }
// `editingId` tracks which row is currently being edited
let rows      = [];
let editingId = null;

// ── INIT ──────────────────────────────────────
// Load saved data from localStorage when page loads
window.addEventListener('DOMContentLoaded', function () {
  loadFromStorage();
  renderTable();

  // Allow pressing Enter key in the name input to add a row
  document.getElementById('name-input').addEventListener('keydown', function (e) {
    if (e.key === 'Enter') addRow();
  });

  // Allow pressing Enter in edit modal to save
  document.getElementById('edit-input').addEventListener('keydown', function (e) {
    if (e.key === 'Enter') saveEdit();
  });

  // Close modal when clicking the dark overlay background
  document.getElementById('edit-modal').addEventListener('click', function (e) {
    if (e.target === this) closeModal();
  });
});

// ── LOAD FROM LOCALSTORAGE ────────────────────
function loadFromStorage() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    try {
      rows = JSON.parse(saved);   // parse the saved JSON string back into array
    } catch (err) {
      rows = [];                  // if data is corrupted, start fresh
    }
  }
}

// ── SAVE TO LOCALSTORAGE ──────────────────────
function saveToStorage() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
}

// ── GENERATE UNIQUE ID ────────────────────────
// Simple timestamp-based ID — unique enough for local use
function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
}

// ── ADD A NEW ROW ─────────────────────────────
function addRow() {
  const input = document.getElementById('name-input');
  const name  = input.value.trim();

  // Validate: don't add empty names
  if (!name) {
    input.focus();
    input.style.borderColor = '#b81c1c';
    setTimeout(() => input.style.borderColor = '', 700);
    return;
  }

  // Build the new row object
  const newRow = {
    id:     generateId(),
    name:   name,
    status: 'NO'        // default status is unpaid / NO
  };

  rows.push(newRow);     // add to array
  saveToStorage();       // persist to localStorage
  renderTable();         // refresh the table

  // Clear the input and return focus
  input.value = '';
  input.focus();
}

// ── TOGGLE STATUS (YES ↔ NO) ──────────────────
function toggleStatus(id) {
  const row = rows.find(r => r.id === id);
  if (!row) return;
  row.status = (row.status === 'YES') ? 'NO' : 'YES';   // flip the value
  saveToStorage();
  renderTable();
}

// ── DELETE A ROW ──────────────────────────────
function deleteRow(id) {
  const row   = rows.find(r => r.id === id);
  const label = row ? `"${row.name}"` : 'this record';

  // Simple confirmation before deleting
  if (!confirm(`Delete ${label}?`)) return;

  rows = rows.filter(r => r.id !== id);   // remove matching row
  saveToStorage();
  renderTable();
}

// ── OPEN EDIT MODAL ───────────────────────────
function openEdit(id) {
  const row = rows.find(r => r.id === id);
  if (!row) return;

  editingId = id;   // remember which row we're editing

  const editInput = document.getElementById('edit-input');
  editInput.value = row.name;

  // Show the modal
  document.getElementById('edit-modal').style.display = 'flex';
  setTimeout(() => editInput.focus(), 50);   // small delay so focus works reliably
}

// ── SAVE EDIT ─────────────────────────────────
function saveEdit() {
  const newName = document.getElementById('edit-input').value.trim();
  if (!newName) return;   // ignore empty saves

  const row = rows.find(r => r.id === editingId);
  if (row) {
    row.name = newName;
    saveToStorage();
    renderTable();
  }

  closeModal();
}

// ── CLOSE MODAL ───────────────────────────────
function closeModal() {
  document.getElementById('edit-modal').style.display = 'none';
  editingId = null;
}

// ── EXPORT TO EXCEL (.xlsx) ───────────────────
// Converts the `rows` array into an Excel file and triggers a download
function exportExcel() {
  if (rows.length === 0) {
    alert('Nothing to export. Add some rows first!');
    return;
  }

  // Build the data array: header row + one row per entry
  const sheetData = [
    ['Serial No', 'Name', 'Status']    // header row
  ];
  rows.forEach(function (row, index) {
    sheetData.push([index + 1, row.name, row.status]);
  });

  // Create a worksheet and workbook using SheetJS
  const ws = XLSX.utils.aoa_to_sheet(sheetData);   // aoa = array of arrays

  // Style column widths (SheetJS supports !cols)
  ws['!cols'] = [
    { wch: 10 },   // Serial No
    { wch: 30 },   // Name
    { wch: 10 }    // Status
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Distribution List');

  // Generate filename with today's date
  const today = new Date().toISOString().slice(0, 10);   // e.g. 2025-01-15
  XLSX.writeFile(wb, `meat-distribution-${today}.xlsx`);
}

// ── IMPORT FROM EXCEL (.xlsx) ─────────────────
// Reads an uploaded .xlsx file and appends rows to the current list
function importExcel(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const data  = new Uint8Array(e.target.result);
      const wb    = XLSX.read(data, { type: 'array' });

      // Read the first sheet
      const sheetName = wb.SheetNames[0];
      const ws        = wb.Sheets[sheetName];

      // Convert sheet to array of objects using first row as header
      const jsonData = XLSX.utils.sheet_to_json(ws, { defval: '' });

      if (jsonData.length === 0) {
        alert('The Excel file appears to be empty.');
        return;
      }

      let added = 0;
      let skipped = 0;

      jsonData.forEach(function (rowObj) {
        // Support both "Name" and "name" column headings (case-insensitive)
        const nameRaw   = rowObj['Name']   || rowObj['name']   || rowObj['NAME']   || '';
        const statusRaw = rowObj['Status'] || rowObj['status'] || rowObj['STATUS'] || 'NO';

        const name   = String(nameRaw).trim();
        const status = String(statusRaw).trim().toUpperCase() === 'YES' ? 'YES' : 'NO';

        if (!name) { skipped++; return; }   // skip rows with no name

        rows.push({ id: generateId(), name, status });
        added++;
      });

      saveToStorage();
      renderTable();

      // Reset file input so the same file can be re-imported if needed
      event.target.value = '';

      alert(`✅ Imported ${added} row(s) successfully.` + (skipped ? `\n⚠ ${skipped} row(s) skipped (empty name).` : ''));

    } catch (err) {
      alert('❌ Could not read the file. Make sure it is a valid .xlsx file.\n\nError: ' + err.message);
      event.target.value = '';
    }
  };

  reader.readAsArrayBuffer(file);   // read file as binary
}


function clearAll() {
  if (rows.length === 0) return;
  if (!confirm('Delete ALL records? This cannot be undone.')) return;
  rows = [];
  saveToStorage();
  renderTable();
}

// ── RENDER TABLE ──────────────────────────────
// Rebuilds the entire table from the `rows` array
function renderTable() {
  const tbody      = document.getElementById('table-body');
  const emptyState = document.getElementById('empty-state');

  // Clear existing rows
  tbody.innerHTML = '';

  if (rows.length === 0) {
    // Show the "no records" message, hide table body
    document.getElementById('distribution-table').style.display = 'none';
    emptyState.classList.add('visible');
    updateStats();
    return;
  }

  // Rows exist — show the table
  document.getElementById('distribution-table').style.display = 'table';
  emptyState.classList.remove('visible');

  // Build each table row
  rows.forEach(function (row, index) {
    const isYes = row.status === 'YES';

    const tr = document.createElement('tr');

    // ── Serial number cell
    const tdSerial = document.createElement('td');
    tdSerial.className = 'cell-serial';
    tdSerial.textContent = index + 1;   // 1-based serial number

    // ── Name cell
    const tdName = document.createElement('td');
    tdName.className = 'cell-name';
    tdName.textContent = row.name;

    // ── Status cell (toggle button)
    const tdStatus = document.createElement('td');
    const btnStatus = document.createElement('button');
    btnStatus.className = 'btn-status ' + (isYes ? 'yes' : 'no');
    btnStatus.textContent = isYes ? '✓ YES' : '✗ NO';
    btnStatus.title = isYes ? 'Click to mark Unpaid' : 'Click to mark Paid';
    btnStatus.onclick = function () { toggleStatus(row.id); };
    tdStatus.appendChild(btnStatus);

    // ── Actions cell (Edit + Delete buttons)
    const tdActions = document.createElement('td');
    tdActions.className = 'cell-actions';

    const btnEdit = document.createElement('button');
    btnEdit.className = 'btn-edit';
    btnEdit.textContent = '✏ Edit';
    btnEdit.onclick = function () { openEdit(row.id); };

    const btnDelete = document.createElement('button');
    btnDelete.className = 'btn-delete';
    btnDelete.textContent = '✕ Del';
    btnDelete.onclick = function () { deleteRow(row.id); };

    tdActions.appendChild(btnEdit);
    tdActions.appendChild(btnDelete);

    // Append all cells to the row
    tr.appendChild(tdSerial);
    tr.appendChild(tdName);
    tr.appendChild(tdStatus);
    tr.appendChild(tdActions);

    tbody.appendChild(tr);
  });

  updateStats();
}

// ── UPDATE STATS BADGES ───────────────────────
function updateStats() {
  const total   = rows.length;
  const yesCount = rows.filter(r => r.status === 'YES').length;
  const noCount  = total - yesCount;

  document.getElementById('stat-total').textContent = total;
  document.getElementById('stat-yes').textContent   = yesCount;
  document.getElementById('stat-no').textContent    = noCount;
}
