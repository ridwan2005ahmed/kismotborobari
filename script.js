/* =============================================
   MEAT DISTRIBUTION LIST — script.js
   All data lives in localStorage as JSON array.
   ============================================= */

// ── STORAGE KEY ──────────────────────────────
const STORAGE_KEY = 'meatDistList_v1';

// ── STATE ─────────────────────────────────────
let rows       = [];
let editingId  = null;
let modalStatus = 'NO';   // tracks YES/NO selection inside the Edit modal

// ── INIT ──────────────────────────────────────
window.addEventListener('DOMContentLoaded', function () {
  loadFromStorage();
  renderTable();

  // Enter key in name input → add row
  document.getElementById('name-input').addEventListener('keydown', function (e) {
    if (e.key === 'Enter') addRow();
  });

  // Enter key in edit modal → save
  document.getElementById('edit-input').addEventListener('keydown', function (e) {
    if (e.key === 'Enter') saveEdit();
  });

  // Click dark overlay → close modal
  document.getElementById('edit-modal').addEventListener('click', function (e) {
    if (e.target === this) closeModal();
  });
});

// ── LOAD FROM LOCALSTORAGE ────────────────────
function loadFromStorage() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    try {
      rows = JSON.parse(saved);
    } catch (err) {
      rows = [];
    }
  }
}

// ── SAVE TO LOCALSTORAGE ──────────────────────
function saveToStorage() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
}

// ── GENERATE UNIQUE ID ────────────────────────
function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
}

// ── ADD A NEW ROW ─────────────────────────────
function addRow() {
  const input = document.getElementById('name-input');
  const name  = input.value.trim();

  if (!name) {
    input.focus();
    input.style.borderColor = '#b81c1c';
    setTimeout(() => input.style.borderColor = '', 700);
    return;
  }

  rows.push({ id: generateId(), name: name, status: 'NO', locked: false });
  saveToStorage();
  renderTable();

  input.value = '';
  input.focus();
}

// ── SET STATUS INSIDE MODAL ───────────────────
// Called when YES or NO button is pressed inside Edit modal
function setModalStatus(val) {
  modalStatus = val;
  document.getElementById('toggle-yes').classList.toggle('active', val === 'YES');
  document.getElementById('toggle-no').classList.toggle('active',  val === 'NO');
}

// ── FIRST-TIME STATUS TOGGLE (table button) ───
// Only works if the row has never been confirmed (locked: false)
// After first press → locked: true → button becomes a plain badge
function firstTimeToggle(id) {
  const row = rows.find(r => r.id === id);
  if (!row || row.locked) return;   // already locked → ignore click
  row.status = (row.status === 'YES') ? 'NO' : 'YES';
  row.locked = true;                // lock after first change
  saveToStorage();
  renderTable();
}

// ── DELETE A ROW ──────────────────────────────
function deleteRow(id) {
  const row   = rows.find(r => r.id === id);
  const label = row ? '"' + row.name + '"' : 'this record';
  if (!confirm('Delete ' + label + '?')) return;
  rows = rows.filter(r => r.id !== id);
  saveToStorage();
  renderTable();
}

// ── OPEN EDIT MODAL ───────────────────────────
function openEdit(id) {
  const row = rows.find(r => r.id === id);
  if (!row) return;

  editingId   = id;
  modalStatus = row.status;   // pre-fill current status

  document.getElementById('edit-input').value = row.name;

  // Highlight the correct YES/NO button
  document.getElementById('toggle-yes').classList.toggle('active', row.status === 'YES');
  document.getElementById('toggle-no').classList.toggle('active',  row.status === 'NO');

  const modal = document.getElementById('edit-modal');
  modal.hidden = false;
  modal.style.display = 'flex';
  setTimeout(() => document.getElementById('edit-input').focus(), 50);
}

// ── SAVE EDIT ─────────────────────────────────
function saveEdit() {
  const newName = document.getElementById('edit-input').value.trim();
  if (!newName) return;

  const row = rows.find(r => r.id === editingId);
  if (row) {
    row.name   = newName;
    row.status = modalStatus;   // save the status chosen inside the modal
    saveToStorage();
    renderTable();
  }

  closeModal();
}

// ── CLOSE MODAL ───────────────────────────────
function closeModal() {
  const modal = document.getElementById('edit-modal');
  modal.style.display = 'none';
  modal.hidden = true;
  editingId = null;
}

// ── EXPORT TO EXCEL (.xlsx) ───────────────────
function exportExcel() {
  if (rows.length === 0) {
    alert('Nothing to export. Add some rows first!');
    return;
  }

  const sheetData = [['Serial No', 'Name', 'Status']];
  rows.forEach(function (row, index) {
    sheetData.push([index + 1, row.name, row.status]);
  });

  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  ws['!cols'] = [{ wch: 10 }, { wch: 30 }, { wch: 10 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Distribution List');

  const today = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, 'meat-distribution-' + today + '.xlsx');
}

// ── IMPORT FROM EXCEL (.xlsx) ─────────────────
function importExcel(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data     = new Uint8Array(e.target.result);
      const wb       = XLSX.read(data, { type: 'array' });
      const ws       = wb.Sheets[wb.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(ws, { defval: '' });

      if (jsonData.length === 0) { alert('The Excel file appears to be empty.'); return; }

      let added = 0, skipped = 0;
      jsonData.forEach(function (rowObj) {
        const nameRaw   = rowObj['Name'] || rowObj['name'] || rowObj['NAME'] || '';
        const statusRaw = rowObj['Status'] || rowObj['status'] || rowObj['STATUS'] || 'NO';
        const name      = String(nameRaw).trim();
        const status    = String(statusRaw).trim().toUpperCase() === 'YES' ? 'YES' : 'NO';
        if (!name) { skipped++; return; }
        rows.push({ id: generateId(), name: name, status: status, locked: true });
        added++;
      });

      saveToStorage();
      renderTable();
      event.target.value = '';
      alert('✅ Imported ' + added + ' row(s).' + (skipped ? '\n⚠ ' + skipped + ' skipped (empty name).' : ''));
    } catch (err) {
      alert('❌ Could not read the file.\n\nError: ' + err.message);
      event.target.value = '';
    }
  };
  reader.readAsArrayBuffer(file);
}

// ── CLEAR ALL DATA ────────────────────────────
function clearAll() {
  if (rows.length === 0) return;
  if (!confirm('Delete ALL records? This cannot be undone.')) return;
  rows = [];
  saveToStorage();
  renderTable();
}

// ── RENDER TABLE ──────────────────────────────
function renderTable() {
  const tbody      = document.getElementById('table-body');
  const emptyState = document.getElementById('empty-state');

  tbody.innerHTML = '';

  if (rows.length === 0) {
    document.getElementById('distribution-table').style.display = 'none';
    emptyState.classList.add('visible');
    updateStats();
    return;
  }

  document.getElementById('distribution-table').style.display = 'table';
  emptyState.classList.remove('visible');

  rows.forEach(function (row, index) {
    const isYes = row.status === 'YES';
    const tr    = document.createElement('tr');

    // Serial number
    const tdSerial = document.createElement('td');
    tdSerial.className   = 'cell-serial';
    tdSerial.textContent = index + 1;

    // Name
    const tdName = document.createElement('td');
    tdName.className   = 'cell-name';
    tdName.textContent = row.name;

    // Status cell:
    // • locked: false → clickable button (first-time set)
    // • locked: true  → plain badge (use Edit to change)
    const tdStatus = document.createElement('td');

    if (!row.locked) {
      // Unlocked: show as a clickable toggle button
      const btn = document.createElement('button');
      btn.className   = 'btn-status ' + (isYes ? 'yes' : 'no');
      btn.textContent = isYes ? '✓ YES' : '✗ NO';
      btn.title       = 'Click to confirm status (one-time)';
      btn.onclick     = function () { firstTimeToggle(row.id); };
      tdStatus.appendChild(btn);
    } else {
      // Locked: show as non-clickable badge
      const badge = document.createElement('span');
      badge.className   = 'status-badge ' + (isYes ? 'yes' : 'no');
      badge.textContent = isYes ? '✓ YES' : '✗ NO';
      badge.title       = 'Use Edit button to change';
      tdStatus.appendChild(badge);
    }

    // Actions
    const tdActions = document.createElement('td');
    tdActions.className = 'cell-actions';

    const btnEdit = document.createElement('button');
    btnEdit.className   = 'btn-edit';
    btnEdit.textContent = '✏ Edit';
    btnEdit.onclick     = function () { openEdit(row.id); };

    const btnDelete = document.createElement('button');
    btnDelete.className   = 'btn-delete';
    btnDelete.textContent = '✕ Del';
    btnDelete.onclick     = function () { deleteRow(row.id); };

    tdActions.appendChild(btnEdit);
    tdActions.appendChild(btnDelete);

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
  const total    = rows.length;
  const yesCount = rows.filter(r => r.status === 'YES').length;
  const noCount  = total - yesCount;

  document.getElementById('stat-total').textContent = total;
  document.getElementById('stat-yes').textContent   = yesCount;
  document.getElementById('stat-no').textContent    = noCount;
}
