const CLIENT_ID = 'db8d33ea-a736-4152-be60-714f71fd7a0f';
const TENANT_ID = '5ed35cd3-7f0d-4691-b1f9-56d422d5b3ca';
// Derived from wherever the app is actually being served from, so the same
// code works unchanged on GitHub Pages, Azure Static Web Apps, localhost, etc.
// Whatever this resolves to at runtime must also be added to the "Redirect URIs"
// list on the Azure AD app registration (Entra ID > App registrations > this app
// > Authentication) — see README.md for details.
const REDIRECT_URI = window.location.origin + window.location.pathname;
const SCOPES = ['Files.ReadWrite', 'User.Read'];
const POUCHES_PER_UNIT = 225; // default fallback
function factor(item) { return item.pouchesPerUnit || POUCHES_PER_UNIT; }
const FOLDER_NAME = 'SGC Stock Tracker';
const FILE_NAME = 'inventory.json';
const GRAPH_FILE_URL = 'https://graph.microsoft.com/v1.0/me/drive/root:/' +
  encodeURIComponent(FOLDER_NAME) + '/' + encodeURIComponent(FILE_NAME) + ':/content';

const msalConfig = {
  auth: { clientId: CLIENT_ID, authority: 'https://login.microsoftonline.com/' + TENANT_ID, redirectUri: REDIRECT_URI },
  cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: false }
};
const msalInstance = new msal.PublicClientApplication(msalConfig);

const SAMPLES = [
  { id: 1, name: 'Sour gummy worms', current: 12, target: 100 },
  { id: 2, name: 'Peach rings', current: 85, target: 100 },
  { id: 3, name: 'Gummy bears', current: 45, target: 100 },
  { id: 4, name: 'Mango bites', current: 8, target: 200 },
  { id: 5, name: 'Watermelon slices', current: 150, target: 200 },
  { id: 6, name: 'Cola bottles', current: 3, target: 500 },
];

let items = [];
let log = [];
let search = '';
let editMode = {};
let pouchMode = false;

// Some take-outs were logged under an old/inconsistent product name before it
// was corrected. This maps those old names to the current correct name so the
// log (and its export) always display the right name, regardless of when the
// take-out was recorded.
const NAME_ALIASES = {
  'Mocca Shots Mint': 'Mint',
  'Mocca Shots Dutch': 'Dutch',
};
function normalizeItemName(name) { return NAME_ALIASES[name] || name; }

function pct(item) { return item.target > 0 ? Math.round((item.current / item.target) * 100) : 0; }
function status(p) {
  if (p <= 30) return { key: 'danger', label: 'Critical' };
  if (p <= 70) return { key: 'warn', label: 'Low' };
  return { key: 'good', label: 'Good' };
}
const styleMap = {
  danger: { bg: 'var(--danger-bg)', bar: 'var(--danger-bar)', text: 'var(--danger-text)', border: 'var(--danger-border)' },
  warn:   { bg: 'var(--warn-bg)',   bar: 'var(--warn-bar)',   text: 'var(--warn-text)',   border: 'var(--warn-border)' },
  good:   { bg: 'var(--good-bg)',   bar: 'var(--good-bar)',   text: 'var(--good-text)',   border: 'var(--good-border)' },
};
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}
function fmt(n) { return Number(n).toLocaleString(); }
function displayVal(units, f) { return pouchMode ? fmt(units * f) : fmt(units); }
function displayUnit() { return pouchMode ? 'pouches' : 'units'; }

function setStatus(text, kind) {
  document.getElementById('statusText').textContent = text;
  const dot = document.getElementById('statusDot');
  dot.className = 'status-dot ' + (kind === 'ok' ? 'dot-ok' : kind === 'bad' ? 'dot-bad' : 'dot-mid');
}

async function getAccessToken() {
  const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  if (!account) throw new Error('Not signed in');
  try {
    const r = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account });
    return r.accessToken;
  } catch {
    const r = await msalInstance.acquireTokenPopup({ scopes: SCOPES });
    msalInstance.setActiveAccount(r.account);
    return r.accessToken;
  }
}

async function loadFromOneDrive() {
  document.getElementById('loadingScreen').style.display = 'block';
  document.getElementById('appShell').style.display = 'none';
  try {
    const token = await getAccessToken();
    const res = await fetch(GRAPH_FILE_URL, { headers: { Authorization: 'Bearer ' + token } });
    if (res.status === 404) {
      items = SAMPLES; log = [];
      await saveToOneDrive();
    } else if (res.ok) {
      const data = await res.json();
      // Handle both old format (plain array) and new format ({ items, log })
      if (Array.isArray(data)) { items = data; log = []; }
      else { items = data.items || []; log = data.log || []; }
    } else { throw new Error('Graph error ' + res.status); }
    setStatus('Synced', 'ok');
  } catch (e) {
    console.error(e); items = []; log = [];
    setStatus('Could not load from OneDrive', 'bad');
  }
  document.getElementById('loadingScreen').style.display = 'none';
  document.getElementById('appShell').style.display = 'block';
  render();
}

async function saveToOneDrive() {
  setStatus('Saving...', 'mid');
  try {
    const token = await getAccessToken();
    const res = await fetch(GRAPH_FILE_URL, {
      method: 'PUT',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ items, log }, null, 2)
    });
    if (!res.ok) throw new Error('Save failed ' + res.status);
    const now = new Date();
    setStatus('Synced ' + now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }), 'ok');
  } catch (e) { console.error(e); setStatus('Save failed — check connection', 'bad'); }
}

function persist() { render(); saveToOneDrive(); }

const RING_R = 26;
const RING_C = 2 * Math.PI * RING_R;

let sortMode = 'pct-asc'; // 'pct-asc' | 'pct-desc' | 'units' | 'pouches' | 'alpha'
let groupByCategory = false;

function sortItemsList(arr) {
  const out = [...arr];
  if (sortMode === 'pct-desc') out.sort((a, b) => b._pct - a._pct);
  else if (sortMode === 'units') out.sort((a, b) => b.current - a.current);
  else if (sortMode === 'pouches') out.sort((a, b) => (b.current * factor(b)) - (a.current * factor(a)));
  else if (sortMode === 'alpha') out.sort((a, b) => a.name.localeCompare(b.name));
  else out.sort((a, b) => a._pct - b._pct); // 'pct-asc' default
  return out;
}

function groupItemsList(arr) {
  if (!groupByCategory) return [{ category: null, items: arr }];
  const map = new Map();
  arr.forEach(item => {
    const cat = (item.category || '').trim() || 'Uncategorized';
    if (!map.has(cat)) map.set(cat, []);
    map.get(cat).push(item);
  });
  const cats = Array.from(map.keys()).sort((a, b) => {
    if (a === 'Uncategorized') return 1;
    if (b === 'Uncategorized') return -1;
    return a.localeCompare(b);
  });
  return cats.map(cat => ({ category: cat, items: map.get(cat) }));
}

function renderCard(item) {
      const st = status(item._pct);
      const sm = styleMap[st.key];
      const clamp = Math.min(item._pct, 100);
      const dashLen = +(RING_C * clamp / 100).toFixed(1);
      const gapLen = +(RING_C - dashLen).toFixed(1);
      const mode = editMode[item.id];
      const badgeClass = pouchMode ? 'unit-badge pouch-badge' : 'unit-badge';
      const badgeLabel = pouchMode ? 'pouches' : 'units';

      let actionsHtml = '';
      if (mode === 'count') {
        actionsHtml = `
          <div class="crate-row">
            <button class="crate-step step-btn" data-id="${item.id}" data-delta="-3">-3</button>
            <button class="crate-step step-btn" data-id="${item.id}" data-delta="-1">-1</button>
            <input type="number" min="0" class="crate-input" id="countInput-${item.id}" value="${item.current}" placeholder="Count">
            <button class="crate-step step-btn" data-id="${item.id}" data-delta="1">+1</button>
            <button class="crate-step step-btn" data-id="${item.id}" data-delta="3">+3</button>
          </div>
          <p class="crate-hint">&plusmn;1 unit &middot; &plusmn;3 units (1 crate) &middot; or type any number</p>
          <div class="save-cancel-row">
            <button class="cancel-mode-btn" data-id="${item.id}">Cancel</button>
            <button class="primary save-count-btn" data-id="${item.id}">Save</button>
          </div>`;
      } else if (mode === 'takeout') {
        actionsHtml = `
          <div class="takeout-form">
            <label>Lot number</label>
            <input id="lotInput-${item.id}" placeholder="e.g. LOT-2024-001" autocomplete="off">
            <div class="row2">
              <div>
                <label>Units to remove</label>
                <input type="number" min="1" id="takeQty-${item.id}" value="1">
              </div>
              <div>
                <label>= Pouches</label>
                <input type="number" id="takePouches-${item.id}" placeholder="${factor(item)}" readonly style="color:var(--text-muted);">
              </div>
            </div>
          </div>
          <div class="save-cancel-row" style="margin-top:8px;">
            <button class="cancel-mode-btn" data-id="${item.id}">Cancel</button>
            <button class="primary confirm-takeout-btn" data-id="${item.id}">Confirm take-out</button>
          </div>`;
      } else if (mode === 'edit') {
        actionsHtml = `
          <div style="margin-top:12px;">
            <input id="editName-${item.id}" value="${escapeHtml(item.name)}" placeholder="Item name" style="margin-bottom:6px;">
            <div class="row2" style="margin-top:0; margin-bottom:8px;">
              <input type="number" min="0" id="editCurrent-${item.id}" value="${item.current}" placeholder="Current qty">
              <input type="number" min="1" id="editTarget-${item.id}" value="${item.target}" placeholder="Target qty">
            </div>
            <div class="row2" style="margin-top:0; margin-bottom:8px;">
              <input type="number" min="1" id="editPouches-${item.id}" value="${item.pouchesPerUnit || POUCHES_PER_UNIT}" placeholder="Pouches per unit">
              <input id="editCategory-${item.id}" value="${escapeHtml(item.category || '')}" placeholder="Category (optional)">
            </div>
            <div class="card-actions">
              <button class="danger-text delete-btn" data-id="${item.id}" style="flex:0 0 auto;">Delete</button>
              <div style="flex:1;"></div>
              <button class="cancel-mode-btn" data-id="${item.id}" style="flex:0 0 auto;">Cancel</button>
              <button class="primary save-edit-btn" data-id="${item.id}" style="flex:0 0 auto;">Save</button>
            </div>
          </div>`;
      } else {
        actionsHtml = `
          <div class="card-actions">
            <button class="takeout-btn" data-id="${item.id}">Take out</button>
            <button class="count-btn" data-id="${item.id}">Update count</button>
            <button class="icon-btn edit-btn" data-id="${item.id}" aria-label="Edit">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>
            </button>
          </div>`;
      }

  return `
    <div class="card">
      <div class="card-body">
        <div class="ring-wrap">
          <svg viewBox="0 0 64 64" width="54" height="54">
            <circle cx="32" cy="32" r="${RING_R}" fill="none" stroke="var(--border-strong)" stroke-width="6"/>
            <circle class="ring-fill" cx="32" cy="32" r="${RING_R}" fill="none" stroke="${sm.bar}" stroke-width="6" stroke-linecap="round" transform="rotate(-90 32 32)" style="stroke-dasharray:${dashLen} ${gapLen};"/>
          </svg>
          <div class="ring-pct" style="color:${sm.text};">${item._pct}%</div>
        </div>
        <div class="card-info">
          <div class="card-name">${escapeHtml(item.name)} <span class="${badgeClass}">${badgeLabel}</span></div>
          <div class="card-qty">${displayVal(item.current, factor(item))} / ${displayVal(item.target, factor(item))}</div>
          <span class="status-pill" style="background:${sm.bg}; color:${sm.text}; border-color:${sm.border};">${st.label}</span>
          ${item.category ? `<div class="card-cat">${escapeHtml(item.category)}</div>` : ''}
        </div>
      </div>
      ${actionsHtml}
    </div>`;
}

function render() {
  const filtered = items.filter(i => i.name.toLowerCase().includes(search.toLowerCase()));
  const withPct = filtered.map(i => ({ ...i, _pct: pct(i) }));
  const sorted = sortItemsList(withPct);
  const groups = groupItemsList(sorted);

  document.getElementById('sumCritical').textContent = items.filter(i => pct(i) <= 30).length;
  document.getElementById('sumLow').textContent = items.filter(i => { const p = pct(i); return p > 30 && p <= 70; }).length;
  document.getElementById('sumGood').textContent = items.filter(i => pct(i) > 70).length;

  const pouchBtn = document.getElementById('pouchToggleBtn');
  if (pouchMode) { pouchBtn.textContent = 'View in units'; pouchBtn.classList.add('active-toggle'); }
  else { pouchBtn.textContent = 'View in pouches'; pouchBtn.classList.remove('active-toggle'); }

  document.getElementById('sortSelect').value = sortMode;
  const groupBtn = document.getElementById('groupToggleBtn');
  if (groupByCategory) { groupBtn.classList.add('active-toggle'); }
  else { groupBtn.classList.remove('active-toggle'); }

  const list = document.getElementById('list');
  if (sorted.length === 0) {
    list.innerHTML = '<div class="empty">' + (items.length === 0 ? 'No items yet. Add your first item.' : 'No items match your search.') + '</div>';
  } else {
    list.innerHTML = groups.map(g => {
      const cardsHtml = g.items.map(renderCard).join('');
      if (g.category === null) return `<div class="card-grid">${cardsHtml}</div>`;
      const count = g.items.length;
      return `
        <section>
          <h3 class="item-group-title">${escapeHtml(g.category)} <span class="item-group-count">${count} item${count !== 1 ? 's' : ''}</span></h3>
          <div class="card-grid">${cardsHtml}</div>
        </section>`;
    }).join('');
  }

  document.getElementById('footerCount').textContent = items.length > 0
    ? items.length + ' item' + (items.length !== 1 ? 's' : '') + ' tracked' : '';

  attachCardListeners();
}

function attachCardListeners() {
  document.querySelectorAll('.count-btn').forEach(b => b.onclick = () => { editMode[b.dataset.id] = 'count'; render(); });
  document.querySelectorAll('.takeout-btn').forEach(b => b.onclick = () => { editMode[b.dataset.id] = 'takeout'; render(); updatePouchPreview(b.dataset.id); });
  document.querySelectorAll('.edit-btn').forEach(b => b.onclick = () => { editMode[b.dataset.id] = 'edit'; render(); });
  document.querySelectorAll('.cancel-mode-btn').forEach(b => b.onclick = () => { delete editMode[b.dataset.id]; render(); });
  document.querySelectorAll('.step-btn').forEach(b => b.onclick = () => {
    const inp = document.getElementById('countInput-' + b.dataset.id);
    inp.value = Math.max(0, (parseInt(inp.value) || 0) + parseInt(b.dataset.delta));
  });
  document.querySelectorAll('.delete-btn').forEach(b => b.onclick = () => {
    items = items.filter(i => String(i.id) !== b.dataset.id);
    delete editMode[b.dataset.id]; persist();
  });
  document.querySelectorAll('.save-count-btn').forEach(b => b.onclick = () => {
    const id = b.dataset.id;
    const val = parseInt(document.getElementById('countInput-' + id).value);
    if (!isNaN(val) && val >= 0) { const it = items.find(i => String(i.id) === id); if (it) it.current = val; }
    delete editMode[id]; persist();
  });
  document.querySelectorAll('.save-edit-btn').forEach(b => b.onclick = () => {
    const id = b.dataset.id;
    const name = document.getElementById('editName-' + id).value.trim();
    const cur = parseInt(document.getElementById('editCurrent-' + id).value);
    const tgt = parseInt(document.getElementById('editTarget-' + id).value);
    if (name && tgt > 0) {
      const it = items.find(i => String(i.id) === id);
      if (it) {
        it.name = name; it.current = isNaN(cur) ? 0 : cur; it.target = tgt;
        it.pouchesPerUnit = parseInt(document.getElementById('editPouches-' + id).value) || POUCHES_PER_UNIT;
        it.category = document.getElementById('editCategory-' + id).value.trim();
      }
      delete editMode[id]; persist();
    }
  });

  // Take-out: live pouch preview
  document.querySelectorAll('[id^="takeQty-"]').forEach(inp => {
    inp.oninput = () => updatePouchPreview(inp.id.replace('takeQty-', ''));
  });

  document.querySelectorAll('.confirm-takeout-btn').forEach(b => b.onclick = () => {
    const id = b.dataset.id;
    const lot = document.getElementById('lotInput-' + id).value.trim();
    const qty = parseInt(document.getElementById('takeQty-' + id).value) || 1;
    if (!lot) { document.getElementById('lotInput-' + id).focus(); return; }
    const it = items.find(i => String(i.id) === id);
    if (!it) return;
    const removed = Math.min(qty, it.current);
    it.current = Math.max(0, it.current - qty);
    log.unshift({
      id: Date.now(),
      itemId: it.id,
      itemName: it.name,
      lotNumber: lot,
      unitsRemoved: removed,
      pouchesRemoved: removed * factor(it),
      date: new Date().toISOString()
    });
    delete editMode[id];
    persist();
  });

  document.querySelectorAll('[id^="countInput-"]').forEach(inp => {
    inp.onkeydown = e => { if (e.key === 'Enter') document.querySelector('.save-count-btn[data-id="' + inp.id.split('-')[1] + '"]').click(); };
  });
}

function updatePouchPreview(id) {
  const qtyInp = document.getElementById('takeQty-' + id);
  const pouchInp = document.getElementById('takePouches-' + id);
  const it = items.find(i => String(i.id) === String(id));
  if (qtyInp && pouchInp && it) pouchInp.value = ((parseInt(qtyInp.value) || 0) * factor(it)).toLocaleString();
}

function renderLog() {
  const el = document.getElementById('logList');
  if (log.length === 0) { el.innerHTML = '<div class="log-empty">No take-outs recorded yet.</div>'; return; }
  el.innerHTML = log.map(e => {
    const d = new Date(e.date);
    const dateStr = d.toLocaleDateString([], { month: 'short', day: 'numeric', year: 'numeric' }) +
      ' ' + d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    return `
      <div class="log-entry">
        <div class="log-cell-name">
          <div class="log-item-name">${escapeHtml(normalizeItemName(e.itemName))}</div>
          <div class="log-lot">Lot <span class="mono">${escapeHtml(e.lotNumber)}</span></div>
        </div>
        <div class="log-cell-qty">
          <div class="log-qty-num">-${fmt(e.unitsRemoved)} units</div>
          <div class="log-qty-label">${fmt(e.pouchesRemoved)} pouches</div>
        </div>
        <div class="log-cell-date">${dateStr}</div>
      </div>`;
  }).join('');
}

// Formats a date like "27-Feb" (day, abbreviated month — no year/time), used in the CSV export.
function formatShortDate(dateInput) {
  const d = new Date(dateInput);
  return d.getDate() + '-' + d.toLocaleString('en-US', { month: 'short' });
}

function exportCSV() {
  // Keep only the first (earliest) take-out recorded for each item + lot number
  // pairing — the same lot number used on two different products still counts
  // as two separate lots.
  const firstByKey = new Map();
  log.forEach(e => {
    const key = normalizeItemName(e.itemName) + '||' + e.lotNumber;
    const existing = firstByKey.get(key);
    if (!existing || new Date(e.date) < new Date(existing.date)) {
      firstByKey.set(key, e);
    }
  });
  const uniqueEntries = Array.from(firstByKey.values())
    .sort((a, b) => new Date(b.date) - new Date(a.date));

  const header = ['Date', 'Item', 'Lot Number'];
  const rows = uniqueEntries.map(e => [
    formatShortDate(e.date), normalizeItemName(e.itemName), e.lotNumber
  ].map(v => '"' + String(v).replace(/"/g, '""') + '"'));
  const csv = [header, ...rows].map(r => r.join(',')).join('\n');
  const blob = new Blob([csv], { type: 'text/csv' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'sgc-takeout-log-' + new Date().toISOString().slice(0,10) + '.csv';
  a.click();
}

function setActiveNav(target) {
  document.querySelectorAll('[data-nav]').forEach(b => {
    const isMatch = b.dataset.nav === target;
    b.classList.toggle('nav-active', isMatch && b.classList.contains('nav-link'));
    b.classList.toggle('tabbar-active', isMatch && b.classList.contains('tabbar-btn'));
  });
}

function showLogScreen() {
  document.getElementById('appShell').style.display = 'none';
  document.getElementById('logScreen').style.display = 'block';
  setActiveNav('log');
  renderLog();
}
function showInventoryScreen() {
  document.getElementById('logScreen').style.display = 'none';
  document.getElementById('appShell').style.display = 'block';
  setActiveNav('stock');
}

function exportInventoryExcel() {
  const sorted = [...items].sort((a, b) => a.name.localeCompare(b.name));

  const rows = sorted.map(item => ({
    'Item':             item.name,
    'Current (pouches)': item.current * factor(item),
  }));

  const ws = XLSX.utils.json_to_sheet(rows);

  ws['!cols'] = [
    { wch: 28 }, // Item
    { wch: 18 }, // Current pouches
  ];

  ws['!autofilter'] = { ref: 'A1:B' + (rows.length + 1) };
  ws['!tables'] = [{
    name: 'StockInventory',
    ref:  'A1:B' + (rows.length + 1),
    headerRow: true,
    totalsRow: false,
  }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Stock Levels');

  const now = new Date();
  const month = now.toLocaleString('en-US', { month: 'long', year: 'numeric' }).replace(' ', '-');
  XLSX.writeFile(wb, 'sgc-inventory-' + month + '.xlsx');
}

document.getElementById('exportExcelBtn').addEventListener('click', exportInventoryExcel);
document.getElementById('pouchToggleBtn').addEventListener('click', () => { pouchMode = !pouchMode; render(); });
document.querySelectorAll('[data-nav="stock"]').forEach(b => b.addEventListener('click', showInventoryScreen));
document.querySelectorAll('[data-nav="log"]').forEach(b => b.addEventListener('click', showLogScreen));
document.getElementById('backFromLog').addEventListener('click', showInventoryScreen);
document.getElementById('exportCsvBtn').addEventListener('click', exportCSV);
document.getElementById('searchInput').addEventListener('input', e => { search = e.target.value; render(); });
document.getElementById('addBtn').addEventListener('click', () => {
  document.getElementById('addPanel').style.display = 'block';
  document.getElementById('newName').focus();
});
document.getElementById('cancelAdd').addEventListener('click', () => {
  document.getElementById('addPanel').style.display = 'none';
  ['newName','newCurrent','newTarget','newPouches','newCategory'].forEach(id => document.getElementById(id).value = '');
});
document.getElementById('confirmAdd').addEventListener('click', () => {
  const name = document.getElementById('newName').value.trim();
  const current = parseInt(document.getElementById('newCurrent').value) || 0;
  const target = parseInt(document.getElementById('newTarget').value) || 0;
  const ppu = parseInt(document.getElementById('newPouches').value) || POUCHES_PER_UNIT;
  const category = document.getElementById('newCategory').value.trim();
  if (!name || target <= 0) return;
  items.push({ id: Date.now(), name, current, target, pouchesPerUnit: ppu, category });
  document.getElementById('cancelAdd').click();
  persist();
});
document.getElementById('sortSelect').addEventListener('change', e => { sortMode = e.target.value; render(); });
document.getElementById('groupToggleBtn').addEventListener('click', () => { groupByCategory = !groupByCategory; render(); });
document.getElementById('signOutBtn').addEventListener('click', () => msalInstance.logoutPopup().catch(() => {}));

async function onSignedIn() {
  document.getElementById('signinScreen').style.display = 'none';
  document.getElementById('authBar').classList.remove('hidden-until-auth');
  document.getElementById('sidebar').classList.remove('hidden-until-auth');
  document.getElementById('tabbar').classList.remove('hidden-until-auth');
  await loadFromOneDrive();
}
document.getElementById('signInBtn').addEventListener('click', async () => {
  document.getElementById('signinError').textContent = '';
  try {
    const r = await msalInstance.loginPopup({ scopes: SCOPES });
    msalInstance.setActiveAccount(r.account);
    await onSignedIn();
  } catch (e) {
    console.error(e);
    document.getElementById('signinError').textContent = 'Sign-in failed. Please try again.';
  }
});

(function init() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) { msalInstance.setActiveAccount(accounts[0]); onSignedIn(); }
})();
