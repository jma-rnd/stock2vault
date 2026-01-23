// -----------------------------
// Stock: Group Desc + Date filtering
// -----------------------------
function computeCounts(rows, idx) {
  const counts = new Map();
  if (idx === -1) return counts;
  for (const r of rows) {
    const key = String((r[idx] ?? '')).trim();
    if (!key) continue;
    counts.set(key, (counts.get(key) || 0) + 1);
  }
  const sorted = Array.from(counts.entries()).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
  return new Map(sorted);
}

function parseLooseDate(v) {
  if (v == null) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;

  // Excel serial
  if (typeof v === 'number' && isFinite(v)) {
    try {
      const d = XLSX.SSF.parse_date_code(v);
      if (d && d.y && d.m && d.d) {
        // Use UTC to avoid timezone drift
        return new Date(Date.UTC(d.y, d.m - 1, d.d));
      }
    } catch (_) {}
  }

  const s = String(v).trim();
  if (!s) return null;

  // YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const [y, m, d] = s.split('-').map(n => Number(n));
    return new Date(Date.UTC(y, m - 1, d));
  }

  // DD/MM/YYYY or D/M/YYYY
  const m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m1) {
    const dd = Number(m1[1]);
    const mm = Number(m1[2]);
    const yy = Number(m1[3]);
    return new Date(Date.UTC(yy, mm - 1, dd));
  }

  // Fallback: Date.parse
  const t = Date.parse(s);
  if (!isNaN(t)) return new Date(t);

  return null;
}

function setDateFilterMode(mode) {
  state.stock.dateFilter.mode = mode;
  // exclusive checkboxes
  el.f6m.checked = (mode === '6m');
  el.f12m.checked = (mode === '12m');
  el.fcustom.checked = (mode === 'custom');

  const enableCustom = (mode === 'custom');
  el.fStart.disabled = !enableCustom;
  el.fEnd.disabled = !enableCustom;

  updateDateFilterRange();
  scheduleAuditRun();
}

function updateDateFilterRange() {
  const mode = state.stock.dateFilter.mode;
  const now = new Date();
  const utcToday = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));

  if (mode === 'none') {
    state.stock.dateFilter.start = null;
    state.stock.dateFilter.end = null;
    el.fMeta.textContent = 'No date filtering applied.';
    return;
  }

  if (mode === '6m' || mode === '12m') {
    const months = (mode === '6m') ? 6 : 12;
    const start = new Date(Date.UTC(utcToday.getUTCFullYear(), utcToday.getUTCMonth() - months, utcToday.getUTCDate()));
    state.stock.dateFilter.start = start;
    state.stock.dateFilter.end = utcToday;
    el.fStart.value = start.toISOString().slice(0,10);
    el.fEnd.value = utcToday.toISOString().slice(0,10);
    el.fMeta.textContent = `Filtering to last ${months} months (inclusive).`;
    return;
  }

  if (mode === 'custom') {
    const start = el.fStart.value ? parseLooseDate(el.fStart.value) : null;
    const end = el.fEnd.value ? parseLooseDate(el.fEnd.value) : null;
    state.stock.dateFilter.start = start;
    state.stock.dateFilter.end = end;
    el.fMeta.textContent = 'Custom date filter applied (inclusive).';
    return;
  }
}

function rowPassesDateFilter(row) {
  const mode = state.stock.dateFilter.mode;
  if (mode === 'none') return true;
  const idx = state.audit.stockMoveDateIdx;
  if (idx === -1) return true; // no date column -> don't filter
  const d = parseLooseDate(row[idx]);
  if (!d) return false; // if filtering, rows without a date do not pass

  const start = state.stock.dateFilter.start;
  const end = state.stock.dateFilter.end;
  if (start && d < start) return false;
  if (end && d > end) return false;
  return true;
}

function rebuildGroupSelectionDefaultAll() {
  state.stock.groupSelected = new Map();
  for (const [g] of state.stock.groupCounts.entries()) {
    state.stock.groupSelected.set(g, true);
  }
}

function applyDefaultGroupSelection() {
  const wanted = new Set(DEFAULT_GROUP_DESCS.map(x => normalizeKey(x)));
  let selected = 0;
  for (const g of state.stock.groupSelected.keys()) {
    const isOn = wanted.has(normalizeKey(g));
    state.stock.groupSelected.set(g, isOn);
    if (isOn) selected++;
  }
  return { selected };
}

function renderGroupList() {
  const q = normalizeKey(el.groupSearch.value || '');
  const list = el.groupList;
  clearNode(list);

  const head = document.createElement('div');
  head.className = 'listHeader';
  head.innerHTML = '<div></div><div>Group Desc</div><div style="text-align:right;">Rows</div>';
  list.appendChild(head);

  if (!state.stock.fileName || !state.stock.groupCounts.size) {
    const row = document.createElement('div');
    row.className = 'row';
    row.innerHTML = '<div></div><div class="muted">Load a stock file to see Group Desc values.</div><div></div>';
    list.appendChild(row);
    return;
  }

  let shown = 0;
  for (const [g, cnt] of state.stock.groupCounts.entries()) {
    if (q && !normalizeKey(g).includes(q)) continue;
    shown++;
    const row = document.createElement('div');
    row.className = 'row';

    const cbWrap = document.createElement('div');
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = state.stock.groupSelected.get(g) !== false;
    cb.addEventListener('change', () => {
      state.stock.groupSelected.set(g, cb.checked);
      scheduleAuditRun();
    });
    cbWrap.appendChild(cb);

    const name = document.createElement('div');
    name.className = 'name';
    name.textContent = g;

    const c = document.createElement('div');
    c.className = 'count';
    c.textContent = String(cnt);

    row.appendChild(cbWrap);
    row.appendChild(name);
    row.appendChild(c);
    list.appendChild(row);
  }

  if (shown === 0) {
    const row = document.createElement('div');
    row.className = 'row';
    row.innerHTML = '<div></div><div class="muted">No Group Desc values match the current search.</div><div></div>';
    list.appendChild(row);
  }
}

function getFilteredStockRows() {
  const rows = state.stock.rows || [];
  const gIdx = state.audit.stockGroupIdx;
  const out = [];
  for (const r of rows) {
    // Date filter first
    if (!rowPassesDateFilter(r)) continue;

    // Group filter
    if (gIdx !== -1 && state.stock.groupSelected && state.stock.groupSelected.size) {
      const g = safeCell(r, gIdx);
      if (!g) continue;
      if (state.stock.groupSelected.get(g) === false) continue;
    }
    out.push(r);
  }
  return out;
}
