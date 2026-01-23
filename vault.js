// -----------------------------
// Vault: indexes & filetype summary
// -----------------------------
function computeGroupCountsGeneric(rows, idx) {
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

function renderVaultFiletypeSummary() {
  const list = el.vaultFiletypeList;
  if (!list) return;

  clearNode(list);

  const header = document.createElement('div');
  header.className = 'listHeader';
  header.style.gridTemplateColumns = '1fr 110px';
  header.innerHTML = '<div>Filetype</div><div style="text-align:right;">Rows</div>';
  list.appendChild(header);

  if (!state.vault.fileName || !state.vault.headers.length) {
    const row = document.createElement('div');
    row.className = 'row';
    row.style.gridTemplateColumns = '1fr 110px';
    row.innerHTML = '<div class="muted">Load a vault file to see Filetype values.</div><div></div>';
    list.appendChild(row);
    return;
  }

  const ftIdx = findColumnIndex(state.vault.headers, 'Filetype');
  if (ftIdx === -1) {
    const row = document.createElement('div');
    row.className = 'row';
    row.style.gridTemplateColumns = '1fr 110px';
    row.innerHTML = '<div class="muted">Column "Filetype" not found — summary unavailable.</div><div></div>';
    list.appendChild(row);
    return;
  }

  const counts = computeGroupCountsGeneric(state.vault.rows, ftIdx);
  if (!counts.size) {
    const row = document.createElement('div');
    row.className = 'row';
    row.style.gridTemplateColumns = '1fr 110px';
    row.innerHTML = '<div class="muted">No Filetype values found.</div><div></div>';
    list.appendChild(row);
    return;
  }

  for (const [name, count] of counts.entries()) {
    const row = document.createElement('div');
    row.className = 'row';
    row.style.gridTemplateColumns = '1fr 110px';

    const label = document.createElement('div');
    label.className = 'name';
    label.textContent = name;

    const c = document.createElement('div');
    c.className = 'count';
    c.textContent = String(count);

    row.appendChild(label);
    row.appendChild(c);
    list.appendChild(row);
  }
}

function updateAuditMappings() {
  state.audit.stockMatchIdx = findColumnIndex(state.stock.headers, state.audit.stockMatchCol || '');
  state.audit.stockGroupIdx = findColumnIndex(state.stock.headers, 'Group Desc');
  state.audit.stockDescIdx = findColumnIndex(state.stock.headers, 'Part Desc');
  state.audit.stockMoveDateIdx = findColumnIndex(state.stock.headers, 'Last Movement Date');

  state.audit.vaultMatchIdx = findColumnIndex(state.vault.headers, state.audit.vaultMatchCol || '');
  state.audit.vaultTypeIdx = findColumnIndex(state.vault.headers, 'Filetype');
  state.audit.vaultStateIdx = findColumnIndex(state.vault.headers, 'State');
  state.audit.vaultTitleIdx = findColumnIndex(state.vault.headers, 'Title');
}

// Normalize any path-ish string to a base filename (no directories)
function stripPath(nameVal) {
  const s = String(nameVal || '').trim();
  const unified = s.split('\\\\').join('/');
  const parts = unified.split('/');
  return parts.length ? parts[parts.length - 1] : unified;
}

function baseNameFromAny(nameVal) {
  const just = stripPath(nameVal);
  const low = just.toLowerCase();
  const dot = low.lastIndexOf('.');
  return (dot > 0 ? low.slice(0, dot) : low);
}

function detectExt(filetypeVal, nameVal) {
  const t = String(filetypeVal || '').toLowerCase();
  const n = String(nameVal || '').toLowerCase();
  const exts = ['.idw', '.pdf', '.ipt', '.iam'];
  for (const e of exts) {
    if (t.includes(e)) return e;
    if (n.endsWith(e)) return e;
  }
  return '';
}

function buildVaultIndex() {
  state.audit.vaultIndex = new Map();
  if (!state.vault.rows.length) return;
  updateAuditMappings();

  const mIdx = state.audit.vaultMatchIdx;
  const tIdx = state.audit.vaultTypeIdx;
  const sIdx = state.audit.vaultStateIdx;

  for (const r of state.vault.rows) {
    const key = normalizeKey(safeCell(r, mIdx));
    if (!key) continue;
    const nameVal = safeCell(r, mIdx);
    const typeVal = safeCell(r, tIdx);
    const stateVal = safeCell(r, sIdx);
    const ext = detectExt(typeVal, nameVal);
    const base = baseNameFromAny(nameVal);
    const entry = { key, name: nameVal, base, ext, filetype: String(typeVal || ''), state: String(stateVal || ''), row: r };
    if (!state.audit.vaultIndex.has(key)) state.audit.vaultIndex.set(key, []);
    state.audit.vaultIndex.get(key).push(entry);
  }
}
