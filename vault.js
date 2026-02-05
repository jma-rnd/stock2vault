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
  state.audit.vaultPartNumberIdx = findColumnIndex(state.vault.headers, 'Part Number');
  state.audit.vaultNameIdx = findColumnIndex(state.vault.headers, 'Name');
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

function extractDrawingNumber(nameVal) {
  const s = String(nameVal || '').toUpperCase();
  const m = s.match(/\b[A-Z]{1,3}\d{4}\b/);
  return m ? m[0] : '';
}

function parsePdfNameInfo(nameVal) {
  const raw = stripPath(nameVal);
  if (!raw) return { drawingNumber: '', partNumber: '' };

  const upper = raw.toUpperCase();
  const noExt = upper.replace(/\.[A-Z0-9]+$/, '');
  const cleaned = noExt.replace(/\bREV(?:ISION)?\s*\d+\b/g, ' ');

  const drawingNumber = extractDrawingNumber(cleaned);
  const parts = cleaned.match(/[A-Z0-9]{6,}/g) || [];
  const partNumber = parts.length ? parts[parts.length - 1] : '';

  return { drawingNumber, partNumber };
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

function escapeRegexLiteral(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function hasWildcardToken(code) {
  return /\([^)]+\)/.test(code);
}

function buildWildcardPattern(code) {
  const raw = String(code || '').trim();
  if (!raw) return null;

  const tokenRe = /\([^)]*\)/g;
  let lastIdx = 0;
  const parts = [];
  let kind = '';
  let m;
  while ((m = tokenRe.exec(raw)) !== null) {
    const before = raw.slice(lastIdx, m.index);
    if (before) parts.push(escapeRegexLiteral(before.toUpperCase()));

    const token = m[0].slice(1, -1).trim().toUpperCase();
    const isLast = (m.index + m[0].length === raw.length);
    if (token === 'LENGTH') {
      parts.push('\\d{2,4}');
      kind = kind || 'length';
    } else if (token === 'G') {
      parts.push('G?');
      if (!kind) kind = 'galv';
    } else {
      const max = isLast ? 5 : 10;
      parts.push(`[A-Z]{0,${max}}`);
      if (!kind) kind = 'letters';
    }
    lastIdx = m.index + m[0].length;
  }
  const tail = raw.slice(lastIdx);
  if (tail) parts.push(escapeRegexLiteral(tail.toUpperCase()));

  if (!parts.length) return null;
  const pattern = '^' + parts.join('') + '$';
  const wildcardCount = (raw.match(tokenRe) || []).length;
  return { pattern, wildcardCount, kind };
}

function renderWildcardUI() {
  if (!el.wildcardSummary || !el.wildcardPatterns || !el.wildcardMatches) return;

  const patterns = state.audit.vaultPatternIndex || [];
  el.wildcardSummary.textContent = patterns.length
    ? `${patterns.length} wildcard pattern${patterns.length === 1 ? '' : 's'} loaded.`
    : 'No wildcard patterns loaded.';

  clearNode(el.wildcardPatterns);
  if (!patterns.length) {
    const span = document.createElement('div');
    span.className = 'hint muted';
    span.textContent = 'No wildcard patterns detected in the Vault file.';
    el.wildcardPatterns.appendChild(span);
  } else {
    for (const pat of patterns) {
      const chip = document.createElement('div');
      chip.className = 'chip small';
      const name = document.createElement('span');
      name.textContent = pat.template;
      const mono = document.createElement('span');
      mono.className = 'mono';
      mono.textContent = pat.pattern;
      chip.appendChild(name);
      chip.appendChild(mono);
      el.wildcardPatterns.appendChild(chip);
    }
  }

  renderWildcardMatches('');
}

function renderWildcardMatches(input) {
  if (!el.wildcardMatches) return;
  clearNode(el.wildcardMatches);

  const code = String(input || '').trim();
  if (!code) {
    const span = document.createElement('div');
    span.className = 'hint muted';
    span.textContent = 'Enter a Part Code to test matching.';
    el.wildcardMatches.appendChild(span);
    return;
  }

  const matches = findVaultMatches(code);
  if (!matches.length) {
    const span = document.createElement('div');
    span.className = 'hint muted';
    span.textContent = 'No matches for this Part Code.';
    el.wildcardMatches.appendChild(span);
    return;
  }

  for (const m of matches) {
    const chip = document.createElement('div');
    chip.className = 'chip small';
    const name = document.createElement('span');
    name.textContent = m.name || m.key;
    const meta = document.createElement('span');
    meta.className = 'mono';
    meta.textContent = [m.filetype, m.state].filter(Boolean).join(' • ');
    chip.appendChild(name);
    if (meta.textContent) chip.appendChild(meta);
    el.wildcardMatches.appendChild(chip);
  }
}

function buildVaultIndex() {
  state.audit.vaultIndex = new Map();
  state.audit.vaultPatternIndex = [];
  state.audit.pdfNameIndex = new Map();
  state.audit.idwDrawingNumbers = new Set();
  if (!state.vault.rows.length) return;
  updateAuditMappings();

  const mIdx = state.audit.vaultMatchIdx;
  const tIdx = state.audit.vaultTypeIdx;
  const sIdx = state.audit.vaultStateIdx;
  const nIdx = state.audit.vaultNameIdx;

  for (const r of state.vault.rows) {
    const key = normalizeKey(safeCell(r, mIdx));
    if (!key) continue;
    const nameVal = safeCell(r, mIdx);
    const nameColVal = nIdx !== -1 ? safeCell(r, nIdx) : '';
    const typeVal = safeCell(r, tIdx);
    const stateVal = safeCell(r, sIdx);
    const ext = detectExt(typeVal, nameColVal || nameVal);
    const displayName = nameVal || nameColVal;
    const base = baseNameFromAny(displayName);
    const entry = { key, name: displayName, base, ext, filetype: String(typeVal || ''), state: String(stateVal || ''), row: r };
    if (!state.audit.vaultIndex.has(key)) state.audit.vaultIndex.set(key, []);
    state.audit.vaultIndex.get(key).push(entry);

    if (hasWildcardToken(nameVal)) {
      const built = buildWildcardPattern(nameVal);
      if (built && built.pattern) {
        state.audit.vaultPatternIndex.push({
          regex: new RegExp(built.pattern),
          template: nameVal,
          pattern: built.pattern,
          entry,
          kind: built.kind || 'letters',
          wildcardCount: built.wildcardCount,
          literalLen: nameVal.length,
        });
        if (built.kind) entry.wildcardKind = built.kind;
      }
    }
  }

  if (nIdx !== -1) {
    for (const r of state.vault.rows) {
      const nameColVal = safeCell(r, nIdx);
      if (!nameColVal) continue;
      const typeVal = safeCell(r, tIdx);
      const ext = detectExt(typeVal, nameColVal);
      if (ext !== '.idw') continue;
      const drawing = extractDrawingNumber(nameColVal);
      if (drawing) state.audit.idwDrawingNumbers.add(drawing);
    }

    for (const r of state.vault.rows) {
      const nameColVal = safeCell(r, nIdx);
      if (!nameColVal) continue;
      const typeVal = safeCell(r, tIdx);
      const stateVal = safeCell(r, sIdx);
      const ext = detectExt(typeVal, nameColVal);
      if (ext !== '.pdf') continue;

      const info = parsePdfNameInfo(nameColVal);
      if (!info.partNumber) continue;
      if (info.drawingNumber && state.audit.idwDrawingNumbers.has(info.drawingNumber)) continue;

      const key = normalizeKey(info.partNumber);
      if (!key) continue;

      const nameVal = safeCell(r, mIdx);
      const displayName = nameVal || nameColVal;
      const base = baseNameFromAny(displayName);
      const entry = {
        key,
        name: displayName,
        base,
        ext,
        filetype: String(typeVal || ''),
        state: String(stateVal || ''),
        row: r,
        pdfNameMatch: true,
        drawingNumber: info.drawingNumber,
      };
      if (!state.audit.pdfNameIndex.has(key)) state.audit.pdfNameIndex.set(key, []);
      state.audit.pdfNameIndex.get(key).push(entry);
    }
  }

  renderWildcardUI();
}

function findVaultMatches(stockCodeRaw) {
  const exactKey = normalizeKey(stockCodeRaw);
  const exact = state.audit.vaultIndex.get(exactKey) || [];
  if (exact.length) return exact;

  const code = String(stockCodeRaw || '').trim().toUpperCase();
  if (!code) return [];
  if (!state.audit.vaultPatternIndex.length) {
    const pdfOnly = state.audit.pdfNameIndex.get(exactKey) || [];
    return pdfOnly;
  }

  const matches = [];
  for (const pat of state.audit.vaultPatternIndex) {
    if (pat.regex.test(code)) matches.push(pat);
  }
  if (!matches.length) {
    const pdfOnly = state.audit.pdfNameIndex.get(exactKey) || [];
    return pdfOnly;
  }

  matches.sort((a, b) => (a.wildcardCount - b.wildcardCount) || (b.literalLen - a.literalLen));

  const bestKey = matches[0].entry.key;
  const out = [];
  const seen = new Set();
  for (const m of matches) {
    if (m.entry.key !== bestKey) continue;
    const key = m.entry.key + '|' + m.entry.name;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(m.entry);
  }
  if (out.length) return out;

  const pdfOnly = state.audit.pdfNameIndex.get(exactKey) || [];
  if (pdfOnly.length) return pdfOnly;

  return [];
}
