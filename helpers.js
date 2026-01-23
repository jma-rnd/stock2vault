// -----------------------------
// Helpers
// -----------------------------
function clearNode(node) {
  while (node.firstChild) node.removeChild(node.firstChild);
}

function log(msg, level='info') {
  const ts = new Date().toLocaleTimeString();
  const prefix = level === 'error' ? '⛔' : level === 'warn' ? '⚠️' : 'ℹ️';
  el.log.textContent = `${prefix} [${ts}] ${msg}\n` + el.log.textContent;
  el.logStatus.textContent = level === 'error' ? 'Error' : level === 'warn' ? 'Warning' : 'OK';
  el.logStatus.className = 'pill ' + (level === 'error' ? 'error' : level === 'warn' ? 'warn' : 'ok');
}

function setStatus(which, text, kind='muted') {
  const node = which === 'stock' ? el.stockStatus : el.vaultStatus;
  node.textContent = text;
  node.className = 'pill ' + (kind === 'ok' ? 'ok' : kind === 'warn' ? 'warn' : kind === 'error' ? 'error' : '');
}

function setAuditStatus(text, kind='muted') {
  el.auditStatus.textContent = text;
  el.auditStatus.className = 'pill ' + (kind === 'ok' ? 'ok' : kind === 'warn' ? 'warn' : kind === 'error' ? 'error' : '');
}

function normalizeHeader(h) {
  return String(h || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function findColumnIndex(headers, desired) {
  const want = normalizeHeader(desired);
  for (let i=0; i<headers.length; i++) {
    if (normalizeHeader(headers[i]) === want) return i;
  }
  return -1;
}

function safeCell(row, idx) {
  if (idx < 0) return '';
  return String((row[idx] ?? '')).trim();
}

function normalizeKey(v) {
  return String(v || '').trim().toLowerCase();
}

function renderChips(target, headers) {
  clearNode(target);
  if (!headers || headers.length === 0) {
    const span = document.createElement('div');
    span.className = 'hint muted';
    span.textContent = 'No headers detected (empty sheet?)';
    target.appendChild(span);
    return;
  }
  headers.forEach((h, i) => {
    const chip = document.createElement('div');
    chip.className = 'chip';
    const name = document.createElement('span');
    name.textContent = String(h || '').trim() || '(blank)';
    const idx = document.createElement('span');
    idx.className = 'mono';
    idx.textContent = `#${i+1}`;
    chip.appendChild(name);
    chip.appendChild(idx);
    target.appendChild(chip);
  });
}

function hasNonEmpty(v) {
  return String(v ?? '').trim() !== '';
}

function filterRowsByRequiredColumn(rows, idx) {
  if (idx === -1) return { rows, dropped: 0 };
  const out = [];
  let dropped = 0;
  for (const r of rows) {
    if (hasNonEmpty(r[idx])) out.push(r);
    else dropped++;
  }
  return { rows: out, dropped };
}

function unique(arr) {
  return Array.from(new Set(arr));
}

function escapeHtml(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function overlapTokenSet(aText, bText) {
  const a = new Set(unique(tokenizeLoose(aText)));
  const b = new Set(unique(tokenizeLoose(bText)));
  const inter = new Set();
  for (const x of a) if (b.has(x)) inter.add(x);
  return inter;
}

// Highlight tokens by splitting into alphanumeric runs so we only highlight whole tokens.
function highlightTokensInText(text, tokenSet) {
  const s = String(text ?? '');
  const parts = s.split(/([a-zA-Z0-9]+)/g);
  let out = '';
  for (const part of parts) {
    if (!part) continue;
    if (/^[a-zA-Z0-9]+$/.test(part)) {
      const k = part.toLowerCase();
      if (tokenSet && tokenSet.has(k)) {
        out += `<mark class="tok">${escapeHtml(part)}</mark>`;
      } else {
        out += escapeHtml(part);
      }
    } else {
      out += escapeHtml(part);
    }
  }
  return out;
}
