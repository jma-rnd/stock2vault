// -----------------------------
// Audit classification
// -----------------------------
function isReleasedState(stateVal) {
  return normalizeHeader(stateVal) === 'released';
}

function isFolderType(typeVal) {
  return normalizeHeader(typeVal) === normalizeHeader('Folder (Folder)');
}

function classifyStockRow(vaultMatches, stockKeyBase) {
  if (!vaultMatches || vaultMatches.length === 0) return 'missing';

  const hasFolder = vaultMatches.some(v => isFolderType(v.filetype));
  const hasNonFolder = vaultMatches.some(v => !isFolderType(v.filetype));
  if (hasFolder && !hasNonFolder) return 'folder';

  const idws = vaultMatches.filter(v => v.ext === '.idw');
  if (idws.length) {
    const anyRel = idws.some(v => isReleasedState(v.state));
    return anyRel ? 'released' : 'unreleased';
  }

  if (vaultMatches.some(v => v.ext === '.pdf')) return 'pdf';

  const models = vaultMatches.filter(v => v.ext === '.ipt' || v.ext === '.iam');
  if (models.length) {
    const base = stockKeyBase || '';
    const anyIdwSameBase = vaultMatches.some(v => v.ext === '.idw' && v.base === base);
    if (!anyIdwSameBase) return 'modelled';
  }

  return 'missing';
}

function resetCounts() {
  state.audit.counts = {
    released: 0,
    unreleased: 0,
    pdf: 0,
    modelled: 0,
    folder: 0,
    missing: 0,
    totalConsidered: 0,
  };
}

function resetReview() {
  state.audit.review.items = [];
  state.audit.review.cursor = 0;
  // keep decisions map across runs unless you want it reset; for now keep it.
  // state.audit.review.decisions = new Map();
}

function setReviewStatus(text, kind='muted') {
  if (!el.reviewStatus) return;
  el.reviewStatus.textContent = text;
  el.reviewStatus.className = 'pill ' + (kind === 'ok' ? 'ok' : kind === 'warn' ? 'warn' : kind === 'error' ? 'error' : '');
}

function reviewKey(item) {
  // Stable-ish key: stock part code + matched title + rule
  return `${normalizeKey(item.partCode)}|${normalizeKey(item.title)}|${item.rule}`;
}

function ensureReviewVisibility() {
  const anyRule = !!state.audit.useDescTitleRule || !!state.audit.useDescTitleTunableRule;
  const hasItems = state.audit.review.items.length > 0;
  if (!el.reviewCard) return;

  if (!anyRule) {
    el.reviewCard.style.display = 'none';
    return;
  }

  el.reviewCard.style.display = 'block';

  if (!hasItems) {
    setReviewStatus('No fuzzy matches', 'warn');
    el.reviewHint.textContent = 'Part Desc ↔ Title rules are enabled, but no matches were produced in the last run.';
    el.reviewNext.disabled = true;
    clearNode(el.reviewList);
    const empty = document.createElement('div');
    empty.className = 'hint muted';
    empty.textContent = 'Nothing to review yet. Try loosening thresholds or running the audit on a wider filter set.';
    el.reviewList.appendChild(empty);
    return;
  }

  setReviewStatus('Ready', 'ok');
  el.reviewNext.disabled = false;
}

function renderReviewBatch() {
  ensureReviewVisibility();
  if (!el.reviewCard || el.reviewCard.style.display === 'none') return;

  const items = state.audit.review.items;
  const start = state.audit.review.cursor;
  const end = Math.min(items.length, start + 10);

  clearNode(el.reviewList);

  const hdr = document.createElement('div');
  hdr.className = 'hint';
  hdr.textContent = `Showing ${start + 1}-${end} of ${items.length} fuzzy matches from the last audit run.`;
  el.reviewList.appendChild(hdr);

  for (let i = start; i < end; i++) {
    const it = items[i];
    const key = reviewKey(it);
    const decision = state.audit.review.decisions.get(key) || null;

    const card = document.createElement('div');
    card.className = 'reviewItem';

    const top = document.createElement('div');
    top.className = 'reviewTop';

    const left = document.createElement('div');
    left.style.display = 'flex';
    left.style.gap = '10px';
    left.style.flexWrap = 'wrap';
    left.style.alignItems = 'center';

    const b1 = document.createElement('div');
    b1.className = 'badge';
    b1.textContent = `${it.rule} • shared ${it.shared} • Jaccard ${it.score.toFixed(2)}`;

    const b2 = document.createElement('div');
    b2.className = 'badge';
    b2.textContent = `Part Code: ${it.partCode}`;

    left.appendChild(b1);
    left.appendChild(b2);

    const right = document.createElement('div');
    right.className = 'reviewBtns';

    const approve = document.createElement('button');
    approve.className = 'good';
    approve.textContent = 'Approve';
    approve.disabled = !!decision;
    approve.addEventListener('click', () => {
      state.audit.review.decisions.set(key, 'approved');
      renderReviewBatch();
    });

    const flag = document.createElement('button');
    flag.className = 'bad';
    flag.textContent = 'False positive';
    flag.disabled = !!decision;
    flag.addEventListener('click', () => {
      state.audit.review.decisions.set(key, 'flagged');
      renderReviewBatch();
    });

    if (decision) {
      const d = document.createElement('div');
      d.className = 'decision ' + (decision === 'approved' ? 'approved' : 'flagged');
      d.textContent = decision === 'approved' ? 'Approved' : 'Flagged as false positive';
      right.appendChild(d);
    } else {
      right.appendChild(approve);
      right.appendChild(flag);
    }

    top.appendChild(left);
    top.appendChild(right);
    card.appendChild(top);

    const tokSet = it.tokens || new Set();

    const a = document.createElement('div');
    a.innerHTML = `<div class="reviewLabel">Stock Part Desc</div><div class="reviewText">${highlightTokensInText(it.partDesc, tokSet)}</div>`;

    const b = document.createElement('div');
    b.innerHTML = `<div class="reviewLabel">Vault Title (matched)</div><div class="reviewText">${highlightTokensInText(it.title, tokSet)}</div>`;

    const c = document.createElement('div');
    c.className = 'hint muted';
    const toks = Array.from(tokSet).sort();
    c.textContent = toks.length ? `Matched tokens: ${toks.join(', ')}` : 'Matched tokens: (none)';

    card.appendChild(a);
    card.appendChild(b);
    card.appendChild(c);

    el.reviewList.appendChild(card);
  }

  el.reviewNext.disabled = (end >= items.length);
}

function renderAuditCounts() {
  const c = state.audit.counts;
  const rows = [
    { key: 'released', label: 'Released', sub: 'Filetype .idw and State = Released' },
    { key: 'unreleased', label: 'Unreleased', sub: 'Filetype .idw and State ≠ Released' },
    { key: 'pdf', label: 'Only PDF located', sub: 'Filetype .pdf' },
    { key: 'modelled', label: 'Modelled', sub: 'Filetype .ipt/.iam and no .idw' },
    { key: 'folder', label: 'Folder', sub: 'Only Folder (Folder) exists' },
    { key: 'missing', label: 'Missing', sub: 'No matching Vault row found' },
  ];
  el.auditRows.innerHTML = '';
  for (const r of rows) {
    const div = document.createElement('div');
    div.className = 'auditRow';
    div.innerHTML = '<div class="cat">' + r.label + '<span class="subtxt">' + r.sub + '</span></div>' +
                    '<div class="num">' + Number(c[r.key] || 0).toLocaleString() + '</div>';
    el.auditRows.appendChild(div);
  }
}

let _auditTimer = null;
function scheduleAuditRun() {
  if (_auditTimer) clearTimeout(_auditTimer);
  _auditTimer = setTimeout(runAuditNow, 150);
}

function runAuditNow() {
  // Requirements: both files loaded
  if (!state.stock.fileName || !state.vault.fileName) {
    setAuditStatus('Waiting for files', 'warn');
    el.auditHint.textContent = 'Load both files to compute results.';
    resetCounts();
    renderAuditCounts();
    return;
  }

  updateAuditMappings();

  // Vault must have Filetype and State columns.
  if (state.audit.vaultTypeIdx === -1 || state.audit.vaultStateIdx === -1) {
    setAuditStatus('Vault export unexpected', 'error');
    el.auditHint.textContent = 'Vault export is missing required columns: Filetype and/or State.';
    log('Vault export is missing required columns. Expected headers: "Filetype" and "State".', 'error');
    return;
  }

  // Determine match columns
  if (state.audit.useDefaultMatchRule) {
    state.audit.stockMatchCol = 'Part Code';
    state.audit.vaultMatchCol = 'Stock Number';
    updateAuditMappings();
  }

  if (state.audit.stockMatchIdx === -1 || state.audit.vaultMatchIdx === -1) {
    setAuditStatus('Missing required match columns', 'error');
    el.auditHint.textContent = 'Cannot run: match column(s) missing.';
    return;
  }

  // Build indexes
  buildVaultIndex();
  if (state.audit.useDescTitleRule || state.audit.useDescTitleTunableRule) buildTitleIndex();

  const filteredStock = getFilteredStockRows();
  resetCounts();
  resetReview();

  let fuzzyUsed = 0;
  let fuzzyUsedTunable = 0;

  for (const r of filteredStock) {
    const stockKey = normalizeKey(safeCell(r, state.audit.stockMatchIdx));
    if (!stockKey) continue;

    let matches = state.audit.vaultIndex.get(stockKey) || [];

    // Secondary fuzzy (conservative)
    if ((!matches || matches.length === 0) && state.audit.useDescTitleRule && state.audit.stockDescIdx !== -1) {
      const desc = safeCell(r, state.audit.stockDescIdx);
      const best = findBestTitleMatch(desc);
      if (best.titleKey) {
        matches = state.audit.titleToEntries.get(best.titleKey) || [];
        if (matches.length) {
          fuzzyUsed++;
          const titleText = (matches[0] && matches[0].name) ? matches[0].name : '';
          const tokSet = overlapTokenSet(desc, titleText);
          state.audit.review.items.push({
            rule: 'Title match',
            partCode: safeCell(r, state.audit.stockMatchIdx),
            partDesc: desc,
            title: titleText,
            shared: best.shared,
            score: best.score,
            tokens: tokSet,
          });
        }
      }
    }

    // Tertiary fuzzy (tunable)
    if ((!matches || matches.length === 0) && state.audit.useDescTitleTunableRule && state.audit.stockDescIdx !== -1) {
      const desc = safeCell(r, state.audit.stockDescIdx);
      const best = findBestTitleMatch(desc, {
        minSharedTokens: Number(state.audit.tunableMinSharedTokens) || FUZZY_MIN_SHARED_TOKENS,
        minJaccard: Number(state.audit.tunableMinJaccard) || FUZZY_MIN_JACCARD,
      });
      if (best.titleKey) {
        matches = state.audit.titleToEntries.get(best.titleKey) || [];
        if (matches.length) {
          fuzzyUsedTunable++;
          const titleText = (matches[0] && matches[0].name) ? matches[0].name : '';
          const tokSet = overlapTokenSet(desc, titleText);
          state.audit.review.items.push({
            rule: 'Tunable match',
            partCode: safeCell(r, state.audit.stockMatchIdx),
            partDesc: desc,
            title: titleText,
            shared: best.shared,
            score: best.score,
            tokens: tokSet,
          });
        }
      }
    }

    const base = baseNameFromAny(stockKey);
    const cat = classifyStockRow(matches, base);
    state.audit.counts[cat] = (state.audit.counts[cat] || 0) + 1;
    state.audit.counts.totalConsidered++;
  }

  renderAuditCounts();

  // Review panel (only relevant if fuzzy rules enabled)
  ensureReviewVisibility();
  if (state.audit.review.items.length) {
    state.audit.review.cursor = 0;
    renderReviewBatch();
  }

  let extra = '';
  if (state.audit.useDescTitleRule) extra += ` (Title matches used: ${fuzzyUsed.toLocaleString()})`;
  if (state.audit.useDescTitleTunableRule) extra += ` (Tunable Title matches used: ${fuzzyUsedTunable.toLocaleString()})`;

  setAuditStatus('Complete', 'ok');
  el.auditHint.textContent = `Considered ${state.audit.counts.totalConsidered.toLocaleString()} stock rows after filters.${extra}`;
}
