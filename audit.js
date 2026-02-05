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

function exportRulesJson() {
  const payload = {
    blockedPairs: Array.from(state.audit.review.blockedPairs),
    approvedPairs: Array.from(state.audit.review.approvedPairs),
    rules: state.audit.review.rules,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'vault_audit_rules.json';
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

async function importRulesJson(file) {
  if (!file) return;
  try {
    const text = await file.text();
    const data = JSON.parse(text);
    const blocked = Array.isArray(data.blockedPairs) ? data.blockedPairs : [];
    const approved = Array.isArray(data.approvedPairs) ? data.approvedPairs : [];
    const rules = data.rules && typeof data.rules === 'object' ? data.rules : {};

    state.audit.review.blockedPairs = new Set(blocked);
    state.audit.review.approvedPairs = new Set(approved);
    state.audit.review.rules.conflictPairs = Array.isArray(rules.conflictPairs) ? rules.conflictPairs : [];
    state.audit.review.rules.conflictGroups = Array.isArray(rules.conflictGroups) ? rules.conflictGroups : [];
    state.audit.review.rules.requiredTokens = Array.isArray(rules.requiredTokens) ? rules.requiredTokens : [];
    state.audit.review.rules.requiredGroups = Array.isArray(rules.requiredGroups) ? rules.requiredGroups : [];
    state.audit.review.rules.approvedTokens = rules.approvedTokens && typeof rules.approvedTokens === 'object'
      ? rules.approvedTokens
      : {};

    renderRuleLists();
    if (el.rulesImportStatus) {
      el.rulesImportStatus.textContent = 'Imported';
      el.rulesImportStatus.className = 'pill ok';
    }
  } catch (e) {
    if (el.rulesImportStatus) {
      el.rulesImportStatus.textContent = 'Import failed';
      el.rulesImportStatus.className = 'pill error';
    }
    log(`Rules import failed: ${e && e.message ? e.message : e}`, 'error');
  }
}

function recordReviewDecision(item, decision) {
  const pairKey = reviewPairKey(item.partDesc, item.title);
  if (decision === 'flagged') {
    state.audit.review.blockedPairs.add(pairKey);
    state.audit.review.approvedPairs.delete(pairKey);
  } else if (decision === 'approved') {
    state.audit.review.approvedPairs.add(pairKey);
    state.audit.review.blockedPairs.delete(pairKey);
  }
  renderRuleLists();
}

function renderRuleLists() {
  if (!el.ruleConflicts || !el.ruleRequired || !el.ruleApproved) return;

  clearNode(el.ruleConflicts);
  clearNode(el.ruleRequired);
  clearNode(el.ruleApproved);

  const conflicts = state.audit.review.rules.conflictPairs || [];
  const conflictGroups = state.audit.review.rules.conflictGroups || [];
  const required = state.audit.review.rules.requiredTokens || [];
  const requiredGroups = state.audit.review.rules.requiredGroups || [];
  const approved = state.audit.review.rules.approvedTokens || {};

  if (!conflicts.length && !conflictGroups.length) {
    const span = document.createElement('div');
    span.className = 'hint muted';
    span.textContent = 'No conflict pairs saved.';
    el.ruleConflicts.appendChild(span);
  } else {
    for (const c of conflicts) {
      const chip = document.createElement('div');
      chip.className = 'chip small';
      chip.textContent = `${c.a} <-> ${c.b}`;
      el.ruleConflicts.appendChild(chip);
    }
    for (const g of conflictGroups) {
      const chip = document.createElement('div');
      chip.className = 'chip small';
      const aText = g.aText || (g.aTokens || []).join(' ');
      const bText = g.bText || (g.bTokens || []).join(' ');
      chip.textContent = `${aText} <-> ${bText}`;
      el.ruleConflicts.appendChild(chip);
    }
  }

  if (!required.length && !requiredGroups.length) {
    const span = document.createElement('div');
    span.className = 'hint muted';
    span.textContent = 'No required tokens saved.';
    el.ruleRequired.appendChild(span);
  } else {
    for (const t of required) {
      const chip = document.createElement('div');
      chip.className = 'chip small';
      chip.textContent = t;
      el.ruleRequired.appendChild(chip);
    }
    for (const g of requiredGroups) {
      const chip = document.createElement('div');
      chip.className = 'chip small';
      chip.textContent = g.text || (g.tokens || []).join(' ');
      el.ruleRequired.appendChild(chip);
    }
  }

  const approvedTokens = Object.entries(approved).sort((a, b) => b[1] - a[1]);
  if (!approvedTokens.length) {
    const span = document.createElement('div');
    span.className = 'hint muted';
    span.textContent = 'No approved tokens saved.';
    el.ruleApproved.appendChild(span);
  } else {
    for (const [tok, count] of approvedTokens) {
      const chip = document.createElement('div');
      chip.className = 'chip small';
      chip.textContent = `${tok} * ${count}`;
      el.ruleApproved.appendChild(chip);
    }
  }
}

function clearReviewRules() {
  state.audit.review.blockedPairs = new Set();
  state.audit.review.approvedPairs = new Set();
  state.audit.review.rules.conflictPairs = [];
  state.audit.review.rules.conflictGroups = [];
  state.audit.review.rules.requiredTokens = [];
  state.audit.review.rules.requiredGroups = [];
  state.audit.review.rules.approvedTokens = {};
  renderRuleLists();
  if (el.rulesImportStatus) {
    el.rulesImportStatus.textContent = 'Local';
    el.rulesImportStatus.className = 'pill';
  }
  log('Cleared active match rules.', 'info');
}

function ensureReviewVisibility() {
  const anyRule = !!state.audit.useDescTitleRule;
  const hasItems = state.audit.review.items.length > 0;
  if (!el.reviewCard) return;

  if (!anyRule) {
    el.reviewCard.style.display = 'none';
    return;
  }

  el.reviewCard.style.display = 'block';

  if (!hasItems) {
    setReviewStatus('No fuzzy matches', 'warn');
    el.reviewHint.textContent = 'Part Desc <-> Title rules are enabled, but no matches were produced in the last run.';
    el.reviewNext.disabled = true;
    clearNode(el.reviewList);
    const empty = document.createElement('div');
    empty.className = 'hint muted';
    empty.textContent = 'Nothing to review yet. Try loosening thresholds or running the audit on a wider filter set.';
    el.reviewList.appendChild(empty);
    renderRuleLists();
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
    const tokSet = it.tokens || new Set();
    const sel = state.audit.review.selections.get(key) || { stock: null, vault: null };

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
    b1.textContent = `${it.rule} * shared ${it.shared} * score ${it.score.toFixed(2)}`;

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
      recordReviewDecision(it, 'approved');
      const matched = Array.from(tokSet || []);
      for (const tok of matched) {
        state.audit.review.rules.approvedTokens[tok] = (state.audit.review.rules.approvedTokens[tok] || 0) + 1;
      }
      renderRuleLists();
      renderReviewBatch();
    });

    const flag = document.createElement('button');
    flag.className = 'bad';
    flag.textContent = 'False positive';
    flag.disabled = !!decision;
    flag.addEventListener('click', () => {
      state.audit.review.decisions.set(key, 'flagging');
      renderReviewBatch();
    });

    if (decision && decision !== 'flagging') {
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
    const onSelect = (side, text, node) => {
      const selection = window.getSelection ? window.getSelection() : null;
      if (!selection || selection.isCollapsed) return;
      if (!node.contains(selection.anchorNode) || !node.contains(selection.focusNode)) return;
      const selectedText = selection.toString().trim();
      if (!selectedText) return;
      const tokens = unique(tokenizeLoose(selectedText));
      if (!tokens.length) return;
      const current = state.audit.review.selections.get(key) || { stock: null, vault: null };
      if (side === 'stock') {
        current.stock = { text: selectedText, tokens };
      } else {
        current.vault = { text: selectedText, tokens };
      }
      state.audit.review.selections.set(key, current);
      renderReviewBatch();
    };

    const a = document.createElement('div');
    const aLabel = document.createElement('div');
    aLabel.className = 'reviewLabel';
    aLabel.textContent = 'Stock Part Desc';
    const aText = document.createElement('div');
    aText.className = 'reviewText';
    aText.innerHTML = highlightTokensInText(it.partDesc, tokSet);
    aText.addEventListener('mouseup', () => onSelect('stock', it.partDesc, aText));
    a.appendChild(aLabel);
    a.appendChild(aText);

    const b = document.createElement('div');
    const bLabel = document.createElement('div');
    bLabel.className = 'reviewLabel';
    bLabel.textContent = 'Vault Title (matched)';
    const bText = document.createElement('div');
    bText.className = 'reviewText';
    bText.innerHTML = highlightTokensInText(it.title, tokSet);
    bText.addEventListener('mouseup', () => onSelect('vault', it.title, bText));
    b.appendChild(bLabel);
    b.appendChild(bText);
    const dn = document.createElement('div');
    dn.className = 'hint';
    dn.textContent = `Drawing Number: ${it.drawingNumber || '(none)'}`;
    b.appendChild(dn);

    const pn = document.createElement('div');
    pn.className = 'hint';
    pn.textContent = `Part Number: ${it.partNumber || '(none)'}`;
    b.appendChild(pn);

    const c = document.createElement('div');
    c.className = 'hint muted';
    const toks = Array.from(tokSet).sort();
    c.textContent = toks.length ? `Matched tokens: ${toks.join(', ')}` : 'Matched tokens: (none)';

    card.appendChild(a);
    card.appendChild(b);
    card.appendChild(c);

    if (decision === 'flagging' || decision === 'flagged') {
      const actions = document.createElement('div');
      actions.className = 'reviewBtns';

      const inst = document.createElement('div');
      inst.className = 'hint';
      inst.textContent = 'Select a phrase in the Stock and Vault text, then click "Save conflict rule".';
      actions.appendChild(inst);

        if (sel.stock || sel.vault) {
        const hint = document.createElement('div');
        hint.className = 'hint';
        const parts = [];
        if (sel.stock) parts.push(`Stock: "${sel.stock.text}"`);
        if (sel.vault) parts.push(`Vault: "${sel.vault.text}"`);
        hint.textContent = parts.join(' | ');
        actions.appendChild(hint);
      }

      const save = document.createElement('button');
      save.className = 'secondary';
      save.textContent = 'Save conflict rule';
      save.disabled = !sel.stock && !sel.vault;
      save.addEventListener('click', () => {
        const stockSel = sel.stock;
        const vaultSel = sel.vault;
        if (stockSel && vaultSel) {
          state.audit.review.rules.conflictGroups.push({
            aTokens: stockSel.tokens,
            bTokens: vaultSel.tokens,
            aText: stockSel.text,
            bText: vaultSel.text,
          });
        } else {
          const single = stockSel || vaultSel;
          if (single && !state.audit.review.rules.requiredGroups.some(g => g.text === single.text)) {
            state.audit.review.rules.requiredGroups.push({
              tokens: single.tokens,
              text: single.text,
            });
          }
        }
        state.audit.review.selections.set(key, { stock: null, vault: null });
        state.audit.review.decisions.set(key, 'flagged');
        renderRuleLists();
        renderReviewBatch();
      });

      const clear = document.createElement('button');
      clear.className = 'secondary';
      clear.textContent = 'Clear selection';
      clear.addEventListener('click', () => {
        state.audit.review.selections.set(key, { stock: null, vault: null });
        renderReviewBatch();
      });

      const addAnother = document.createElement('button');
      addAnother.className = 'secondary';
      addAnother.textContent = 'Add another conflict';
      addAnother.addEventListener('click', () => {
        state.audit.review.decisions.set(key, 'flagging');
        renderReviewBatch();
      });

      actions.appendChild(save);
      actions.appendChild(clear);
      if (decision === 'flagged') actions.appendChild(addAnother);
      card.appendChild(actions);
    }
    el.reviewList.appendChild(card);
  }

  el.reviewNext.disabled = (end >= items.length);
  renderRuleLists();
}

function renderAuditCounts() {
  const c = state.audit.counts;
  const rows = [
    { key: 'released', label: 'Released', sub: 'Filetype .idw and State = Released' },
    { key: 'unreleased', label: 'Unreleased', sub: 'Filetype .idw and State != Released' },
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
  if (state.audit.useDescTitleRule) buildTitleIndex();

      const filteredStock = getFilteredStockRows();
  resetCounts();
  resetReview();

  let fuzzyUsed = 0;

  for (const r of filteredStock) {
    const stockKey = normalizeKey(safeCell(r, state.audit.stockMatchIdx));
    if (!stockKey) continue;

        let matches = findVaultMatches(safeCell(r, state.audit.stockMatchIdx));

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
          const drawingNumber = (matches[0] && matches[0].row)
            ? safeCell(matches[0].row, state.audit.vaultPartNumberIdx)
            : '';
          const partNumber = (matches[0] && matches[0].row)
            ? safeCell(matches[0].row, state.audit.vaultMatchIdx)
            : '';
          state.audit.review.items.push({
            rule: 'Title match',
            partCode: safeCell(r, state.audit.stockMatchIdx),
            partDesc: desc,
            title: titleText,
            shared: best.shared,
            score: best.score,
            tokens: tokSet,
            drawingNumber,
            partNumber,
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

  setAuditStatus('Complete', 'ok');
  el.auditHint.textContent = `Considered ${state.audit.counts.totalConsidered.toLocaleString()} stock rows after filters.${extra}`;
}

function exportAuditXlsx() {
  if (!state.stock.fileName || !state.vault.fileName) {
    if (el.exportStatus) {
      el.exportStatus.textContent = 'Missing files';
      el.exportStatus.className = 'pill warn';
    }
    log('Export aborted: load both Stock and Vault files first.', 'warn');
    return;
  }

  updateAuditMappings();
  buildVaultIndex();
  if (state.audit.useDescTitleRule) buildTitleIndex();

  const includeUnmatched = true;
  const includeVaultDetails = true;

  const rows = [];
  const filteredStock = getFilteredStockRows();
  for (const r of filteredStock) {
    const partCode = safeCell(r, state.audit.stockMatchIdx);
    if (!partCode) continue;

    let matches = [];
    let matchType = 'none';

    const exactKey = normalizeKey(partCode);
    const exact = state.audit.vaultIndex.get(exactKey) || [];
    if (exact.length) {
      matches = exact;
      matchType = 'exact';
    } else {
      const wildcardMatches = findVaultMatches(partCode);
      if (wildcardMatches.length) {
        matches = wildcardMatches;
        if (wildcardMatches[0].pdfNameMatch) {
          matchType = 'pdf-name';
        } else if (wildcardMatches[0].wildcardKind === 'length') {
          matchType = 'wildcard-length';
        } else if (wildcardMatches[0].wildcardKind === 'galv') {
          matchType = 'wildcard-galv';
        } else {
          matchType = 'wildcard';
        }
      }
    }

    if (!matches.length && state.audit.useDescTitleRule && state.audit.stockDescIdx !== -1) {
      const desc = safeCell(r, state.audit.stockDescIdx);
      const best = findBestTitleMatch(desc);
      if (best.titleKey) {
        const byTitle = state.audit.titleToEntries.get(best.titleKey) || [];
        if (byTitle.length) {
          matches = byTitle;
          matchType = 'title';
        }
      }
    }

    const base = baseNameFromAny(normalizeKey(partCode));
    const category = classifyStockRow(matches, base);
    if (!includeUnmatched && category === 'missing') continue;

    const out = {
      'Group Desc': safeCell(r, state.audit.stockGroupIdx),
      'Part Code': partCode,
      'Part Desc': safeCell(r, state.audit.stockDescIdx),
      'Category': category,
      'Drawing Number': '',
      'Vault States': '',
      'Vault Filetypes': '',
      'Match Type': matchType,
      'Vault Matched Phrase': '',
    };

    if (includeVaultDetails) {
      const names = unique(matches.map(m => m.name || m.key));
      const types = unique(matches.map(m => m.filetype).filter(Boolean));
      const states = unique(matches.map(m => m.state).filter(Boolean));
      const partNumbers = unique(
        matches
          .map(m => safeCell(m.row || [], state.audit.vaultPartNumberIdx))
          .filter(Boolean)
      );
      out['Drawing Number'] = partNumbers.join('; ');
      out['Vault States'] = states.join('; ');
      out['Vault Filetypes'] = types.join('; ');
      out['Vault Matched Phrase'] = names.join('; ');
    }

    rows.push(out);
  }

  const sheet = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet, 'Audit');

  const stamp = new Date().toISOString().slice(0, 10);
  const fileName = `vault_audit_${stamp}.xlsx`;
  XLSX.writeFile(wb, fileName);

  if (el.exportStatus) {
    el.exportStatus.textContent = 'Exported';
    el.exportStatus.className = 'pill ok';
  }
  log(`Exported audit XLSX: ${fileName}`, 'info');
}
