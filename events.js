// -----------------------------
// Event handlers
// -----------------------------

el.descTitleToggle.addEventListener('change', () => {
  state.audit.useDescTitleRule = !!el.descTitleToggle.checked;
  refreshRulesAvailability();
});

// Review: pagination
if (el.reviewNext) {
  el.reviewNext.addEventListener('click', () => {
    const items = state.audit.review.items;
    if (!items.length) return;
    state.audit.review.cursor = Math.min(items.length, state.audit.review.cursor + 10);
    if (state.audit.review.cursor >= items.length) {
      state.audit.review.cursor = Math.max(0, items.length - 10);
    }
    renderReviewBatch();
  });
}

el.groupSearch.addEventListener('input', renderGroupList);

el.selectAll.addEventListener('click', () => {
  for (const g of state.stock.groupSelected.keys()) state.stock.groupSelected.set(g, true);
  renderGroupList();
  scheduleAuditRun();
});

el.selectNone.addEventListener('click', () => {
  for (const g of state.stock.groupSelected.keys()) state.stock.groupSelected.set(g, false);
  renderGroupList();
  scheduleAuditRun();
});

el.selectDefault.addEventListener('click', () => {
  const out = applyDefaultGroupSelection();
  renderGroupList();
  log(`Default selection applied. ${out.selected} Group Desc values ticked.`, 'info');
  scheduleAuditRun();
});

if (el.wildcardTest) {
  el.wildcardTest.addEventListener('input', () => {
    renderWildcardMatches(el.wildcardTest.value);
  });
}

if (el.exportAudit) {
  el.exportAudit.addEventListener('click', exportAuditXlsx);
}

if (el.exportRules) {
  el.exportRules.addEventListener('click', exportRulesJson);
}

if (el.rulesImportBtn) {
  el.rulesImportBtn.addEventListener('click', () => {
    if (el.rulesImportFile && el.rulesImportFile.files && el.rulesImportFile.files[0]) {
      importRulesJson(el.rulesImportFile.files[0]);
    } else {
      log('Select a rules JSON file to import.', 'warn');
    }
  });
}

// Date filter events
el.f6m.addEventListener('change', () => setDateFilterMode(el.f6m.checked ? '6m' : 'none'));
el.f12m.addEventListener('change', () => setDateFilterMode(el.f12m.checked ? '12m' : 'none'));
el.fcustom.addEventListener('change', () => setDateFilterMode(el.fcustom.checked ? 'custom' : 'none'));
el.fStart.addEventListener('change', () => { updateDateFilterRange(); scheduleAuditRun(); });
el.fEnd.addEventListener('change', () => { updateDateFilterRange(); scheduleAuditRun(); });


// -----------------------------
// File loading
// -----------------------------
el.stockFile.addEventListener('change', async () => {
  const file = el.stockFile.files && el.stockFile.files[0];
  if (!file) return;

  try {
    setStatus('stock', 'Loading…', 'warn');
    const { headers, rows, sheet } = await readXlsx(file);
    state.stock.fileName = file.name;
    state.stock.headers = headers;

    // Option B: Only keep rows where Part Code is non-empty
    const pcIdx = findColumnIndex(headers, 'Part Code');
    const filtered = filterRowsByRequiredColumn(rows, pcIdx);
    state.stock.rawRows = filtered.rows;
    state.stock.rows = filtered.rows;

    if (filtered.dropped) {
      log(`Dropped ${filtered.dropped.toLocaleString()} stock rows with empty Part Code.`, 'info');
    }

    renderChips(el.stockHeaders, headers);

    // Enable date controls
    const dateIdx = findColumnIndex(headers, 'Last Movement Date');
    el.f6m.disabled = (dateIdx === -1);
    el.f12m.disabled = (dateIdx === -1);
    el.fcustom.disabled = (dateIdx === -1);
    el.fStart.disabled = true;
    el.fEnd.disabled = true;
    state.stock.dateFilter.mode = 'none';
    el.f6m.checked = false;
    el.f12m.checked = false;
    el.fcustom.checked = false;
    el.fStart.value = '';
    el.fEnd.value = '';
    el.fMeta.textContent = (dateIdx === -1)
      ? 'Column "Last Movement Date" not found — date filtering unavailable.'
      : 'No date filtering applied.';

    // Group Desc summary
    const gIdx = findColumnIndex(headers, 'Group Desc');
    state.stock.groupCounts = computeCounts(state.stock.rows, gIdx);
    rebuildGroupSelectionDefaultAll();

    // Enable group UI
    const hasGroups = state.stock.groupCounts.size > 0;
    el.groupSearch.disabled = !hasGroups;
    el.selectAll.disabled = !hasGroups;
    el.selectNone.disabled = !hasGroups;
    el.selectDefault.disabled = !hasGroups;

    renderGroupList();

    setStatus('stock', `Loaded (${sheet}) — ${state.stock.rows.length.toLocaleString()} rows`, 'ok');
    log(`Stock file loaded. Headers: ${headers.length}. Rows (with Part Code): ${state.stock.rows.length}. Unique Group Desc: ${state.stock.groupCounts.size}.`, 'info');

    refreshRulesAvailability();
  } catch (e) {
    console.error(e);
    setStatus('stock', 'Failed to load', 'error');
    log(`Failed to parse stock XLSX: ${e && e.message ? e.message : e}`, 'error');
  }
});

el.vaultFile.addEventListener('change', async () => {
  const file = el.vaultFile.files && el.vaultFile.files[0];
  if (!file) return;

  try {
    setStatus('vault', 'Loading…', 'warn');
    const { headers, rows, sheet } = await readXlsx(file);
    state.vault.fileName = file.name;
    state.vault.headers = headers;
    state.vault.rows = rows;

    renderChips(el.vaultHeaders, headers);
    renderVaultFiletypeSummary();

    setStatus('vault', `Loaded (${sheet}) — ${rows.length.toLocaleString()} rows`, 'ok');
    log(`Vault file loaded. Headers: ${headers.length}. Rows: ${rows.length}.`, 'info');

    refreshRulesAvailability();
  } catch (e) {
    console.error(e);
    setStatus('vault', 'Failed to load', 'error');
    log(`Failed to parse vault XLSX: ${e && e.message ? e.message : e}`, 'error');
  }
});
