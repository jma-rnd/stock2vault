function refreshRulesAvailability() {
  const hasStock = state.stock.headers.length > 0;
  const hasVault = state.vault.headers.length > 0;
  const ready = hasStock && hasVault;

  if (el.exportAudit) el.exportAudit.disabled = !ready;
  if (el.exportHint) el.exportHint.textContent = ready ? 'Ready to export.' : 'Load both files to enable export.';
  if (el.exportStatus) {
    el.exportStatus.textContent = ready ? 'Ready' : 'Waiting for files';
    el.exportStatus.className = 'pill ' + (ready ? 'ok' : 'warn');
  }

  if (!ready) return;

  updateAuditMappings();

  // Validate fixed vault columns
  const fixedTypeIdx = findColumnIndex(state.vault.headers, 'Filetype');
  const fixedStateIdx = findColumnIndex(state.vault.headers, 'State');
  if (fixedTypeIdx === -1 || fixedStateIdx === -1) {
    setAuditStatus('Vault export unexpected', 'error');
    log('Vault export is missing required columns. Expected headers: "Filetype" and "State".', 'error');
  }

  // Default match rule is always enforced (manual selection removed).
  state.audit.useDefaultMatchRule = true;
  state.audit.stockMatchCol = 'Part Code';
  state.audit.vaultMatchCol = 'Stock Number';
  if (el.defaultMatchToggle) {
    el.defaultMatchToggle.checked = true;
    el.defaultMatchToggle.disabled = true;
  }

  // Desc↔Title prerequisites
  const canUseDescTitle = (findColumnIndex(state.stock.headers, 'Part Desc') !== -1) && (findColumnIndex(state.vault.headers, 'Title') !== -1);
  el.descTitleToggle.disabled = !canUseDescTitle;
  if (!canUseDescTitle) {
    el.descTitleToggle.checked = false;
    state.audit.useDescTitleRule = false;
  } else {
    state.audit.useDescTitleRule = !!el.descTitleToggle.checked;
  }

  // Meta labels
  el.defaultMatchMeta.textContent = 'Active. Matching is exact after trimming and lowercasing (case-insensitive).';

  el.descTitleMeta.textContent = !canUseDescTitle
    ? 'Unavailable. Requires Stock column "Part Desc" and Vault column "Title".'
    : (state.audit.useDescTitleRule
      ? `On. Using loose word match (min shared tokens ${FUZZY_MIN_SHARED_TOKENS}, min score ${FUZZY_MIN_JACCARD}).`
      : 'Off. If enabled, rows that fail the default match will try a word-overlap match between Stock Part Desc and Vault Title.');

  scheduleAuditRun();
}
