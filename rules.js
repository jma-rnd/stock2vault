function refreshRulesAvailability() {
  const hasStock = state.stock.headers.length > 0;
  const hasVault = state.vault.headers.length > 0;
  const ready = hasStock && hasVault;

  if (el.runAudit) el.runAudit.disabled = !ready;

  if (!ready) return;

  updateAuditMappings();

  // Validate fixed vault columns
  const fixedTypeIdx = findColumnIndex(state.vault.headers, 'Filetype');
  const fixedStateIdx = findColumnIndex(state.vault.headers, 'State');
  if (fixedTypeIdx === -1 || fixedStateIdx === -1) {
    setAuditStatus('Vault export unexpected', 'error');
    if (el.runAudit) el.runAudit.disabled = true;
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
  el.descTitleTunableToggle.disabled = !canUseDescTitle;
  if (!canUseDescTitle) {
    el.descTitleToggle.checked = false;
    el.descTitleTunableToggle.checked = false;
    state.audit.useDescTitleRule = false;
    state.audit.useDescTitleTunableRule = false;
  } else {
    state.audit.useDescTitleRule = !!el.descTitleToggle.checked;
    state.audit.useDescTitleTunableRule = !!el.descTitleTunableToggle.checked;
  }

  // Sync sliders
  el.tunableTok.value = String(state.audit.tunableMinSharedTokens);
  el.tunableJac.value = String(state.audit.tunableMinJaccard);
  el.tunableTokVal.textContent = String(state.audit.tunableMinSharedTokens);
  el.tunableJacVal.textContent = Number(state.audit.tunableMinJaccard).toFixed(2);
  el.tunableControls.style.display = state.audit.useDescTitleTunableRule ? 'grid' : 'none';

  // Meta labels
  el.defaultMatchMeta.textContent = 'Active. Matching is exact after trimming and lowercasing (case-insensitive).';

  el.descTitleMeta.textContent = !canUseDescTitle
    ? 'Unavailable. Requires Stock column "Part Desc" and Vault column "Title".'
    : (state.audit.useDescTitleRule
      ? `On. Using loose word match (min shared tokens ${FUZZY_MIN_SHARED_TOKENS}, min Jaccard ${FUZZY_MIN_JACCARD}).`
      : 'Off. If enabled, rows that fail the default match will try a word-overlap match between Stock Part Desc and Vault Title.');

  el.descTitleTunableMeta.textContent = !canUseDescTitle
    ? 'Unavailable. Requires Stock column "Part Desc" and Vault column "Title".'
    : (state.audit.useDescTitleTunableRule
      ? `On. Using tunable loose word match (min shared tokens ${state.audit.tunableMinSharedTokens}, min Jaccard ${Number(state.audit.tunableMinJaccard).toFixed(2)}).`
      : 'Off. If enabled, uses the sliders below to try to reduce false negatives (may increase false positives).');

  scheduleAuditRun();
}
