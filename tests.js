// -----------------------------
// Self-tests
// -----------------------------
function runSelfTests() {
  const tests = [];
  const t = (name, fn) => tests.push({ name, fn });

  t('stripPath handles backslashes', () => stripPath('C\\\\a\\\\b\\\\file.idw') === 'file.idw');
  t('stripPath trims whitespace', () => stripPath('  C\\\\a\\\\b\\\\file.idw  ') === 'file.idw');
  t('normalizeKey is case-insensitive', () => normalizeKey('AbC123') === 'abc123');
  t('baseNameFromAny drops ext', () => baseNameFromAny('a/b/Thing.IDW') === 'thing');
  t('detectExt finds idw in filetype', () => detectExt('Autodesk Inventor Drawing (.idw)', 'Thing') === '.idw');
  t('detectExt finds pdf in name', () => detectExt('', 'ABC.PDF') === '.pdf');

  t('parseLooseDate parses dd/mm/yyyy', () => {
    const d = parseLooseDate('13/01/2026');
    return !!d && d.getUTCFullYear() === 2026;
  });

  t('filterRowsByRequiredColumn drops empty Part Code rows', () => {
    const rows = [['A'], [''], ['  '], ['B']];
    const r = filterRowsByRequiredColumn(rows, 0);
    return r.rows.length === 2 && r.dropped === 2;
  });

  t('tokenizeLoose keeps m12 and drops stopwords', () => {
    const toks = tokenizeLoose('The M12 bolt and nut');
    return toks.includes('m12') && toks.includes('bolt') && toks.includes('nut') && !toks.includes('the') && !toks.includes('and');
  });

  t('tokenizeLoose splits mixed alphanumerics', () => {
    const toks = tokenizeLoose('Plate 150mm wide');
    return toks.includes('plate') && toks.includes('150') && toks.includes('wide');
  });

  t('findBestTitleMatch picks obvious overlap', () => {
    const prevTok = state.audit.titleTokenIndex;
    const prevMap = state.audit.titleToEntries;
    state.audit.titleTokenIndex = new Map();
    state.audit.titleToEntries = new Map();

    const addTitle = (title) => {
      const k = normalizeKey(title);
      state.audit.titleToEntries.set(k, [{ name: title, ext: '.pdf', filetype: '.pdf', state: '' }]);
      for (const tok of unique(tokenizeLoose(title))) {
        if (!state.audit.titleTokenIndex.has(tok)) state.audit.titleTokenIndex.set(tok, new Set());
        state.audit.titleTokenIndex.get(tok).add(k);
      }
    };
    addTitle('M12 Bolt Assembly 150mm');
    addTitle('Hydraulic Hose Kit');

    const best = findBestTitleMatch('Bolt assembly M12 150 mm');
    const ok = !!best.titleKey && best.titleKey.includes('m12 bolt assembly');

    state.audit.titleTokenIndex = prevTok;
    state.audit.titleToEntries = prevMap;
    return ok;
  });

  t('findBestTitleMatch accepts looser thresholds when supplied', () => {
    const prevTok = state.audit.titleTokenIndex;
    const prevMap = state.audit.titleToEntries;
    state.audit.titleTokenIndex = new Map();
    state.audit.titleToEntries = new Map();

    const addTitle = (title) => {
      const k = normalizeKey(title);
      state.audit.titleToEntries.set(k, [{ name: title, ext: '.pdf', filetype: '.pdf', state: '' }]);
      for (const tok of unique(tokenizeLoose(title))) {
        if (!state.audit.titleTokenIndex.has(tok)) state.audit.titleTokenIndex.set(tok, new Set());
        state.audit.titleTokenIndex.get(tok).add(k);
      }
    };
    addTitle('Hydraulic Hose Kit');

    const best = findBestTitleMatch('Hydraulic kit', { minSharedTokens: 2, minJaccard: 0.20 });
    const ok = !!best.titleKey;

    state.audit.titleTokenIndex = prevTok;
    state.audit.titleToEntries = prevMap;
    return ok;
  });

  t('highlightTokensInText highlights only whole alnum tokens', () => {
    const s = 'M12 bolt - bolt.';
    const set = new Set(['m12','bolt']);
    const out = highlightTokensInText(s, set);
    return out.includes('<mark') && out.toLowerCase().includes('m12') && out.toLowerCase().includes('bolt');
  });

  let pass = 0;
  for (const test of tests) {
    let ok = false;
    try {
      ok = !!test.fn();
    } catch (e) {
      ok = false;
    }
    if (!ok) {
      log(`Self-test failed: ${test.name}`, 'warn');
    } else {
      pass++;
    }
  }

  if (pass !== tests.length) {
    log(`Self-tests: ${pass}/${tests.length} passed`, 'warn');
  } else {
    log(`Self-tests: ${pass}/${tests.length} passed`, 'info');
  }
}

// Boot
renderAuditCounts();
renderVaultFiletypeSummary();
ensureReviewVisibility();
runSelfTests();
log('Ready. Load your XLSX files to begin.', 'info');
