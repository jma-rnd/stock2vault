// -----------------------------
// Fuzzy: Title matching
// -----------------------------
const STOPWORDS = new Set([
  'a','an','and','are','as','at','be','but','by','for','from','has','have','if','in','into','is','it','its','of','on','or','s','such','t','that','the','their','then','there','these','they','this','to','was','will','with'
]);

function tokenizeLoose(s) {
  const text = String(s || '').toLowerCase();
  const raw = text.split(/[^a-z0-9]+/g).filter(Boolean);
  const out = [];

  const splitMixed = (tok) => {
    // Keep short letter+digits identifiers intact (e.g., m12, d10, ab123)
    if (/^[a-z]{1,3}\d{1,6}$/.test(tok)) return [tok];

    // Split number+letters suffix (e.g., 150mm -> 150, mm)
    if (/^\d{1,8}[a-z]{1,4}$/.test(tok)) {
      const m = tok.match(/^(\d{1,8})([a-z]{1,4})$/);
      return m ? [m[1], m[2]] : [tok];
    }

    // General split at letter<->digit boundaries
    const withBars = tok.replace(/([a-z])([0-9])/g, '$1|$2').replace(/([0-9])([a-z])/g, '$1|$2');
    return withBars.split('|').filter(Boolean);
  };

  for (const chunk of raw) {
    const parts = splitMixed(chunk);
    for (const tok of parts) {
      if (tok.length < 3) continue;
      if (STOPWORDS.has(tok)) continue;
      out.push(tok);
    }
  }
  return out;
}

function jaccard(aSet, bSet) {
  let inter = 0;
  for (const x of aSet) if (bSet.has(x)) inter++;
  const union = aSet.size + bSet.size - inter;
  return union ? inter / union : 0;
}

// Conservative defaults
const FUZZY_MIN_SHARED_TOKENS = 4;
const FUZZY_MIN_JACCARD = 0.50;

function buildTitleIndex() {
  state.audit.titleTokenIndex = new Map();
  state.audit.titleToEntries = new Map();
  if (!state.vault.rows.length) return;
  updateAuditMappings();

  const titleIdx = state.audit.vaultTitleIdx;
  const tIdx = state.audit.vaultTypeIdx;
  const sIdx = state.audit.vaultStateIdx;
  if (titleIdx === -1) return;

  for (const r of state.vault.rows) {
    const titleRaw = safeCell(r, titleIdx);
    const titleKey = normalizeKey(titleRaw);
    if (!titleKey) continue;

    const typeVal = safeCell(r, tIdx);
    const stateVal = safeCell(r, sIdx);

    const ext = detectExt(typeVal, titleRaw);
    const base = baseNameFromAny(titleRaw);

    const entry = { key: titleKey, name: titleRaw, base, ext, filetype: String(typeVal || ''), state: String(stateVal || ''), row: r };

    if (!state.audit.titleToEntries.has(titleKey)) state.audit.titleToEntries.set(titleKey, []);
    state.audit.titleToEntries.get(titleKey).push(entry);

    const toks = unique(tokenizeLoose(titleRaw));
    for (const tok of toks) {
      if (!state.audit.titleTokenIndex.has(tok)) state.audit.titleTokenIndex.set(tok, new Set());
      state.audit.titleTokenIndex.get(tok).add(titleKey);
    }
  }
}

function findBestTitleMatch(stockDesc, opts = null) {
  const cfg = {
    minSharedTokens: FUZZY_MIN_SHARED_TOKENS,
    minJaccard: FUZZY_MIN_JACCARD,
    ...(opts || {})
  };

  const toks = unique(tokenizeLoose(stockDesc));
  if (toks.length === 0) return { titleKey: null, score: 0, shared: 0 };

  const candidateCounts = new Map();
  for (const tok of toks) {
    const set = state.audit.titleTokenIndex.get(tok);
    if (!set) continue;
    for (const titleKey of set) {
      candidateCounts.set(titleKey, (candidateCounts.get(titleKey) || 0) + 1);
    }
  }

  let bestKey = null;
  let bestShared = 0;
  let bestScore = 0;

  const aSet = new Set(toks);

  for (const [titleKey, shared] of candidateCounts.entries()) {
    if (shared < cfg.minSharedTokens) continue;
    const entries = state.audit.titleToEntries.get(titleKey);
    if (!entries || entries.length === 0) continue;
    const titleText = entries[0].name;
    const bSet = new Set(unique(tokenizeLoose(titleText)));
    const score = jaccard(aSet, bSet);
    if (score > bestScore || (score === bestScore && shared > bestShared)) {
      bestKey = titleKey;
      bestShared = shared;
      bestScore = score;
    }
  }

  if (!bestKey) return { titleKey: null, score: 0, shared: 0 };
  if (bestShared < cfg.minSharedTokens) return { titleKey: null, score: 0, shared: 0 };
  if (bestScore < cfg.minJaccard) return { titleKey: null, score: 0, shared: 0 };
  return { titleKey: bestKey, score: bestScore, shared: bestShared };
}
