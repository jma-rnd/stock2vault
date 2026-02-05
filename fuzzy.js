// -----------------------------
// Fuzzy: Title matching
// -----------------------------
const STOPWORDS = new Set([
  'a','an','and','are','as','at','be','but','by','for','from','has','have','if','in','into','is','it','its','of','on','or','s','such','t','that','the','their','then','there','these','they','this','to','was','will','with'
]);

const ALLOW_SHORT = new Set(['tc', 'gal']);

function tokenizeLoose(s) {
  const text = String(s || '').toLowerCase();
  const raw = text.split(/[^a-z0-9./]+/g).filter(Boolean);
  const out = [];

  const splitMixed = (tok) => {
    // Keep fractions intact (e.g., 3/4)
    if (/^\d{1,3}\/\d{1,3}$/.test(tok)) return [tok];

    // Keep decimals intact (e.g., 15.8)
    if (/^\d{1,8}\.\d{1,4}$/.test(tok)) return [tok];

    // Keep short letter+digits identifiers intact (e.g., m12, d10, ab123)
    if (/^[a-z]{1,3}\d{1,6}$/.test(tok)) return [tok];

    // Split number+letters suffix (e.g., 150mm -> 150, mm)
    if (/^\d{1,8}(?:\.\d{1,4})?[a-z]{1,4}$/.test(tok)) {
      const m = tok.match(/^(\d{1,8}(?:\.\d{1,4})?)([a-z]{1,4})$/);
      return m ? [m[1], m[2]] : [tok];
    }

    // General split at letter<->digit boundaries
    const withBars = tok.replace(/([a-z])([0-9])/g, '$1|$2').replace(/([0-9])([a-z])/g, '$1|$2');
    return withBars.split('|').filter(Boolean);
  };

  for (const chunk of raw) {
    const parts = splitMixed(chunk);
    for (const tok of parts) {
      if (tok.length < 3) {
        if (!/^\d+(\.\d+)?$/.test(tok) && !/^\d{1,3}\/\d{1,3}$/.test(tok) && !ALLOW_SHORT.has(tok)) continue;
      }
      if (STOPWORDS.has(tok)) continue;
      out.push(tok);
    }
  }
  return out;
}

const CRITICAL_MATERIALS = new Set(['tc','carbide','tungsten','gal','galv','galvanised','galvanized','black']);

function extractCriticalNumbers(text) {
  const s = String(text || '').toLowerCase();
  const counts = new Map();

  const add = (num) => {
    if (!num) return;
    const n = String(num);
    counts.set(n, (counts.get(n) || 0) + 1);
  };

  const qtyRe = /(\d{1,4}(?:\.\d{1,4})?)\s*(per|pallet|pack|pk|qty|quantity)\b/g;
  const qty = new Set();
  let m;
  while ((m = qtyRe.exec(s)) !== null) {
    qty.add(String(m[1]));
  }

  const dimRe = /(\d{1,4}(?:\.\d{1,4})?)\s*[x×]\s*(\d{1,4}(?:\.\d{1,4})?)(?:\s*[x×]\s*(\d{1,4}(?:\.\d{1,4})?))?/g;
  while ((m = dimRe.exec(s)) !== null) {
    add(m[1]); add(m[2]); if (m[3]) add(m[3]);
  }

  const mmRe = /(\d{1,4}(?:\.\d{1,4})?)\s*mm\b/g;
  while ((m = mmRe.exec(s)) !== null) add(m[1]);

  const holeRe = /(\d{1,4}(?:\.\d{1,4})?)\s*mm?\s*(hole|dia|diam|diameter)\b/g;
  while ((m = holeRe.exec(s)) !== null) add(m[1]);

  const holeRe2 = /\b(hole|dia|diam|diameter)\s*(\d{1,4}(?:\.\d{1,4})?)\b/g;
  while ((m = holeRe2.exec(s)) !== null) add(m[2]);

  const thickRe = /(\d{1,4}(?:\.\d{1,4})?)\s*mm?\s*(thick|thickness)\b/g;

  const fracRe = /(\d{1,3}\/\d{1,3})/g;
  while ((m = fracRe.exec(s)) !== null) add(m[1]);
  while ((m = thickRe.exec(s)) !== null) add(m[1]);

  // Remove quantity-related numbers
  for (const q of qty) counts.delete(q);

  return counts;
}

function extractCriticalCodes(tokens) {
  const out = new Set();
  for (const tok of tokens) {
    if (/^[a-z]{1,3}\d{1,4}$/.test(tok)) out.add(tok);
  }
  return out;
}

function extractCriticalMaterials(tokens) {
  const out = new Set();
  for (const tok of tokens) {
    if (CRITICAL_MATERIALS.has(tok)) out.add(tok);
  }
  return out;
}

function countsEqual(a, b) {
  if (a.size !== b.size) return false;
  for (const [k, v] of a.entries()) {
    if (b.get(k) !== v) return false;
  }
  return true;
}

function setsEqual(a, b) {
  if (a.size !== b.size) return false;
  for (const v of a) if (!b.has(v)) return false;
  return true;
}

function buildTokenMeta(text) {
  const tokens = unique(tokenizeLoose(text));
  const criticalNumbers = extractCriticalNumbers(text);
  const criticalCodes = extractCriticalCodes(tokens);
  const criticalMaterials = extractCriticalMaterials(tokens);

  for (const n of criticalNumbers.keys()) tokens.push(n);
  for (const c of criticalCodes) tokens.push(c);
  for (const m of criticalMaterials) tokens.push(m);

  return {
    tokens: unique(tokens),
    criticalNumbers,
    criticalCodes,
    criticalMaterials,
  };
}

function hasConflictRule(aTokens, bTokens) {
  const rules = state.audit.review.rules || {};
  const conflicts = rules.conflictPairs || [];
  const groups = rules.conflictGroups || [];
  if (!conflicts.length && !groups.length) return false;

  const aSet = new Set(aTokens);
  const bSet = new Set(bTokens);

  for (const c of conflicts) {
    if (!c || !c.a || !c.b) continue;
    const a = String(c.a).toLowerCase();
    const b = String(c.b).toLowerCase();
    if ((aSet.has(a) && bSet.has(b)) || (aSet.has(b) && bSet.has(a))) return true;
  }
  for (const g of groups) {
    if (!g || !g.aTokens || !g.bTokens) continue;
    const aGroup = g.aTokens.map(t => String(t).toLowerCase());
    const bGroup = g.bTokens.map(t => String(t).toLowerCase());
    const aIn = aGroup.every(t => aSet.has(t));
    const bIn = bGroup.every(t => bSet.has(t));
    const aInRev = aGroup.every(t => bSet.has(t));
    const bInRev = bGroup.every(t => aSet.has(t));
    if ((aIn && bIn) || (aInRev && bInRev)) return true;
  }
  return false;
}

function hasMissingRequiredToken(aTokens, bTokens) {
  const rules = state.audit.review.rules || {};
  const required = rules.requiredTokens || [];
  const groups = rules.requiredGroups || [];
  if (!required.length && !groups.length) return false;

  const aSet = new Set(aTokens);
  const bSet = new Set(bTokens);
  for (const t of required) {
    const tok = String(t).toLowerCase();
    if ((aSet.has(tok) && !bSet.has(tok)) || (!aSet.has(tok) && bSet.has(tok))) return true;
  }
  for (const g of groups) {
    if (!g || !g.tokens) continue;
    const toks = g.tokens.map(t => String(t).toLowerCase());
    const aHas = toks.every(t => aSet.has(t));
    const bHas = toks.every(t => bSet.has(t));
    if (aHas !== bHas) return true;
  }
  return false;
}

function approvalBonus(aTokens, bTokens) {
  const rules = state.audit.review.rules || {};
  const approved = rules.approvedTokens || {};
  const aSet = new Set(aTokens);
  const bSet = new Set(bTokens);
  let total = 0;
  for (const tok of aSet) {
    if (!bSet.has(tok)) continue;
    const count = approved[tok] || 0;
    total += count;
  }
  return Math.min(0.2, total * 0.01);
}

function weightedJaccard(aTokens, bTokens, aMeta, bMeta) {
  const aSet = new Set(aTokens);
  const bSet = new Set(bTokens);
  const all = new Set([...aSet, ...bSet]);

  const weightFor = (tok, meta) => {
    if (meta.criticalNumbers.has(tok)) return 3;
    if (meta.criticalCodes.has(tok)) return 3;
    if (meta.criticalMaterials.has(tok)) return 2;
    return 1;
  };

  let inter = 0;
  let union = 0;
  for (const tok of all) {
    const w = Math.max(weightFor(tok, aMeta), weightFor(tok, bMeta));
    union += w;
    if (aSet.has(tok) && bSet.has(tok)) inter += w;
  }
  return union ? inter / union : 0;
}

// Conservative defaults
const FUZZY_MIN_SHARED_TOKENS = 3;
const FUZZY_MIN_JACCARD = 0.40;

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

  const stockMeta = buildTokenMeta(stockDesc);
  const toks = stockMeta.tokens;
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
    const pairKey = reviewPairKey(stockDesc, titleText);
    if (state.audit.review.blockedPairs && state.audit.review.blockedPairs.has(pairKey)) continue;

    const titleMeta = buildTokenMeta(titleText);

    if (stockMeta.criticalNumbers.size && titleMeta.criticalNumbers.size) {
      if (!countsEqual(stockMeta.criticalNumbers, titleMeta.criticalNumbers)) continue;
    }
    if (stockMeta.criticalCodes.size && titleMeta.criticalCodes.size) {
      if (!setsEqual(stockMeta.criticalCodes, titleMeta.criticalCodes)) continue;
    }
    if (stockMeta.criticalMaterials.size && titleMeta.criticalMaterials.size) {
      if (!setsEqual(stockMeta.criticalMaterials, titleMeta.criticalMaterials)) continue;
    }

    if (hasConflictRule(stockMeta.tokens, titleMeta.tokens)) continue;
    if (hasMissingRequiredToken(stockMeta.tokens, titleMeta.tokens)) continue;

    const bSet = new Set(titleMeta.tokens);
    let score = weightedJaccard(toks, titleMeta.tokens, stockMeta, titleMeta);
    let sharedCount = 0;
    for (const t of aSet) if (bSet.has(t)) sharedCount++;

    score = Math.min(1, score + approvalBonus(stockMeta.tokens, titleMeta.tokens));

    if (state.audit.review.approvedPairs && state.audit.review.approvedPairs.has(pairKey)) {
      score = 1;
      sharedCount = Math.max(sharedCount, cfg.minSharedTokens);
    }

    if (score > bestScore || (score === bestScore && sharedCount > bestShared)) {
      bestKey = titleKey;
      bestShared = sharedCount;
      bestScore = score;
    }
  }

  if (!bestKey) return { titleKey: null, score: 0, shared: 0 };
  if (bestShared < cfg.minSharedTokens) return { titleKey: null, score: 0, shared: 0 };
  if (bestScore < cfg.minJaccard) return { titleKey: null, score: 0, shared: 0 };
  return { titleKey: bestKey, score: bestScore, shared: bestShared };
}
