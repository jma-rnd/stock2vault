// -----------------------------
// XLSX parsing
// -----------------------------
async function readXlsx(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  const firstSheetName = wb.SheetNames[0];
  const ws = wb.Sheets[firstSheetName];
  if (!ws) return { headers: [], rows: [], sheet: firstSheetName };

  const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  if (!data || data.length === 0) return { headers: [], rows: [], sheet: firstSheetName };

  const headers = (data[0] || []).map(h => String(h ?? '').trim());
  const rows = data.slice(1);

  // Trim trailing completely-empty rows (common when Excel has formatting)
  const isEmptyRow = (r) => {
    if (!r) return true;
    for (const cell of r) {
      if (String(cell ?? '').trim() !== '') return false;
    }
    return true;
  };
  let end = rows.length;
  while (end > 0 && isEmptyRow(rows[end - 1])) end--;

  return { headers, rows: rows.slice(0, end), sheet: firstSheetName };
}
