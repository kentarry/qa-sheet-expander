import * as XLSX from 'xlsx';

// 解析 Excel buffer → { sheetNames, sheets: { [name]: { data: string[][] } } }
export function xlParse(buf) {
  const wb = XLSX.read(buf, { type: 'array' });
  const sheets = {};
  wb.SheetNames.forEach(n => {
    const ws = wb.Sheets[n];
    if (!ws['!ref']) { sheets[n] = { data: [[]] }; return; }
    const rng = XLSX.utils.decode_range(ws['!ref']);
    const data = [];
    for (let r = rng.s.r; r <= rng.e.r; r++) {
      const row = [];
      for (let c = rng.s.c; c <= rng.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        row.push(cell?.v !== undefined ? String(cell.v) : '');
      }
      data.push(row);
    }
    sheets[n] = { data };
  });
  return { sheetNames: wb.SheetNames, sheets };
}

// 欄位索引轉字母 (0→A, 25→Z, 26→AA)
function colToLetter(ci) {
  let s = '';
  let n = ci;
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

// 產出文字摘要（送給 AI），壓縮版
export function xlSummary(sheets, maxRows) {
  maxRows = maxRows || 60;
  let out = '';
  Object.entries(sheets).forEach(([name, s]) => {
    out += '\n[Sheet:' + name + '](' + s.data.length + 'rows, ' + (s.data[0]?.length || 0) + 'cols)\n';
    s.data.slice(0, maxRows).forEach((row, i) => {
      const cells = row.map((v, ci) => {
        if (!v) return '';
        return colToLetter(ci) + (i + 1) + '=' + v.slice(0, 60);
      }).filter(Boolean);
      if (cells.length) out += cells.join(' | ') + '\n';
    });
    if (s.data.length > maxRows) out += '...(' + (s.data.length - maxRows) + ' more rows)\n';
  });
  return out;
}

// 提取簽章
export function xlSignature(sheets) {
  const main = Object.values(sheets)[0];
  if (!main) return '';
  const keywords = [];
  main.data.slice(0, 40).forEach(row => {
    row.forEach(cell => {
      if (cell && cell.length > 1 && cell.length < 60) keywords.push(cell);
    });
  });
  return keywords.slice(0, 30).join('|');
}

// ═══ 結構性差異比對（展表前 vs 展表後）═══
// 重點：描述「結構變化模式」而非逐格差異
// 排除因欄位移動造成的假差異
export function xlDiff(beforeSheets, afterSheets) {
  let diff = '';
  const bNames = Object.keys(beforeSheets);
  const aNames = Object.keys(afterSheets);

  // 比對同名或相同位置的 sheet
  const paired = [];
  for (let i = 0; i < Math.max(bNames.length, aNames.length); i++) {
    if (i < bNames.length && i < aNames.length) {
      // 同名或相同位置
      const bName = bNames[i];
      const aName = aNames.find(n => n === bName) || aNames[i];
      if (aName && !paired.some(p => p.aName === aName)) {
        paired.push({ bName, aName });
      }
    }
  }
  // 確保沒遺漏
  for (const an of aNames) {
    if (!paired.some(p => p.aName === an)) {
      const matchB = bNames.find(bn => bn === an);
      if (matchB) paired.push({ bName: matchB, aName: an });
      else diff += '\nNEW_SHEET: ' + an + ' (' + afterSheets[an].data.length + ' rows)\n';
    }
  }

  for (const { bName, aName } of paired) {
    const bd = beforeSheets[bName].data;
    const ad = afterSheets[aName].data;
    const sheetDiff = structuralDiff(bd, ad);
    if (sheetDiff) {
      const label = bName === aName ? bName : bName + ' → ' + aName;
      diff += '\n--- [Sheet:' + label + '] ---\n' + sheetDiff;
    }
  }

  return diff || '無差異';
}

// 結構性 diff：偵測新增欄、新增行
// 提供更詳細的插入位置上下文，幫助 AI 精確定位
function structuralDiff(beforeData, afterData) {
  let result = '';

  const bCols = beforeData[0]?.length || 0;
  const aCols = afterData[0]?.length || 0;
  const bRows = beforeData.length;
  const aRows = afterData.length;

  // 1. 欄位差異 — 詳細列出每個新欄在每行的標題
  if (aCols > bCols) {
    result += '新增欄位: ' + (aCols - bCols) + ' 欄 (展前 ' + bCols + ' → 展後 ' + aCols + ')\n';
    // 找出新欄是在哪個欄之後插入的
    // 比對展前展後的欄頭，找到分歧點
    const headerRow = afterData[0] || [];
    const beforeHeader = beforeData[0] || [];
    let insertAfterCol = bCols - 1; // 預設插在最後
    for (let c = 0; c < bCols; c++) {
      if ((beforeHeader[c] || '').trim() !== (headerRow[c] || '').trim()) {
        insertAfterCol = c - 1;
        break;
      }
    }
    result += '  插入位置: 在 ' + colToLetter(Math.max(0, insertAfterCol)) + ' 欄之後\n';

    // 列出所有新欄在各行的值
    for (let c = bCols; c < aCols; c++) {
      result += '  新欄 ' + colToLetter(c) + ':\n';
      for (let r = 0; r < Math.min(30, aRows); r++) {
        const val = afterData[r]?.[c];
        if (val && val.trim()) {
          result += '    Row' + (r + 1) + '=' + val.slice(0, 50) + '\n';
        }
      }
    }
  }

  // 2. 建立行簽章索引
  const beforeRowSigs = new Map();
  beforeData.forEach((row, ri) => {
    const sig = row.filter(v => v.trim()).slice(0, 4).join('|');
    if (sig) {
      if (!beforeRowSigs.has(sig)) beforeRowSigs.set(sig, []);
      beforeRowSigs.get(sig).push(ri);
    }
  });

  // 3. 行對應 (寬鬆匹配)
  const rowMapping = [];
  const usedBefore = new Set();
  for (let ar = 0; ar < aRows; ar++) {
    const aRow = afterData[ar];
    // 用前幾個非空 cell 作匹配
    const sig = aRow.filter(v => v.trim()).slice(0, 4).join('|');
    const candidates = sig ? (beforeRowSigs.get(sig) || []) : [];
    const match = candidates.find(idx => !usedBefore.has(idx));
    if (match !== undefined) {
      rowMapping[ar] = match;
      usedBefore.add(match);
    } else {
      // 嘗試寬鬆匹配（前2個非空cell）
      const sig2 = aRow.filter(v => v.trim()).slice(0, 2).join('|');
      let found = false;
      for (const [key, indices] of beforeRowSigs.entries()) {
        if (key.startsWith(sig2)) {
          const idx = indices.find(i => !usedBefore.has(i));
          if (idx !== undefined) {
            rowMapping[ar] = idx;
            usedBefore.add(idx);
            found = true;
            break;
          }
        }
      }
      if (!found) rowMapping[ar] = -1;
    }
  }

  // 4. 報告新增行 — 重要：顯示「在哪個現有行之後」插入
  let changes = 0;
  const maxChanges = 120;

  if (aRows > bRows) {
    result += '新增行: ' + (aRows - bRows) + ' 行 (展前 ' + bRows + ' → 展後 ' + aRows + ')\n';
  }

  for (let ar = 0; ar < aRows && changes < maxChanges; ar++) {
    if (rowMapping[ar] !== -1) continue;
    const aRow = afterData[ar];
    const content = aRow.map((v, ci) => {
      if (!v || !v.trim()) return '';
      return colToLetter(ci) + '=' + v.slice(0, 40);
    }).filter(Boolean);
    if (content.length === 0) continue;

    // 找到這行在哪個「現有行」之後
    let afterContext = '';
    for (let prev = ar - 1; prev >= 0; prev--) {
      if (rowMapping[prev] >= 0) {
        const prevRow = afterData[prev];
        const prevContent = prevRow.filter(v => v.trim()).slice(0, 3).join(' | ');
        afterContext = ' (在「' + prevContent.slice(0, 60) + '」之後)';
        break;
      }
    }

    result += 'INSERTED_ROW ' + (ar + 1) + afterContext + ': ' + content.join(' | ') + '\n';
    changes++;
  }

  // 5. 報告已有行中新增的 cell 值
  for (let ar = 0; ar < aRows && changes < maxChanges; ar++) {
    const br = rowMapping[ar];
    if (br < 0) continue;
    const bRow = beforeData[br] || [];
    const aRow = afterData[ar];
    // 只關注新增的值（展後有、展前沒有），不報告被刪除或修改的
    for (let c = bCols; c < aRow.length && changes < maxChanges; c++) {
      const va = (aRow[c] || '').trim();
      if (va) {
        result += 'NEW_CELL ' + colToLetter(c) + (ar + 1) + '=' + va.slice(0, 40) + '\n';
        changes++;
      }
    }
  }

  return result;
}

// ═══ 頁籤相似度比對 ═══
export function xlSheetSimilarity(inputSheets, historySheets) {
  const result = {};
  const inputNames = Object.keys(inputSheets);
  const histNames = Object.keys(historySheets);

  for (const iName of inputNames) {
    let bestMatch = null, bestScore = 0, bestInfo = '';
    const iData = inputSheets[iName].data;
    const iCells = new Set();
    iData.slice(0, 60).forEach(row => {
      row.forEach(v => { if (v && v.trim().length > 1) iCells.add(v.trim()); });
    });
    if (iCells.size === 0) {
      result[iName] = { matchedSheet: null, similarity: 0, note: '空白頁籤' };
      continue;
    }

    for (const hName of histNames) {
      const hData = historySheets[hName].data;
      const hCells = new Set();
      hData.slice(0, 60).forEach(row => {
        row.forEach(v => { if (v && v.trim().length > 1) hCells.add(v.trim()); });
      });
      if (hCells.size === 0) continue;

      let intersection = 0;
      for (const v of iCells) { if (hCells.has(v)) intersection++; }
      const coverage = iCells.size > 0 ? intersection / iCells.size : 0;
      const jaccard = new Set([...iCells, ...hCells]).size > 0 ? intersection / new Set([...iCells, ...hCells]).size : 0;
      const score = jaccard * 0.4 + coverage * 0.6;

      if (score > bestScore) {
        bestScore = score;
        bestMatch = hName;
        bestInfo = `匹配${intersection}/${iCells.size}項 (${(coverage * 100).toFixed(0)}%)`;
      }
    }

    result[iName] = {
      matchedSheet: bestScore > 0.15 ? bestMatch : null,
      similarity: bestScore,
      note: bestScore > 0.15 ? bestInfo : '無足夠相似的歷史頁籤',
    };
  }
  return result;
}

// JSON 清理
export function sanitizeJsonString(raw) {
  const chars = [];
  let inStr = false, escaped = false;
  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i], code = raw.charCodeAt(i);
    if (escaped) { chars.push(ch); escaped = false; continue; }
    if (ch === '\\' && inStr) { chars.push(ch); escaped = true; continue; }
    if (ch === '"') { inStr = !inStr; chars.push(ch); continue; }
    if (inStr && code < 0x20) {
      if (ch === '\n') chars.push('\\', 'n');
      else if (ch === '\r') chars.push('\\', 'r');
      else if (ch === '\t') chars.push('\\', 't');
      continue;
    }
    chars.push(ch);
  }
  return chars.join('');
}

// 解析 AI JSON（強化版：處理超大被截斷的 JSON）
export function parseAiJson(raw) {
  let tx = raw.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
  const si = tx.indexOf('{');
  if (si < 0) throw new Error('AI回應中找不到JSON');
  tx = tx.slice(si);
  tx = sanitizeJsonString(tx).replace(/,\s*([}\]])/g, '$1');

  // 先找完整 JSON 物件
  let dep = 0, ei = -1;
  for (let i = 0; i < tx.length; i++) {
    const c = tx[i];
    if (c === '"') { i++; while (i < tx.length) { if (tx[i] === '\\') i++; else if (tx[i] === '"') break; i++; } continue; }
    if (c === '{') dep++; else if (c === '}') { dep--; if (dep === 0) { ei = i; break; } }
  }
  if (ei > 0) tx = tx.slice(0, ei + 1);
  tx = tx.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '');

  // 嘗試直接解析
  try { return JSON.parse(tx); } catch (e1) { /* fallthrough */ }

  // 基本修復嘗試
  let fixed = tx;
  let inS = false, esc = false;
  for (let i = 0; i < fixed.length; i++) {
    if (esc) { esc = false; continue; }
    if (fixed[i] === '\\') { esc = true; continue; }
    if (fixed[i] === '"') inS = !inS;
  }
  if (inS) fixed += '"';
  fixed = fixed.replace(/,\s*"[^"]*"?\s*:?\s*"?[^"{}[\]]*$/, '');
  fixed = fixed.replace(/,\s*$/, '');

  let brackets = 0, braces = 0;
  let inS2 = false, esc2 = false;
  for (let i = 0; i < fixed.length; i++) {
    if (esc2) { esc2 = false; continue; } if (fixed[i] === '\\') { esc2 = true; continue; }
    if (fixed[i] === '"') { inS2 = !inS2; continue; } if (inS2) continue;
    if (fixed[i] === '[') brackets++; else if (fixed[i] === ']') brackets--;
    if (fixed[i] === '{') braces++; else if (fixed[i] === '}') braces--;
  }
  for (let i = 0; i < brackets; i++) fixed += ']';
  for (let i = 0; i < braces; i++) fixed += '}';

  try { return JSON.parse(fixed); } catch (e2) { /* fallthrough to extraction */ }

  // ═══ 最後手段：從截斷的 JSON 中逐欄位提取 ═══
  console.warn('JSON 整體解析失敗，嘗試逐欄位提取...');
  const result = { analysis: '(JSON被截斷，已部分提取)', cellUpdates: [], rowInsertions: [], sheetMapping: [] };

  // 提取 analysis
  const analysisMatch = tx.match(/"analysis"\s*:\s*"([^"]*?)"/);
  if (analysisMatch) result.analysis = analysisMatch[1];

  // 提取 cellUpdates 陣列中的每個物件
  const cuStart = tx.indexOf('"cellUpdates"');
  if (cuStart >= 0) {
    const arrStart = tx.indexOf('[', cuStart);
    if (arrStart >= 0) {
      // 逐一提取 {sheet, cell, value} 物件
      const itemRegex = /\{\s*"sheet"\s*:\s*"([^"]*)"\s*,\s*"cell"\s*:\s*"([^"]*)"\s*,\s*"value"\s*:\s*"([^"]*)"\s*\}/g;
      // 也匹配不同欄位順序
      const itemRegex2 = /\{\s*"cell"\s*:\s*"([^"]*)"\s*,\s*"value"\s*:\s*"([^"]*)"\s*,\s*"sheet"\s*:\s*"([^"]*)"\s*\}/g;
      let m;
      while ((m = itemRegex.exec(tx)) !== null) {
        result.cellUpdates.push({ sheet: m[1], cell: m[2], value: m[3] });
      }
      while ((m = itemRegex2.exec(tx)) !== null) {
        result.cellUpdates.push({ sheet: m[3], cell: m[1], value: m[2] });
      }
    }
  }

  // 提取 rowInsertions
  const riStart = tx.indexOf('"rowInsertions"');
  if (riStart >= 0) {
    const riRegex = /\{\s*"sheet"\s*:\s*"([^"]*)"\s*,\s*"afterRow"\s*:\s*(\d+)\s*,\s*"cells"\s*:\s*(\{[^}]*\})/g;
    let m;
    while ((m = riRegex.exec(tx)) !== null) {
      try {
        const cells = JSON.parse(m[3]);
        result.rowInsertions.push({ sheet: m[1], afterRow: parseInt(m[2]), cells });
      } catch (_) { /* skip malformed */ }
    }
  }

  // 提取 sheetMapping
  const smRegex = /\{\s*"inputSheet"\s*:\s*"([^"]*)"\s*,\s*"action"\s*:\s*"([^"]*)"\s*,\s*"reason"\s*:\s*"([^"]*)"\s*\}/g;
  let sm;
  while ((sm = smRegex.exec(tx)) !== null) {
    result.sheetMapping.push({ inputSheet: sm[1], action: sm[2], reason: sm[3] });
  }

  if (result.cellUpdates.length === 0 && result.rowInsertions.length === 0) {
    throw new Error('JSON解析失敗且無法提取任何更新。回應長度: ' + raw.length + ' 字元。建議換用 gemini-2.5-flash 模型。');
  }

  return result;
}

