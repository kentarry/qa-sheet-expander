import JSZip from 'jszip';
import { xlParse } from './excel.js';

// ═══ 展表前/展表後 自動配對 ═══
export function pairFiles(files) {
  const pairs = [];
  const used = new Set();
  const beforeFiles = [];
  const afterFiles = [];
  const otherFiles = [];

  for (const f of files) {
    const name = f.name;
    if (/展表前/.test(name)) {
      beforeFiles.push(f);
    } else if (/展表後/.test(name)) {
      afterFiles.push(f);
    } else if (/品檢/.test(name)) {
      afterFiles.push(f);
    } else {
      otherFiles.push(f);
    }
  }

  function coreName(name) {
    return name
      .replace(/\[品檢\]\s*/g, '')
      .replace(/[_\s]*(展表前|展表後|品檢單?|品檢)[^.]*/gi, '')
      .replace(/\.(xlsx?|xls)$/i, '')
      .replace(/v\d+(\.\d+)?/gi, '')
      .replace(/[_\s]+/g, ' ')
      .trim();
  }

  for (const after of afterFiles) {
    const aCore = coreName(after.name);
    let bestMatch = null, bestScore = 0;
    for (const before of beforeFiles) {
      if (used.has(before.id)) continue;
      const bCore = coreName(before.name);
      let score = 0;
      const shorter = Math.min(aCore.length, bCore.length);
      for (let i = 0; i < shorter; i++) {
        if (aCore[i] === bCore[i]) score++; else break;
      }
      const aWords = aCore.split(/\s+/).filter(w => w.length > 1);
      const bWords = bCore.split(/\s+/).filter(w => w.length > 1);
      aWords.forEach(w => { if (bWords.some(bw => bw.includes(w) || w.includes(bw))) score += 5; });
      if (score > bestScore) { bestScore = score; bestMatch = before; }
    }
    if (bestMatch && bestScore > 3) {
      pairs.push({ before: bestMatch, after, score: bestScore });
      used.add(bestMatch.id); used.add(after.id);
    } else {
      pairs.push({ before: null, after, score: 0 });
      used.add(after.id);
    }
  }
  for (const f of [...beforeFiles, ...otherFiles]) {
    if (!used.has(f.id)) pairs.push({ before: f, after: null, score: 0 });
  }
  return pairs;
}

// ═══ 欄位字母 ↔ 索引 ═══
function colLetterToIndex(col) {
  let idx = 0;
  for (let i = 0; i < col.length; i++) {
    idx = idx * 26 + col.charCodeAt(i) - 64;
  }
  return idx - 1; // 0-based
}

function colIndexToLetter(idx) {
  let s = '';
  let n = idx;
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

function cellToRC(addr) {
  const m = addr.match(/^([A-Z]+)(\d+)$/);
  if (!m) return null;
  return { r: parseInt(m[2]), c: colLetterToIndex(m[1]) };
}

function rcToCell(r, c) {
  return colIndexToLetter(c) + r;
}

// ═══ 解析 workbook.xml 取得 sheet 名稱與 rId 映射 ═══
function parseWorkbook(xml) {
  const sheets = [];
  const regex = /<sheet\s+name="([^"]+)"\s+sheetId="(\d+)"\s+r:id="([^"]+)"/g;
  let m;
  while ((m = regex.exec(xml)) !== null) {
    sheets.push({ name: m[1], sheetId: m[2], rId: m[3] });
  }
  return sheets;
}

// ═══ 解析 workbook.xml.rels ═══
function parseRels(xml) {
  const rels = {};
  const regex = /<Relationship\s+[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g;
  let m;
  while ((m = regex.exec(xml)) !== null) {
    rels[m[1]] = m[2];
  }
  const regex2 = /<Relationship\s+[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"/g;
  while ((m = regex2.exec(xml)) !== null) {
    if (!rels[m[2]]) rels[m[2]] = m[1];
  }
  return rels;
}

// ═══ DOMParser 解析 sheet XML ═══
function parseSheetXml(xmlStr) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlStr, 'application/xml');
  return doc;
}

function serializeXml(doc) {
  const serializer = new XMLSerializer();
  let xml = serializer.serializeToString(doc);
  // 確保 XML 聲明存在
  if (!xml.startsWith('<?xml')) {
    xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + xml;
  }
  return xml;
}

// ═══ 從 XML 取得所有 row 和 cell 的結構 ═══
function getSheetData(doc) {
  const rows = [];
  const ns = doc.documentElement.namespaceURI;
  const rowEls = doc.getElementsByTagNameNS(ns, 'row');
  for (let i = 0; i < rowEls.length; i++) {
    const rowEl = rowEls[i];
    const rowNum = parseInt(rowEl.getAttribute('r'));
    const cells = [];
    const cellEls = rowEl.getElementsByTagNameNS(ns, 'c');
    for (let j = 0; j < cellEls.length; j++) {
      const cellEl = cellEls[j];
      const ref = cellEl.getAttribute('r');
      const rc = cellToRC(ref);
      if (rc) cells.push({ ref, r: rc.r, c: rc.c, el: cellEl });
    }
    rows.push({ rowNum, el: rowEl, cells });
  }
  return rows;
}

// ═══ 在 row 中插入新的 cell，保持欄位順序 ═══
function insertCellInRow(doc, rowEl, colIndex, rowNum, value, copyStyleFromRow) {
  const ns = doc.documentElement.namespaceURI;
  const ref = rcToCell(rowNum, colIndex);

  // 檢查是否已存在此 cell（不覆蓋已有的 cell）
  const existingCells = rowEl.getElementsByTagNameNS(ns, 'c');
  for (let i = 0; i < existingCells.length; i++) {
    const existRef = existingCells[i].getAttribute('r');
    const existRC = cellToRC(existRef);
    if (existRC && existRC.c === colIndex) {
      return false; // 此格已有資料，不覆蓋
    }
  }

  // 建立新 cell
  const newCell = doc.createElementNS(ns, 'c');
  newCell.setAttribute('r', ref);

  // 複製鄰近 cell 的 style
  if (copyStyleFromRow) {
    const adjacentCells = rowEl.getElementsByTagNameNS(ns, 'c');
    if (adjacentCells.length > 0) {
      const lastCell = adjacentCells[adjacentCells.length - 1];
      const style = lastCell.getAttribute('s');
      if (style) newCell.setAttribute('s', style);
    }
  }

  if (value && value !== '') {
    newCell.setAttribute('t', 'inlineStr');
    const is = doc.createElementNS(ns, 'is');
    const t = doc.createElementNS(ns, 't');
    t.textContent = value;
    is.appendChild(t);
    newCell.appendChild(is);
  }

  // 找到正確的插入位置（保持欄位排序）
  let insertBefore = null;
  for (let i = 0; i < existingCells.length; i++) {
    const existRef = existingCells[i].getAttribute('r');
    const existRC = cellToRC(existRef);
    if (existRC && existRC.c > colIndex) {
      insertBefore = existingCells[i];
      break;
    }
  }

  if (insertBefore) {
    rowEl.insertBefore(newCell, insertBefore);
  } else {
    rowEl.appendChild(newCell);
  }

  return true;
}

// ═══ 插入新行 ═══
function insertRowAfter(doc, sheetDataEl, afterRowNum, cells) {
  const ns = doc.documentElement.namespaceURI;
  const newRowNum = afterRowNum + 1;

  // 先把所有 >= newRowNum 的 row 往下移一格
  const rows = sheetDataEl.getElementsByTagNameNS(ns, 'row');
  const rowsToShift = [];
  for (let i = 0; i < rows.length; i++) {
    const rn = parseInt(rows[i].getAttribute('r'));
    if (rn >= newRowNum) rowsToShift.push(rows[i]);
  }

  // 從最大的 row 開始往下移
  rowsToShift.sort((a, b) => parseInt(b.getAttribute('r')) - parseInt(a.getAttribute('r')));
  for (const rowEl of rowsToShift) {
    const oldNum = parseInt(rowEl.getAttribute('r'));
    const shifted = oldNum + 1;
    rowEl.setAttribute('r', String(shifted));
    // 更新該 row 下所有 cell 的 ref
    const cellEls = rowEl.getElementsByTagNameNS(ns, 'c');
    for (let j = 0; j < cellEls.length; j++) {
      const ref = cellEls[j].getAttribute('r');
      const rc = cellToRC(ref);
      if (rc) {
        cellEls[j].setAttribute('r', rcToCell(shifted, rc.c));
      }
    }
  }

  // 建立新行
  const newRowEl = doc.createElementNS(ns, 'row');
  newRowEl.setAttribute('r', String(newRowNum));

  // 從 afterRowNum 複製 span 等屬性
  let afterRowEl = null;
  const allRows = sheetDataEl.getElementsByTagNameNS(ns, 'row');
  for (let i = 0; i < allRows.length; i++) {
    // 注意：afterRowNum 的 row 現在可能已經被移到 afterRowNum+1
    // 所以要找原始的 afterRowNum (如果沒被移動) 或 afterRowNum+1 (如果被移動)
    const rn = parseInt(allRows[i].getAttribute('r'));
    if (rn === afterRowNum || rn === afterRowNum + 1) {
      afterRowEl = allRows[i];
      if (rn === afterRowNum) break; // 找到未被移動的原始行
    }
  }

  // 複製 spans 和 style 屬性
  if (afterRowEl) {
    const spans = afterRowEl.getAttribute('spans');
    if (spans) newRowEl.setAttribute('spans', spans);
    const ht = afterRowEl.getAttribute('ht');
    if (ht) newRowEl.setAttribute('ht', ht);
    const customHeight = afterRowEl.getAttribute('customHeight');
    if (customHeight) newRowEl.setAttribute('customHeight', customHeight);
  }

  // 填入 cells
  for (const [colStr, val] of Object.entries(cells)) {
    const colIdx = colLetterToIndex(colStr.toUpperCase());
    const ref = rcToCell(newRowNum, colIdx);
    const cell = doc.createElementNS(ns, 'c');
    cell.setAttribute('r', ref);

    // 從 afterRowEl 找同欄的 style
    if (afterRowEl) {
      const afterCells = afterRowEl.getElementsByTagNameNS(ns, 'c');
      for (let j = 0; j < afterCells.length; j++) {
        const aRef = afterCells[j].getAttribute('r');
        const arc = cellToRC(aRef);
        if (arc && arc.c === colIdx) {
          const s = afterCells[j].getAttribute('s');
          if (s) cell.setAttribute('s', s);
          break;
        }
      }
    }

    if (val && val !== '') {
      cell.setAttribute('t', 'inlineStr');
      const is = doc.createElementNS(ns, 'is');
      const t = doc.createElementNS(ns, 't');
      t.textContent = val;
      is.appendChild(t);
      cell.appendChild(is);
    }

    newRowEl.appendChild(cell);
  }

  // 找插入位置
  let insertBefore = null;
  const sortedRows = [];
  for (let i = 0; i < allRows.length; i++) {
    sortedRows.push(allRows[i]);
  }
  for (const r of sortedRows) {
    if (parseInt(r.getAttribute('r')) > newRowNum) {
      insertBefore = r;
      break;
    }
  }

  if (insertBefore) {
    sheetDataEl.insertBefore(newRowEl, insertBefore);
  } else {
    sheetDataEl.appendChild(newRowEl);
  }

  return true;
}

// ═══ 更新 dimension ref ═══
function updateDimension(doc) {
  const ns = doc.documentElement.namespaceURI;
  const dimEl = doc.getElementsByTagNameNS(ns, 'dimension')[0];
  if (!dimEl) return;

  const rows = doc.getElementsByTagNameNS(ns, 'row');
  if (rows.length === 0) return;

  let maxRow = 0, maxCol = 0, minRow = Infinity, minCol = Infinity;
  for (let i = 0; i < rows.length; i++) {
    const rn = parseInt(rows[i].getAttribute('r'));
    if (rn > maxRow) maxRow = rn;
    if (rn < minRow) minRow = rn;
    const cells = rows[i].getElementsByTagNameNS(ns, 'c');
    for (let j = 0; j < cells.length; j++) {
      const ref = cells[j].getAttribute('r');
      const rc = cellToRC(ref);
      if (rc) {
        if (rc.c > maxCol) maxCol = rc.c;
        if (rc.c < minCol) minCol = rc.c;
      }
    }
  }

  const newRef = rcToCell(minRow, minCol) + ':' + rcToCell(maxRow, maxCol);
  dimEl.setAttribute('ref', newRef);
}

// ═══ 更新 mergeCells 中的行號（因為插入行導致偏移）═══
function shiftMergeCells(doc, insertedRowNum) {
  const ns = doc.documentElement.nameSpakerURI || doc.documentElement.namespaceURI;
  const mergeCells = doc.getElementsByTagNameNS(ns, 'mergeCell');
  for (let i = 0; i < mergeCells.length; i++) {
    const ref = mergeCells[i].getAttribute('ref');
    if (!ref) continue;
    const parts = ref.split(':');
    if (parts.length !== 2) continue;
    const rc1 = cellToRC(parts[0]);
    const rc2 = cellToRC(parts[1]);
    if (!rc1 || !rc2) continue;

    let changed = false;
    if (rc1.r >= insertedRowNum) { rc1.r++; changed = true; }
    if (rc2.r >= insertedRowNum) { rc2.r++; changed = true; }
    if (changed) {
      mergeCells[i].setAttribute('ref', rcToCell(rc1.r, rc1.c) + ':' + rcToCell(rc2.r, rc2.c));
    }
  }
}

// ═══ 主要函數：只做插入，不修改原有資料 ═══
export async function xlBuild(origBuf, plan) {
  const zip = await JSZip.loadAsync(origBuf);

  // 讀取 workbook 和 rels
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  if (!workbookXml) throw new Error('找不到 workbook.xml');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  if (!relsXml) throw new Error('找不到 workbook.xml.rels');

  const wbSheets = parseWorkbook(workbookXml);
  const rels = parseRels(relsXml);

  // 建立 sheetName → XML 路徑映射
  const sheetPathMap = {};
  for (const s of wbSheets) {
    const target = rels[s.rId];
    if (target) {
      const fullPath = target.startsWith('/') ? target.slice(1) : 'xl/' + target;
      sheetPathMap[s.name] = fullPath;
    }
  }

  // ═══ 處理 AI 計畫 ═══
  const columnInsertions = plan.columnInsertions || [];
  const rowInsertions = plan.rowInsertions || [];
  // 向後相容舊格式的 cellUpdates（只處理新增 cell，不覆蓋）
  const cellUpdates = plan.cellUpdates || [];

  // 按 sheet 分組
  const colInsBySheet = {};
  for (const ci of columnInsertions) {
    const sn = ci.sheet || wbSheets[0]?.name || '';
    if (!colInsBySheet[sn]) colInsBySheet[sn] = [];
    colInsBySheet[sn].push(ci);
  }

  const rowInsBySheet = {};
  for (const ri of rowInsertions) {
    const sn = ri.sheet || wbSheets[0]?.name || '';
    if (!rowInsBySheet[sn]) rowInsBySheet[sn] = [];
    rowInsBySheet[sn].push(ri);
  }

  const cellUpBySheet = {};
  for (const cu of cellUpdates) {
    if (!cu.cell || cu.value === undefined) continue;
    const sn = cu.sheet || wbSheets[0]?.name || '';
    if (!cellUpBySheet[sn]) cellUpBySheet[sn] = [];
    cellUpBySheet[sn].push(cu);
  }

  let applied = 0, skipped = 0;
  const processedSheets = [];

  const allSheetNames = new Set([
    ...Object.keys(colInsBySheet),
    ...Object.keys(rowInsBySheet),
    ...Object.keys(cellUpBySheet),
  ]);

  for (const sheetName of allSheetNames) {
    let xmlPath = sheetPathMap[sheetName];

    // 模糊匹配
    if (!xmlPath) {
      const fuzzyMatch = Object.entries(sheetPathMap).find(([name]) =>
        name.includes(sheetName) || sheetName.includes(name)
      );
      if (fuzzyMatch) xmlPath = fuzzyMatch[1];
    }
    if (!xmlPath) { skipped++; continue; }

    const sheetXmlStr = await zip.file(xmlPath)?.async('string');
    if (!sheetXmlStr) { skipped++; continue; }

    const doc = parseSheetXml(sheetXmlStr);
    const ns = doc.documentElement.namespaceURI;
    const sheetDataEl = doc.getElementsByTagNameNS(ns, 'sheetData')[0];
    if (!sheetDataEl) { skipped++; continue; }

    // ── 1. 處理 column insertions（在每行的指定位置插入新 cell）──
    const colIns = colInsBySheet[sheetName] || [];
    for (const ci of colIns) {
      const afterColIdx = colLetterToIndex(ci.afterColumn.toUpperCase());
      const newColIdx = afterColIdx + 1;

      // 先把所有 > afterColIdx 的 cell 往右移
      const allRows = sheetDataEl.getElementsByTagNameNS(ns, 'row');
      for (let ri = 0; ri < allRows.length; ri++) {
        const rowEl = allRows[ri];
        const rowNum = parseInt(rowEl.getAttribute('r'));
        const cellEls = rowEl.getElementsByTagNameNS(ns, 'c');

        // 從右到左移動 cell，避免衝突
        const cellsToShift = [];
        for (let ci2 = 0; ci2 < cellEls.length; ci2++) {
          const ref = cellEls[ci2].getAttribute('r');
          const rc = cellToRC(ref);
          if (rc && rc.c > afterColIdx) {
            cellsToShift.push({ el: cellEls[ci2], rc });
          }
        }
        cellsToShift.sort((a, b) => b.rc.c - a.rc.c);
        for (const item of cellsToShift) {
          item.el.setAttribute('r', rcToCell(rowNum, item.rc.c + 1));
        }
      }

      // 在 headers 指定的行填入值
      const headers = ci.headers || [];
      for (const h of headers) {
        const targetRow = h.row;
        // 找到目標行
        let targetRowEl = null;
        const rEls = sheetDataEl.getElementsByTagNameNS(ns, 'row');
        for (let ri = 0; ri < rEls.length; ri++) {
          if (parseInt(rEls[ri].getAttribute('r')) === targetRow) {
            targetRowEl = rEls[ri];
            break;
          }
        }
        if (targetRowEl) {
          insertCellInRow(doc, targetRowEl, newColIdx, targetRow, h.value, true);
          applied++;
        }
      }

      // 在所有沒有指定 header 的行也加入空 cell（保持結構一致）
      const headerRows = new Set(headers.map(h => h.row));
      const allRows2 = sheetDataEl.getElementsByTagNameNS(ns, 'row');
      for (let ri = 0; ri < allRows2.length; ri++) {
        const rowEl = allRows2[ri];
        const rowNum = parseInt(rowEl.getAttribute('r'));
        if (!headerRows.has(rowNum)) {
          insertCellInRow(doc, rowEl, newColIdx, rowNum, '', true);
        }
      }
    }

    // ── 2. 處理 row insertions ──
    const riList = (rowInsBySheet[sheetName] || []).slice();
    // 從最大的 afterRow 開始插入，避免偏移問題
    riList.sort((a, b) => (b.afterRow || 0) - (a.afterRow || 0));
    for (const ins of riList) {
      if (!ins.afterRow || !ins.cells) continue;
      const success = insertRowAfter(doc, sheetDataEl, ins.afterRow, ins.cells);
      if (success) applied++;
    }

    // ── 3. 處理 cellUpdates（只在空 cell 填入值，不覆蓋已有資料）──
    const cuList = cellUpBySheet[sheetName] || [];
    for (const cu of cuList) {
      const ref = cu.cell.toUpperCase();
      const rc = cellToRC(ref);
      if (!rc) continue;

      // 找到對應的 row
      const allRows = sheetDataEl.getElementsByTagNameNS(ns, 'row');
      let targetRowEl = null;
      for (let ri = 0; ri < allRows.length; ri++) {
        if (parseInt(allRows[ri].getAttribute('r')) === rc.r) {
          targetRowEl = allRows[ri];
          break;
        }
      }

      if (!targetRowEl) continue;

      // 檢查 cell 是否已經有值
      const existingCells = targetRowEl.getElementsByTagNameNS(ns, 'c');
      let cellExists = false;
      let cellHasValue = false;

      for (let ci2 = 0; ci2 < existingCells.length; ci2++) {
        const existRef = existingCells[ci2].getAttribute('r');
        if (existRef === ref) {
          cellExists = true;
          // 檢查是否有實際內容
          const vEl = existingCells[ci2].getElementsByTagNameNS(ns, 'v')[0];
          const isEl = existingCells[ci2].getElementsByTagNameNS(ns, 'is')[0];
          if ((vEl && vEl.textContent) || isEl) {
            cellHasValue = true;
          }
          break;
        }
      }

      // 只有在 cell 不存在或 cell 存在但沒有值時才填入
      if (!cellHasValue) {
        if (cellExists) {
          // cell 存在但空 → 向已存在的 cell 添加值
          for (let ci2 = 0; ci2 < existingCells.length; ci2++) {
            if (existingCells[ci2].getAttribute('r') === ref && cu.value) {
              existingCells[ci2].setAttribute('t', 'inlineStr');
              const is = doc.createElementNS(ns, 'is');
              const t = doc.createElementNS(ns, 't');
              t.textContent = cu.value;
              is.appendChild(t);
              existingCells[ci2].appendChild(is);
              applied++;
              break;
            }
          }
        } else {
          // cell 不存在 → 插入新 cell
          const success = insertCellInRow(doc, targetRowEl, rc.c, rc.r, cu.value, true);
          if (success) applied++;
        }
      }
    }

    // ── 4. 更新 dimension ──
    updateDimension(doc);

    // 寫回 ZIP
    const updatedXml = serializeXml(doc);
    zip.file(xmlPath, updatedXml);
    processedSheets.push(sheetName);
  }

  // 產生 blob
  const out = await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  // 預覽用
  const previewBuf = await out.arrayBuffer();
  const previewParsed = xlParse(new Uint8Array(previewBuf));

  return {
    es: previewParsed.sheets,
    blob: out,
    applied,
    skipped,
    sheetCount: previewParsed.sheetNames.length,
    processedSheets,
  };
}
