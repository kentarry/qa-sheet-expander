import { useState, useRef, useEffect, useCallback } from 'react';
import { DRIVE_ROOT_ID, MODELS } from './config.js';
import { driveListFolder, driveDownloadFile } from './drive.js';
import { gemInit, gemCall } from './gemini.js';
import { xlParse, xlSummary, xlSignature, xlDiff, xlSheetSimilarity, parseAiJson } from './excel.js';
import { pairFiles, xlBuild } from './xlBuild.js';

// ═══ Sub Components ═══

function StepBar({ steps, current }) {
  return (
    <div className="step-bar">
      {steps.map((s, i) => {
        const state = i < current ? 'done' : i === current ? 'active' : 'pending';
        return (
          <div key={i} className="step-bar__item">
            <div className={`step-bar__track step-bar__track--${state}`} />
            <div className={`step-bar__label step-bar__label--${state}`}>
              {i < current ? '✓ ' : ''}{s}
            </div>
          </div>
        );
      })}
    </div>
  );
}

function Logs({ logs }) {
  const ref = useRef(null);
  useEffect(() => { ref.current?.scrollIntoView({ behavior: 'smooth' }); }, [logs.length]);
  return (
    <div className="log-panel">
      {logs.map((l, i) => (
        <div key={i} className={`log-line log-line--${l.t === 'e' ? 'err' : l.t === 'o' ? 'ok' : 'info'}`}>
          <span className="log-line__ts">{l.ts}</span>
          <span>{l.m}</span>
        </div>
      ))}
      <div ref={ref} />
    </div>
  );
}

function SheetPreview({ sheets, title, activeTab, onTabChange }) {
  const names = Object.keys(sheets);
  const active = activeTab && sheets[activeTab] ? activeTab : names[0];
  const sd = sheets[active];
  if (!sd) return null;
  return (
    <div>
      {title && <div className="sheet-preview__title">{title}</div>}
      <div className="sheet-tabs">
        {names.map(n => (
          <button key={n} onClick={() => onTabChange(n)} className={`sheet-tab ${n === active ? 'sheet-tab--active' : ''}`}>{n}</button>
        ))}
      </div>
      <div className="sheet-table-wrap">
        <table className="sheet-table">
          <tbody>
            {sd.data.slice(0, 50).map((row, ri) => (
              <tr key={ri}>{row.map((c, ci) => (
                <td key={ci}>{(c || '').slice(0, 50)}</td>
              ))}</tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function SimilarityBadge({ similarity }) {
  if (!similarity) return null;
  const entries = Object.entries(similarity);
  return (
    <div>
      <div className="similarity-section__title">📊 頁籤相似度比對</div>
      <div className="similarity-grid">
        {entries.map(([name, info]) => {
          const matched = info.matchedSheet;
          const pct = Math.round(info.similarity * 100);
          return (
            <div key={name} className={`similarity-item ${matched ? 'similarity-item--matched' : 'similarity-item--unmatched'}`}>
              <div>
                <div className="similarity-item__name">{name}</div>
                <div className={matched ? 'similarity-item__note--matched' : 'similarity-item__note--unmatched'}>{info.note}</div>
              </div>
              <div className={`similarity-item__pct ${matched ? 'similarity-item__pct--matched' : 'similarity-item__pct--unmatched'}`}>
                {matched ? '✓ ' + pct + '%' : '—'}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

const Spinner = () => <span className="spinner" />;

// ═══ Settings Modal ═══
function SettingsModal({ gemKey, gemModel, onSave, onClose }) {
  const keyRef = useRef(null);
  const modelRef = useRef(null);

  const handleSave = () => {
    const k = (keyRef.current?.value || '').trim();
    if (!k) return;
    const m = modelRef.current?.value || 'gemini-2.0-flash';
    onSave(k, m);
  };

  return (
    <div className="modal-overlay" onClick={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="modal">
        <h3 className="modal__title">⚙️ 設定</h3>
        <div className="modal__group">
          <label className="modal__label">🔑 Gemini API Key</label>
          <input ref={keyRef} className="modal__input" defaultValue={gemKey} placeholder="AIzaSy..." spellCheck={false} />
          <a href="https://aistudio.google.com/apikey" target="_blank" rel="noopener" className="modal__link">🔗 前往取得 Key ↗</a>
        </div>
        <div className="modal__group">
          <label className="modal__label">🤖 AI 模型</label>
          <select ref={modelRef} className="modal__select" defaultValue={gemModel}>
            {MODELS.map(m => <option key={m.id} value={m.id}>{m.name} — {m.desc}</option>)}
          </select>
          <div className="modal__hint">
            • 2.0 Flash：免費 1,500 次/天，推薦日常使用<br />
            • 2.5 Flash：免費 20 次/天，品質更好<br />
            • 2.5 Pro：免費 5 次/天，最強但最少
          </div>
        </div>
        <div className="modal__actions">
          <button onClick={handleSave} className="btn btn--primary">儲存</button>
          <button onClick={onClose} className="btn btn--ghost">關閉</button>
        </div>
      </div>
    </div>
  );
}

// ═══ MAIN APP ═══
export default function App() {
  const [step, setStep] = useState(0);
  const [logs, setLogs] = useState([]);
  const [showCfg, setShowCfg] = useState(false);
  const [pvTab, setPvTab] = useState({});

  // Drive navigation
  const [projects, setProjects] = useState(null);
  const [cats, setCats] = useState(null);
  const [files, setFiles] = useState(null);
  const [pairs, setPairs] = useState(null);
  const [projId, setProjId] = useState('');
  const [projName, setProjName] = useState('');
  const [catId, setCatId] = useState('');
  const [catName, setCatName] = useState('');
  const [lProj, setLProj] = useState(false);
  const [lCat, setLCat] = useState(false);
  const [lFile, setLFile] = useState(false);

  // Input / Output
  const [inputFile, setInputFile] = useState(null);
  const [inputParsed, setInputParsed] = useState(null);
  const [inputBuf, setInputBuf] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [expanded, setExpanded] = useState(null);
  const [blob, setBlob] = useState(null);
  const [matchInfo, setMatchInfo] = useState('');
  const [similarity, setSimilarity] = useState(null);

  // Settings
  const [gemKey, setGemKey] = useState(() => localStorage.getItem('qse_key') || '');
  const [gemModel, setGemModel] = useState(() => localStorage.getItem('qse_model') || 'gemini-2.0-flash');

  if (gemKey) gemInit(gemKey);

  const log = useCallback((m, t) => {
    const d = new Date();
    const ts = [d.getHours(), d.getMinutes(), d.getSeconds()].map(v => String(v).padStart(2, '0')).join(':');
    setLogs(p => [...p, { ts, m, t: t || 'i' }]);
  }, []);

  // ── Drive navigation ──
  useEffect(() => {
    (async () => {
      setLProj(true);
      try {
        const d = await driveListFolder(DRIVE_ROOT_ID);
        setProjects(d.folders);
      } catch (e) { setLogs(p => [...p, { ts: '', m: e.message, t: 'e' }]); }
      finally { setLProj(false); }
    })();
  }, []);

  const handleSelProj = async (id, name) => {
    setProjId(id); setProjName(name); setCatId(''); setCatName(''); setCats(null); setFiles(null); setPairs(null); setStep(0); setSimilarity(null);
    setLCat(true);
    try { const d = await driveListFolder(id); setCats(d.folders); log(name + ': ' + d.folders.length + ' 個類型', 'o'); }
    catch (e) { log(e.message, 'e'); }
    finally { setLCat(false); }
  };

  const handleSelCat = async (id, name) => {
    setCatId(id); setCatName(name); setFiles(null); setPairs(null); setStep(0); setSimilarity(null);
    setLFile(true);
    try {
      const d = await driveListFolder(id);
      setFiles(d.files);
      const p = pairFiles(d.files);
      setPairs(p);
      log(projName + '/' + name + ': ' + d.files.length + ' 檔案, ' + p.filter(x => x.before && x.after).length + ' 組配對', 'o');
    } catch (e) { log(e.message, 'e'); }
    finally { setLFile(false); }
  };

  // ── Input ──
  const handleInput = async (file) => {
    setInputFile(file); log('載入: ' + file.name);
    try {
      const buf = await file.arrayBuffer();
      const u8 = new Uint8Array(buf);
      setInputBuf(u8);
      const parsed = xlParse(u8);
      setInputParsed(parsed);
      log('OK ' + parsed.sheetNames.length + ' sheets (' + parsed.sheetNames.join(', ') + ')', 'o');
      setStep(1);
      if (pairs?.length) {
        const completePairs = pairs.filter(p => p.before && p.after);
        if (completePairs.length) {
          log('找到 ' + completePairs.length + ' 組歷史配對，將全部下載比對', 'o');
          setStep(2);
        } else { log('沒有完整的展前/展後配對', 'e'); }
      }
    } catch (e) { log('FAIL: ' + e.message, 'e'); }
  };

  // ── AI Expand ──
  const handleExpand = async () => {
    if (!inputParsed || !inputBuf) return;
    const completePairs = (pairs || []).filter(p => p.before && p.after);
    if (!completePairs.length) { log('沒有完整的歷史配對可用', 'e'); return; }
    if (!gemKey) { setShowCfg(true); return; }
    gemInit(gemKey);
    setProcessing(true); setStep(3);

    try {
      log('📥 並行下載所有歷史配對 (' + completePairs.length + ' 組)...');

      const downloadTasks = completePairs.map((pair, pi) => {
        const label = (pi + 1) + '/' + completePairs.length;
        return (async () => {
          const [r1, r2] = await Promise.all([
            driveDownloadFile(pair.before.id),
            driveDownloadFile(pair.after.id),
          ]);
          return { pair, label, before: xlParse(r1.buf), after: xlParse(r2.buf) };
        })();
      });

      const results = await Promise.allSettled(downloadTasks);

      let allDiffs = '';
      let allSimilarities = {};
      let downloadedCount = 0;

      for (const r of results) {
        if (r.status === 'fulfilled') {
          const { pair, label, before, after } = r.value;
          const diff = xlDiff(before.sheets, after.sheets);
          const diffLines = diff.split('\n').filter(l => l.trim()).length;
          log('[' + label + '] ✓ ' + pair.before.name + ' (' + diffLines + ' 項差異)', 'o');

          allDiffs += '\n======= 配對 ' + label + ': ' + pair.before.name + ' → ' + pair.after.name + ' =======\n' + diff;

          const sim = xlSheetSimilarity(inputParsed.sheets, before.sheets);
          Object.entries(sim).forEach(([sheetName, info]) => {
            if (!allSimilarities[sheetName] || info.similarity > allSimilarities[sheetName].similarity) {
              allSimilarities[sheetName] = { ...info, fromPair: pair.before.name };
            }
          });
          downloadedCount++;
        } else {
          log('❌ 下載失敗: ' + r.reason?.message, 'e');
        }
      }

      if (downloadedCount === 0) {
        log('❌ 所有歷史配對下載失敗，程序停止', 'e');
        setStep(2); setProcessing(false); return;
      }

      log('✅ 成功 ' + downloadedCount + '/' + completePairs.length + ' 組', 'o');
      setSimilarity(allSimilarities);
      const matchedCount = Object.values(allSimilarities).filter(s => s.matchedSheet).length;
      log('頁籤配對: ' + matchedCount + '/' + Object.keys(allSimilarities).length + ' 個頁籤有歷史對應', 'o');

      const currentModel = localStorage.getItem('qse_model') || 'gemini-2.0-flash';
      log('🤖 呼叫 Gemini AI (' + currentModel + ')...');

      let simSummary = '';
      Object.entries(allSimilarities).forEach(([name, info]) => {
        simSummary += '  ' + name + ' → ' + (info.matchedSheet ? '歷史:' + info.matchedSheet + ' (' + info.note + ')' : '無對應歷史（不展開）') + '\n';
      });

      const maxDiffLen = 12000;
      const trimmedDiffs = allDiffs.length > maxDiffLen
        ? allDiffs.slice(0, maxDiffLen) + '\n...(已截斷，以上為代表性差異)'
        : allDiffs;

      const prompt = `分析 ${downloadedCount} 組歷史展前→展後的「結構性差異」，歸納展開模式，套用到新檢驗單。

重要：只能用 columnInsertions 和 rowInsertions 來展開，**絕對不能修改或刪除新檢驗單中已有的任何資料**。

=== 頁籤對應 ===
${simSummary}

=== 歷史差異（展前→展後的結構變化）===
${trimmedDiffs}

=== 新檢驗單（你需要展開的目標）===
${xlSummary(inputParsed.sheets, 50)}

指示：
1. 分析歷史差異中的「INSERTED_ROW」和「新增欄位」，找出展開模式
2. **根據新檢驗單中測試項目的實際文字內容和行號來定位**，不要假設位置和歷史一樣
3. 用 columnInsertions 來新增欄位（afterColumn = 在哪欄之後插入）
4. 用 rowInsertions 來新增行（afterRow = 在新檢驗單的哪行之後插入）
5. 驗證欄位留空字串 ""，只有標題填文字
6. 每個有歷史對應的頁籤都必須展開
7. 展開的詳細度要和歷史一致（如果歷史加了N行，新檢驗單對應位置也要加N行）
8. 不要新增新檢驗單中不存在的測試分類`;

      log('Prompt 大小: ' + (prompt.length / 1024).toFixed(1) + ' KB');

      let tx;
      try {
        tx = await gemCall(prompt, log);
      } catch (e) {
        log(e.message, 'e');
        log('程序已停止。請檢查 API Key 或稍後再試。', 'e');
        setStep(2);
        setProcessing(false);
        return;
      }

      log('AI 回應 (' + tx.length + ' chars)', 'o');

      const plan = parseAiJson(tx);
      log('📋 分析: ' + (plan.analysis || 'OK'));

      if (plan.sheetMapping) {
        plan.sheetMapping.forEach(sm => {
          const icon = sm.action === 'expand' ? '🔄' : '⏭';
          log('  ' + icon + ' ' + sm.inputSheet + ': ' + (sm.reason || sm.action));
        });
      }

      const ciCount = plan.columnInsertions?.length || 0;
      const riCount = plan.rowInsertions?.length || 0;
      const cuCount = plan.cellUpdates?.length || 0;
      log('columnInsertions: ' + ciCount + ', rowInsertions: ' + riCount + (cuCount ? ', cellUpdates: ' + cuCount : ''));

      if (ciCount === 0 && riCount === 0 && cuCount === 0) {
        log('⚠ AI 沒有產出任何變更。可能是歷史差異不足或模型需要升級。', 'e');
        log('建議：試試切換到 gemini-2.5-flash 或 gemini-2.5-pro 模型', 'i');
      }

      const sheetStats = {};
      (plan.columnInsertions || []).forEach(ci => {
        const s = ci.sheet || '未指定';
        sheetStats[s] = (sheetStats[s] || 0) + 1;
      });
      (plan.rowInsertions || []).forEach(r => {
        const s = r.sheet || '未指定';
        sheetStats[s] = (sheetStats[s] || 0) + 1;
      });
      (plan.cellUpdates || []).forEach(u => {
        const s = u.sheet || '未指定';
        sheetStats[s] = (sheetStats[s] || 0) + 1;
      });
      Object.entries(sheetStats).forEach(([s, count]) => {
        log('  Sheet [' + s + ']: ' + count + ' 項變更');
      });

      log('📝 套用展開計畫（僅新增欄/行，不修改已有資料）...');

      const { es, blob: outBlob, applied, skipped, sheetCount, processedSheets } = await xlBuild(inputBuf, plan);
      setExpanded(es); setBlob(outBlob);
      setMatchInfo(completePairs.length + ' 組歷史配對');

      log('✅ 套用 ' + applied + ' 項變更, ' + sheetCount + ' 個 Sheet 完整保留', 'o');
      if (processedSheets?.length) {
        log('已展開的頁籤: ' + processedSheets.join(', '), 'o');
      }
      if (skipped > 0) {
        log('跳過 ' + skipped + ' 個 (無對應頁籤)', 'i');
      }
      log('🎉 展表完成!', 'o');
      setStep(4);
    } catch (e) {
      log('❌ ERROR: ' + e.message, 'e');
      log('程序已停止。', 'e');
      setStep(2);
    } finally {
      setProcessing(false);
    }
  };

  const handleDownload = () => {
    if (!blob || !inputFile) return;
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = inputFile.name.replace(/\.(xls|xlsx)$/i, '') + '_品檢展開.xlsx';
    a.click();
  };

  const handleReset = () => {
    setInputFile(null); setInputParsed(null); setInputBuf(null); setExpanded(null); setBlob(null); setMatchInfo(''); setSimilarity(null); setLogs([]); setStep(0);
  };

  const handleSaveCfg = (k, m) => {
    setGemKey(k); localStorage.setItem('qse_key', k); gemInit(k);
    setGemModel(m); localStorage.setItem('qse_model', m);
    setShowCfg(false);
    log('設定已儲存 ✓ (模型: ' + m + ')', 'o');
  };

  const STEPS = ['選擇專案', '上傳檢驗單', '比對歷史', 'AI 展開', '下載'];
  const completePairs = (pairs || []).filter(p => p.before && p.after);

  return (
    <div className="app-container">
      {/* Settings Modal */}
      {showCfg && <SettingsModal gemKey={gemKey} gemModel={gemModel} onSave={handleSaveCfg} onClose={() => setShowCfg(false)} />}

      {/* Header */}
      <header className="app-header">
        <div className="app-header__badge">
          <span>🔬</span> QA Sheet Expander v7
        </div>
        <h1 className="app-header__title">品檢檢驗單展表工具</h1>
        <p className="app-header__subtitle">下載所有歷史配對 → 全面比對差異 → AI 精確展開</p>
        <button onClick={() => setShowCfg(true)} className={`app-header__key-btn ${gemKey ? 'app-header__key-btn--active' : 'app-header__key-btn--inactive'}`}>
          {gemKey ? '🔑 ' + gemModel + ' ✓' : '🔑 設定 Gemini Key'}
        </button>
      </header>

      <StepBar steps={STEPS} current={Math.min(step, 4)} />

      {/* 專案選擇 */}
      <div className="card">
        <div className="card__title">📁 選擇專案</div>
        {!projects ? (
          <div className={`loading-state ${lProj ? 'loading-state--active' : ''}`}>
            {lProj ? <><Spinner /> 讀取中...</> : '連接 Drive...'}
          </div>
        ) : (
          <div className="selector-grid">
            {projects.map(p => (
              <div key={p.id} onClick={() => handleSelProj(p.id, p.name)} className={`selector-item ${p.id === projId ? 'selector-item--active' : ''}`}>
                {p.id === projId ? '✓ ' : ''}{p.name}
              </div>
            ))}
          </div>
        )}
        {lCat && <div className="loading-inline"><Spinner /> 讀取類型...</div>}
      </div>

      {/* 活動類型 */}
      {cats && (
        <div className="card">
          <div className="card__title">📂 {projName} — 選擇類型</div>
          <div className="selector-grid selector-grid--sm">
            {cats.map(c => (
              <div key={c.id} onClick={() => handleSelCat(c.id, c.name)} className={`selector-item ${c.id === catId ? 'selector-item--active' : ''}`}>
                {c.id === catId ? '✓ ' : ''}{c.name}
              </div>
            ))}
          </div>
          {lFile && <div className="loading-inline"><Spinner /> 讀取檔案...</div>}
        </div>
      )}

      {/* 配對 + 上傳 */}
      {pairs && (
        <div className="card">
          <div className="card__title">📚 歷史展前/展後配對</div>
          <div className="card__subtitle">
            {projName}/{catName} — {completePairs.length} 組完整配對（共 {files?.length || 0} 個檔案）—{' '}
            <strong style={{ color: 'var(--accent-primary-light)' }}>全部都會下載比對</strong>
          </div>

          {completePairs.length > 0 ? (
            <div className="pair-list">
              {completePairs.map((p, i) => (
                <div key={i} className="pair-card">
                  <div>
                    <div className="pair-card__label">展表前：</div>
                    <div className="pair-card__name">{p.before.name}</div>
                    <div className="pair-card__arrow">↓</div>
                    <div className="pair-card__label">展表後：</div>
                    <div className="pair-card__name pair-card__name--after">{p.after.name}</div>
                  </div>
                  <div className="pair-card__badge">✓ 配對 {i + 1}</div>
                </div>
              ))}
            </div>
          ) : (
            <div className="empty-state">找不到完整的展前/展後配對。請確認檔名含「展表前」「展表後」或「品檢」</div>
          )}

          {/* 上傳區 */}
          <hr className="divider" />
          <div className="card__title">📋 上傳待展表檔案</div>

          {!inputFile ? (
            <>
              <input id="fileInput" type="file" accept=".xls,.xlsx" style={{ display: 'none' }} onChange={e => { if (e.target.files?.[0]) handleInput(e.target.files[0]); e.target.value = ''; }} />
              <div
                className="upload-zone"
                onClick={() => document.getElementById('fileInput')?.click()}
                onDragOver={e => { e.preventDefault(); e.currentTarget.classList.add('upload-zone--dragover'); }}
                onDragLeave={e => { e.currentTarget.classList.remove('upload-zone--dragover'); }}
                onDrop={e => { e.preventDefault(); e.currentTarget.classList.remove('upload-zone--dragover'); if (e.dataTransfer.files[0]) handleInput(e.dataTransfer.files[0]); }}
              >
                <span className="upload-zone__icon">📋</span>
                <div className="upload-zone__text">拖放或點擊上傳待展表</div>
                <div className="upload-zone__hint">.xls / .xlsx</div>
              </div>
            </>
          ) : (
            <>
              <div className="upload-info">
                <div className="upload-info__details">
                  <div className="upload-info__name">✓ {inputFile.name}</div>
                  {inputParsed && <div className="upload-info__meta">{inputParsed.sheetNames.length} sheets: {inputParsed.sheetNames.join(', ')}</div>}
                </div>
                <button onClick={() => { setInputFile(null); setInputParsed(null); setInputBuf(null); setSimilarity(null); setStep(0); }} className="upload-info__remove">×</button>
              </div>
              {inputParsed && (
                <div style={{ marginTop: 14 }}>
                  <SheetPreview sheets={inputParsed.sheets} title="待展表預覽" activeTab={pvTab.input} onTabChange={t => setPvTab(p => ({ ...p, input: t }))} />
                </div>
              )}
            </>
          )}
        </div>
      )}

      {/* 展開按鈕 */}
      {step >= 2 && step < 4 && (
        <button onClick={handleExpand} disabled={processing} className={`btn ${processing ? 'btn--disabled' : 'btn--primary'}`}>
          {processing ? <><Spinner /> AI 分析展開中...</> : '🚀 開始 AI 展表（下載全部 ' + completePairs.length + ' 組歷史配對）'}
        </button>
      )}

      {/* Processing */}
      {step === 3 && (
        <div className="card" style={{ marginTop: 16 }}>
          <div className="processing-hero">
            <div className="processing-hero__icon">⚙️</div>
            <div className="processing-hero__text">下載所有歷史檔案 → 逐組比對差異 → AI 分析共同模式 → 套用展開...</div>
            <div className="processing-hero__sub">共 {completePairs.length} 組歷史配對需要下載和比對</div>
          </div>
          <Logs logs={logs} />
        </div>
      )}

      {/* Results */}
      {step === 4 && (
        <div className="card">
          <div className="result-hero">
            <div className="result-hero__icon">✅</div>
            <div className="result-hero__title">展表完成！</div>
            {matchInfo && <p className="result-hero__info">已參考 {matchInfo}</p>}
          </div>
          {similarity && <SimilarityBadge similarity={similarity} />}
          {expanded && (
            <div style={{ marginTop: 16 }}>
              <SheetPreview sheets={expanded} title="展開後預覽（所有頁籤）" activeTab={pvTab.result} onTabChange={t => setPvTab(p => ({ ...p, result: t }))} />
            </div>
          )}
          <div style={{ marginTop: 16 }}>
            <button onClick={handleDownload} className="btn btn--success">📥 下載展開後的 Excel</button>
            <button onClick={handleReset} className="btn btn--ghost">🔄 展開另一份</button>
          </div>
        </div>
      )}

      {/* Logs (non-processing) */}
      {logs.length > 0 && step !== 3 && (
        <div className="card">
          <div className="card__title" style={{ fontSize: 13 }}>處理日誌</div>
          <Logs logs={logs} />
        </div>
      )}

      <div className="app-footer">Powered by Google Gemini — v7</div>
    </div>
  );
}
