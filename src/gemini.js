import { GoogleGenerativeAI } from '@google/generative-ai';

let genAI = null;

export function gemInit(apiKey) {
  genAI = new GoogleGenerativeAI(apiKey);
}

const SYS = `你是品檢檢驗單展表專家。

## 核心原則

你會收到歷史「展前→展後」的結構性差異，以及新的待展表檢驗單。
你的任務是分析歷史展開模式，並在新檢驗單上做同樣的結構性展開。

**最重要規則：絕對不能刪除或修改上傳資料中任何已有的內容。只能透過「新增欄」和「新增行」來展開。**

## 展開操作類型

### 1. columnInsertions — 在指定欄後面插入新的欄位
當歷史展開顯示比展前多了幾欄（如 android、ios、win 驗證欄），你需要在新檢驗單的相應位置也插入同樣的欄位。

格式：
{
  "sheet": "頁籤名",
  "afterColumn": "C",
  "headers": [
    {"row": 2, "value": "android"},
    {"row": 5, "value": "獎勵數量"}
  ]
}

- afterColumn：在哪個欄之後插入新欄
- headers：在哪些行填入標題文字（其他行自動留空）

### 2. rowInsertions — 在指定行後面插入新的行
當歷史展開顯示增加了子測試項目行（如獎勵明細行），你需要在新檢驗單的對應項目後面也插入同樣的行。

格式：
{
  "sheet": "頁籤名",
  "afterRow": 15,
  "cells": {"B": "子測試項目名稱", "C": "", "D": ""}
}

- afterRow：在哪一行之後插入
- cells：新行中每個欄位的值（空字串表示空白格讓使用者自己填）

## 定位規則（極重要）

**不要假設新檢驗單的格子位置和歷史一樣！** 你必須根據新檢驗單中每個測試項目的**實際文字內容和行號**來決定插入位置。

步驟：
1. 從歷史差異中找出「在哪個測試項目旁邊做了什麼展開」
2. 在新檢驗單中找到相同或類似的測試項目（用文字寬鬆匹配）
3. 根據該項目在新檢驗單中的實際行號/欄號來設定 afterRow/afterColumn

例如：歷史在「充值功能」(第10行) 後面加了3行子項目
→ 如果新檢驗單的「充值功能」在第15行，那 afterRow 應該設為 15

## 嚴格禁止事項

1. **禁止使用 cellUpdates 修改已有值** — 如果一個格子已經有文字，絕對不能改它
2. **禁止刪除任何行或列**
3. **禁止修改 A/B 欄的測試項目名稱**
4. **禁止新增歷史中沒有的展開項目** — 只複製歷史展開的結構模式
5. **禁止新增新檢驗單中不存在的測試分類** — 如果新檢驗單沒有「排行榜」，就不要加

## 展開的詳細度

你的展開要和歷史展表一樣詳細：
- 如果歷史加了 3 個驗證欄，新檢驗單也要加 3 個
- 如果歷史在某類測試項目下加了 5 行子項目，新檢驗單的同類項目也要加 5 行
- 標題行的文字要和歷史一致（如 「android」「ios」「win」「獎勵」「數量」等）
- 每個有歷史對應的頁籤都必須展開

## 回傳 JSON 格式

{
  "analysis": "展開策略說明",
  "sheetMapping": [
    {"inputSheet": "頁籤名", "action": "expand|skip", "reason": "原因"}
  ],
  "columnInsertions": [
    {"sheet": "頁籤名", "afterColumn": "C", "headers": [{"row": 2, "value": "android"}]}
  ],
  "rowInsertions": [
    {"sheet": "頁籤名", "afterRow": 15, "cells": {"B": "子項目", "C": "", "D": ""}}
  ]
}

只回傳JSON。`;

export async function gemCall(msg, logFn) {
  if (!genAI) throw new Error('請先設定 Gemini Key');
  const modelId = localStorage.getItem('qse_model') || 'gemini-2.0-flash';
  const m = genAI.getGenerativeModel({
    model: modelId,
    systemInstruction: SYS,
    generationConfig: { maxOutputTokens: 65536, temperature: 0.1, responseMimeType: 'application/json' },
  });
  try {
    const result = await m.generateContent(msg);
    return result.response.text();
  } catch (e) {
    const errMsg = e.message || '';
    if (errMsg.includes('429')) {
      const waitMatch = errMsg.match(/retry\s*in\s*(\d+)/i);
      const waitSec = waitMatch ? Math.min(parseInt(waitMatch[1]) + 5, 90) : 60;
      if (logFn) logFn('⏳ 配額限制，等待 ' + waitSec + ' 秒後重試一次...', 'e');
      for (let i = waitSec; i > 0; i--) {
        if (i % 15 === 0 || i <= 5) { if (logFn) logFn('   剩餘 ' + i + ' 秒...', 'i'); }
        await new Promise(r => setTimeout(r, 1000));
      }
      try {
        if (logFn) logFn('重試中...', 'i');
        return (await m.generateContent(msg)).response.text();
      } catch (e2) {
        if (e2.message?.includes('PerDay')) {
          const alt = modelId === 'gemini-2.0-flash' ? 'gemini-2.5-flash' : 'gemini-2.0-flash';
          throw new Error('❌ 每日配額已用完。請切換模型為 ' + alt + '，或明天再試。');
        }
        throw new Error('❌ API 重試失敗，程序停止。' + e2.message);
      }
    }
    throw new Error('❌ API 失敗，程序停止。' + errMsg);
  }
}
