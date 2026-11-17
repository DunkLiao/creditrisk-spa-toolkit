# 佐證抽樣實驗室 — 技術規格書（SPEC）

| 項目 | 內容 |
|---|---|
| 產品名稱 | 佐證抽樣實驗室（Evidence Sampling Lab） |
| 版本 | V3 |
| 交付檔 | `佐證抽樣實驗室.html`（單一檔案，約 120 KB／1,391 行；檔名不含版本號，版本資訊寫在檔案內） |
| 技術棧 | HTML + CSS + Vanilla JavaScript（ES2020），無框架、無外部函式庫 |
| 儲存 | IndexedDB（資料庫 `bot-sample-lab`，版本 `1`） |
| 網路 | 零網路：不使用 fetch／XHR／WebSocket，不需要也不含任何 API key |
| 支援瀏覽器 | Microsoft Edge／Google Chrome 最新版（需支援 IndexedDB、TextDecoder、`String.prototype.normalize`） |
| 介面語言 | 繁體中文 |
| 配色 | 臺銀酒紅 `#8A1538`＋金 `#C9A24B`（`:root` CSS 變數） |

---

## 1. 目的與範圍

### 1.1 目的
讓查核與覆核人員把要抽樣的資料（母體）貼進來後，用**可重現、可說明、可追蹤**的方式抽出樣本，追蹤佐證調閱、查核結果與替補，最後產出可附在工作底稿的抽樣紀錄。

### 1.2 解決的痛點
| 痛點 | 對應設計 |
|---|---|
| Excel 手動抽樣，事後無法重現同一組樣本 | 種子亂數＋母體依鍵值排序，結果與匯入順序無關 |
| 說不清楚每筆為何被抽中 | 每筆樣本記錄「入選原因」 |
| 母體事後被改動無法察覺 | 母體指紋／執行指紋（SHA-256）＋「驗證重現」 |
| 樣本集中在少數單位、漏掉高風險區塊 | 覆蓋地圖＋未涵蓋高風險區塊清單 |
| 佐證追蹤散落在信件與檔案 | 樣本狀態、期限、附件、查核結果集中管理 |

### 1.3 不在範圍內
- 不連任何後端、資料庫或 API；不支援多人即時共用（以 JSON 備份交換）。
- 不直接讀 `.xlsx`；只支援貼上表格（TSV）與 CSV／TSV 檔。
- 不提供統計推論（誤述推估、上限誤述計算），信賴水準等參數僅記錄用。

---

## 2. 系統架構

```
┌───────────────────────── 佐證抽樣實驗室.html ─────────────────────────┐
│ CSS（:root 主題變數、RWD ≤820px、@media print A4）                     │
│ JS                                                                      │
│  0 設定區   APP / SCHEMA / DATA_STORES / METHODS / ALLOC / METHOD_DESC   │
│  1 DB 封裝  open/get/all/byIndex/count/put/del/clear/countBy/tx/putMany │
│  2 共用工具 esc/fmt*/modal/toast/download/toCSV/toExcelXML/barChart     │
│  3 密碼學   sha256（純 JS）/ xmur3 / mulberry32 / makeRng / shuffle     │
│  4 抽樣引擎 execPlan → selectBy → sSRS/sSys/sStrat/sMUS/sRisk/sKey      │
│  5 路由     hash router（go/refresh）/ State / syncSel / mountGrid      │
│  6–14 頁面  總覽/母體/計畫/執行/覆蓋地圖/追蹤/資料管理/說明             │
│  15 啟動    boot()                                                      │
└─────────────────────────────────────────────────────────────────────────┘
                     │ IndexedDB: bot-sample-lab (v1)
                     ▼
 meta · populations · popRows · plans · runs · samples · attachments
```

### 2.1 設計原則
- 單一 HTML 檔，CSS／JS 全部內嵌，圖示使用 emoji，圖表為手刻 SVG。
- 所有插入 `innerHTML` 的使用者資料、編號（含 HTML 屬性值）一律經 `esc()`。
- 刪除母資料時，子資料與附件在**同一個 IndexedDB 交易**中刪除。
- 新增 store／index 時 `APP.dbVersion + 1`，`onupgradeneeded` 自動補建（V1～V3 皆為 v1，資料互通）。

---

## 3. 資料模型（IndexedDB）

資料庫：`bot-sample-lab`，版本 `1`。所有主鍵為字串。

### 3.1 `meta`（keyPath: `key`）
| key | value |
|---|---|
| `lastPage` | 上次停留頁面 id |
| `sel` | `{ popId, planId, runId }` 上次選取 |
| `lastBackup` | 上次 JSON 備份時間（ISO） |

### 3.2 `populations` 母體（keyPath: `id`）
| 欄位 | 型別 | 說明 |
|---|---|---|
| id | string | `uid()` |
| name | string | 母體名稱（必填） |
| period | string | 期別，例 `115Q1` |
| sourceNote | string | 資料來源說明 |
| columns | string[] | 原始標題列 |
| fieldMap | `{key,amount,stratum,unit,risk}` | 對應的原始欄名，未使用為 `''` |
| rowCount | number | 筆數 |
| totalAmount | number | 金額合計（含正負號） |
| fingerprint | string | 母體指紋（64 碼 hex） |
| createdAt / updatedAt | ISO string | |

Index：`byPeriod(period)`、`byCreated(createdAt)`

### 3.3 `popRows` 母體明細（keyPath: `id`）
| 欄位 | 型別 | 說明 |
|---|---|---|
| id | string | `${popId}_${rowNo}` |
| popId | string | |
| rowNo | number | 匯入順序（1 起） |
| key | string | 鍵值；重複時加 `#2`、`#3`；空白時為 `列{n}` |
| amount | number | 解析後金額；無法解析為 0 |
| stratum / unit / risk | string | 分層／單位／風險（可空白） |
| data | object | 原始整列 `{欄名: 值}` |

Index：`byPop(popId)`、`byPopKey([popId, key])`

### 3.4 `plans` 抽樣計畫（keyPath: `id`）
| 欄位 | 型別 | 說明 |
|---|---|---|
| id, popId, name | string | |
| method | enum | `srs`／`sys`／`strat`／`mus`／`risk`／`keyplus` |
| n | int ≥ 1 | 樣本數（`keyplus` 為全查以外的抽樣數） |
| m | int ≥ 0 | 預備數 |
| seed | string | 種子（必填） |
| params | object | 見 3.4.1 |
| note | string | 備註 |
| createdAt / updatedAt | ISO | |

Index：`byPop(popId)`

#### 3.4.1 `params`（`defaultParams()`）
| 參數 | 預設 | 用途 |
|---|---|---|
| allocation | `prop` | `prop` 比例／`equal` 等額／`neyman` |
| minPerStratum | 1 | 每層最低抽樣數 |
| keyItemThreshold | 0 | 高額全查門檻（以絕對值比較） |
| baseMethod | `srs` | `keyplus` 其餘項目的方法（不可為 `keyplus`） |
| riskWeights | `{高:5, 中:2, 低:1}` | 風險標籤權重 |
| highRiskWeight | 5 | 覆蓋地圖高風險門檻權重 |
| confidence | 0.95 | 記錄用 |
| tolerableRate | 0.05 | 記錄用 |
| tolerableMisstatement | 0 | 記錄用 |

### 3.5 `runs` 抽樣執行（keyPath: `id`）
| 欄位 | 說明 |
|---|---|
| id, planId, popId, planName | |
| seed, method | 執行當下的值 |
| popFingerprint | 執行時的母體指紋 |
| paramsSnapshot | `{method, n, m, seed, params}`，之後編輯計畫不影響 |
| sampleCount / reserveCount | 正選＋全查數／預備數 |
| coverage | `{popCount, popAmount, sampleCount, sampleAmount, countRate, amountRate}`（執行當下快照） |
| info / warnings | string[]，方法說明與警告 |
| locked / lockedAt | 鎖定旗標與時間 |
| runHash | 執行指紋 |
| createdAt | ISO |

Index：`byPlan(planId)`、`byPop(popId)`

### 3.6 `samples` 樣本（keyPath: `id`）
| 欄位 | 說明 |
|---|---|
| id, runId, rowId, key | |
| seq | 1 起序號（正選／全查在前，預備在後） |
| role | `正選`／`全查`／`預備` |
| reason | 入選原因 |
| status | `待調閱`／`已取得`／`已查核`／`無法取得`／`已替補`／`備用` |
| result | `正常`／`異常`／`''`（僅 `已查核` 時有值） |
| finding | 查核說明 |
| dueDate | 調閱期限（`YYYY-MM-DD`，預備樣本初始為空） |
| checkedAt | 查核時間 |
| replacedBy / replaces | 替補關係（樣本 id） |
| replaceReason / replacedAt | 替補原因／時間 |
| amount, stratum, unit, risk | 冗餘自母體列，供追蹤與覆蓋計算 |

Index：`byRun(runId)`、`byStatus(status)`、`byRunRole([runId, role])`

### 3.7 `attachments` 附件（keyPath: `id`）
| 欄位 | 說明 |
|---|---|
| id, sampleId | |
| name, mime, size | |
| blob | `Blob`（JSON 備份時轉 Base64 data URL） |
| createdAt | |

Index：`bySample(sampleId)`

### 3.8 刪除連鎖
| 動作 | 同一交易內刪除 | 限制 |
|---|---|---|
| 刪除母體 | popRows、plans、runs、samples、attachments、population | 含鎖定抽樣時須輸入「刪除」 |
| 刪除計畫 | runs、samples、attachments、plan | 含鎖定抽樣時禁止 |
| 刪除抽樣 | samples、attachments、run | 已鎖定時禁止 |
| 清除全部 | 六個資料 store（不含 meta） | 確認對話框 |

---

## 4. 頁面與功能規格

路由採 hash（`#dash` 等）。`go()` 先渲染到離線容器，以 token 確認仍是最新導覽才替換畫面；同頁重繪保留捲動位置。

| id | 頁面 | 主要功能 |
|---|---|---|
| `dash` | 📊 總覽 | 6 張統計卡、樣本狀態長條圖、近期 8 次抽樣、逾期未取得提醒 |
| `pop` | 🗂️ 母體管理 | 匯入／重新匯入／刪除；摘要統計、分層統計、母體明細 |
| `plan` | 🧮 抽樣計畫 | 計畫 CRUD、配置預覽、屬性抽樣與 MUS 樣本數試算 |
| `run` | 🎲 執行與結果 | 執行、驗證重現、鎖定、刪除、CSV／Excel 底稿／列印 |
| `map` | 🗺️ 覆蓋地圖 | 熱力圖（分層×單位／風險，筆數／金額）、長條圖、未涵蓋高風險區塊 |
| `track` | 📎 佐證追蹤 | 狀態更新、附件、替補、批次設定、篩選 |
| `data` | 💾 資料管理 | JSON 備份、合併匯入、還原、範例資料、清除、儲存空間 |
| `help` | 📖 使用說明 | 操作說明、方法說明、FAQ、修正紀錄 |

### 4.1 選取狀態（`State` 與 `syncSel`）
- `State = { page, popId, planId, runId, mapAxis, mapMetric, trackFilter, tok, rendered }`
- `runId === ''` 表示「刻意未選抽樣」（例如新建計畫後），`null` 表示自動挑選。
- `syncSel(needRun)`：校正不存在的選取；`needRun=true`（覆蓋地圖、佐證追蹤、列印）時，若所選計畫沒有執行紀錄，就退回該母體最近一次抽樣。
- 選取狀態寫入 `meta.sel`，下次開啟還原。

### 4.2 通用清單 `mountGrid(host, cfg)`
- 功能：全文搜尋（防抖 200ms）、點欄位排序（數字比大小；文字為 `localeCompare(zh-Hant, {numeric:true})` 自然排序）、CSV／Excel 匯出（匯出目前篩選與排序後的全部列）。
- 畫面顯示上限 `cfg.limit`（預設 200），超過時顯示提示。
- `selectable`：每列勾選框＋表頭全選（反映目前畫面是否全選，可再按一次取消）。
- `selected()` 只回傳**目前篩選後仍看得到**的勾選列。
- 搜尋、排序、勾選狀態以 `GS[cfg.id]` 跨重繪保存。

### 4.3 母體匯入
1. 輸入名稱（必填）、期別、來源說明。
2. 貼上表格或選 CSV／TSV 檔（先試 UTF-8 fatal，失敗改 Big5）。
3. `parseTable`：第一列含 `\t` 即以 Tab 分隔，否則以逗號；支援雙引號與 `""` 跳脫；略過空白列。
4. 欄位自動猜測（正則）：鍵值 `案號|編號|號碼|代號|帳號|單號|key|id`、金額 `金額|餘額|本金|amount`、分層 `分層|層|類別|種類|產品|stratum|type`、單位 `單位|分行|部門|營業|unit|branch`、風險 `風險|等級|risk|rating`；重新匯入時優先沿用原 fieldMap。
5. 鍵值與金額為必對應欄位。
6. 重新匯入時若指紋改變，需確認後才覆蓋。

#### 4.3.1 金額解析 `parseNum`
- 先 `NFKC` 正規化（全形轉半形），`− – —` 視為 `-`，移除千分位與空白。
- 整體括號 `( )` → 負數；尾端 `-` → 負數（與前置負號互相抵銷）。
- 以正則 `([+-]?)[^\d.+-]*(\d+\.?\d*(?:[eE][+-]?\d+)?|\.\d+)` 擷取：允許負號在貨幣符號之前（`-$1,000`、`-NT$1,000`）。
- 無數字 → `NaN`（匯入時以 0 計並計數提示）。

| 輸入 | 結果 |
|---|---|
| `NT$ 1,000` | 1000 |
| `-$1,000`、`-NT$1,000`、`(1,000)`、`1,000-`、`(NT$1,000)` | -1000 |
| `1.2E+6` | 1200000 |
| `１２３４`、`－１，０００` | 1234、-1000 |
| `.5`、`12.5%` | 0.5、12.5 |
| `abc`、空白 | NaN |

### 4.4 抽樣計畫
- 編輯對話框依方法只顯示相關欄位；`keyplus` 時 n 標籤改為「其餘項目抽樣數 n」。
- 新增計畫後設定 `State.runId = ''`，執行頁顯示「此計畫尚未執行抽樣」。
- 配置預覽：直接以 `execPlan` 計算，列出正選、全查、預備筆數、說明、警告，分層方法另列各層配置表。
- 試算器可「以此建立計畫」，帶入方法、n 與信賴參數。

### 4.5 執行與結果
- 執行前讀取資料庫最新的計畫內容；同計畫已有相同種子的執行時提醒「會得到相同樣本」。
- 預設調閱期限：今天 + 14 天（可留空）；預備樣本不設期限、狀態為 `備用`。
- 結果區顯示即時涵蓋率（排除 `無法取得`、`已替補`、未啟用預備）。
- 樣本清單欄位：`# / 角色 / 鍵值 / 金額 / 分層 / 單位 / 風險 / 入選原因 / 狀態 / 替補關係`。
- 刪除抽樣後停留在同一計畫：有其他執行則選最新一次，否則顯示「尚未執行」。

### 4.6 覆蓋地圖
- 軸向：縱軸分層（最多 15 個，其餘併為「其他」），橫軸單位（依筆數取前 12）或風險等級（依權重排序，同樣上限 12）。
- 指標：筆數涵蓋率 `n/N` 或金額涵蓋率 `a/A`（金額以絕對值）。
- 顏色透明度 `0.08 + 0.92 × (v / max)`，以本圖最高格為基準。
- 紅框條件：該格含高風險項目（`weightOf ≥ highRiskWeight`）且有效涵蓋樣本數為 0。
- 計入涵蓋的樣本：`isCovered = isActive && status ≠ 無法取得`。

### 4.7 佐證追蹤
- 狀態機：

```
                ┌──────────► 已取得 ──────► 已查核（正常／異常）
   待調閱 ──────┤
                └──────────► 無法取得 ──[替補]──► 已替補
   備用（預備）──[被選為替補]──► 待調閱
```

- 單筆更新可在 `待調閱／已取得／已查核／無法取得` 之間任意切換；選 `已查核` 必須指定結果。
- 替補：候選為同次抽樣中 `role=預備 且 status=備用`，同層優先、其次依序號。替補樣本期限沿用原期限；原期限已過則為今天 + 7 天。
- 批次設定：目標狀態 `已取得／待調閱／無法取得`；排除 `備用`、`已替補`；已查核者會清除結果並在確認框提示筆數。
- 篩選：全部、各狀態、`異常`、`逾期`（`待調閱` 且期限 < 今天）。
- 逾期判斷以本地日期字串比較（`YYYY-MM-DD`）。

### 4.8 總覽統計定義
| 指標 | 定義 |
|---|---|
| 有效樣本 | `isActive`：非 `已替補`，且非「預備＋備用」 |
| 待調閱 | 有效樣本中 `status = 待調閱` |
| 逾期未取得 | 待調閱且 `dueDate < 今天` |
| 查核異常 | 有效樣本中 `已查核 且 result = 異常` |
| 完成率 | 已查核 ÷ 有效樣本 |

---

## 5. 抽樣演算法規格

### 5.1 共通
- **前處理**：`canon(rows)` 依 `key` 字典序排序（同鍵依 `rowNo`），使結果與匯入順序無關。
- **亂數**：`makeRng(seed, tag) = mulberry32(xmur3(seed + '|' + tag)())`；不同用途使用不同 tag（如 `srs`、`sysres`、`strat|{層名}`），避免一層的變動影響其他層。
- **洗牌**：Fisher–Yates（`shuffle(arr, rng)`）。
- **n 修正**：n 下限 1；非 `keyplus` 時 n > N 會降為 N 並警告。
- **完整性檢查**：同一筆不可同時出現在正選與預備（重複即拋錯）。
- **不足警告**：非 `strat`／`keyplus` 方法若正選數 < n，加入警告。

### 5.2 各方法
| 方法 | 正選 | 預備 | 入選原因範例 |
|---|---|---|---|
| `srs` 簡單隨機 | 洗牌後前 n 筆 | 接續 m 筆 | 簡單隨機：洗牌後第 3 筆 |
| `sys` 系統抽樣 | `k = N/n`，`start = rng()·k`，第 i 筆取 `floor(start + i·k)` | 未入選者以獨立亂數洗牌取 m 筆 | 系統抽樣：間距 k=10.00、起點 2.05，第 4 次命中（母體順位 33） |
| `strat` 分層抽樣 | 依 `allocate()` 配置，層內洗牌取前 a_h 筆 | 依剩餘容量以比例配置 m，層內接續取 | 分層抽樣（比例配置）：「房貸」層母體 128 筆，配置 10 筆，層內隨機第 2 |
| `mus` 金額單位抽樣 | 見 5.3 | 未命中項目以 `ln(u)/|A|` 金額加權排序取 m 筆 | 金額單位抽樣：累計金額命中點 1,234.00（間距 J=…、起點 …） |
| `risk` 風險加權 | Efraimidis–Spirakis：key = `ln(u)/w`，取最大 n 筆；`w ≤ 0` 不參與 | 排名 n+1～n+m | 風險加權隨機：風險「高」權重 5，加權排名第 1 |
| `keyplus` 高額全查 | `|amount| ≥ threshold` 全查，其餘以 baseMethod 抽 n 筆（tag `k`） | 由 baseMethod 產生 | 金額 9,800,000.00 ≥ 全查門檻 5,000,000.00 |

### 5.3 MUS 細節
1. `pool` = 金額絕對值 > 0 的項目（0 元筆數加入警告）。
2. 反覆：`J = Σ|A| / k`，凡 `|A| ≥ J` 列為 `全查` 並自 pool 移除、`k -= 全查數`，直到沒有新增全查項目。
3. 若 `k > 0`：`J = Σ|A| / k`，`start = rng()·J`，沿排序後的 pool 累計金額，每個命中點 `start + i·J < 累計` 時選中當筆（同筆最多計一次）。
4. 命中筆數 < k 時加入警告。

### 5.4 分層配置 `allocate(groups, n, mode, minPer)`
1. 每層先配 `min(容量, minPer)`。
2. 權重：`equal` = 1；`prop` = N_h；`neyman` = N_h × S_h（S_h 為金額樣本標準差）。
3. 剩餘量依權重比例分配，小數以**最大餘數法**補齊，並受各層容量上限；有剩餘且仍有可配層時重複。
4. 若 `Σ minPer > n`，總數可大於 n（已於說明頁揭露）。

### 5.5 樣本數試算
- 屬性抽樣（零預期偏差）：$$n=\left\lceil \frac{\ln(1-c)}{\ln(1-r)} \right\rceil$$（95%、5% → 59）
- MUS：$$n=\left\lceil \frac{\text{總額}\times RF}{TM} \right\rceil$$，RF：90%=2.31、95%=3.0、99%=4.61；其他信賴水準以 `−ln(1−c)` 近似。

---

## 6. 指紋與可重現性

| 指紋 | 計算 |
|---|---|
| 母體指紋 | `SHA256( sort( rows.map(r => r.key + '\t' + r.amount) ).join('\n') )` |
| 執行指紋 | `SHA256( popFingerprint + '|' + JSON.stringify(paramsSnapshot) + '|' + 樣本清單(role:key 以 \n 串接) )` |

- SHA-256 為純 JS 實作（不依賴 `crypto.subtle`，`file://` 下亦可用），已與 Python `hashlib` 比對一致。
- **驗證重現**比對三項：(1) 目前母體指紋 = 執行時指紋；(2) 以 `paramsSnapshot` 重抽，`role:key` 序列完全相同；(3) 重算執行指紋一致。
- 已知限制：母體指紋只涵蓋鍵值與金額；若分層、單位、風險欄位被修改，(1) 會通過但 (2) 會失敗（說明頁已揭露）。
- 鎖定：`locked=true` 後該執行不可刪除，所屬計畫也不可刪除；所屬母體刪除須輸入「刪除」。

---

## 7. 匯入／匯出規格

### 7.1 CSV（`toCSV`）
- UTF-8 BOM（`\uFEFF`），`\r\n` 換行，含 `,` `"` 換行的欄位以雙引號包覆。
- **公式注入防護**：以 `= + - @ Tab CR` 開頭且不是純數字（含科學記號）的值，前置 `'`。

### 7.2 Excel（`toExcelXML`，SpreadsheetML 2003，`.xls`）
- 多工作表；表頭酒紅底白字粗體。
- 移除 XML 1.0 不合法控制字元（`\x00–\x08 \x0B \x0C \x0E–\x1F`）；工作表名稱移除 `\ / ? * [ ] :` 並截至 31 字元。
- 數字型別自動為 Number，其餘為 String。

#### 抽樣底稿（5 張工作表）
| 工作表 | 內容 |
|---|---|
| 計畫參數 | 母體、期別、筆數、金額、母體指紋、方法、計畫名稱、n、m、種子、全部 params、執行時間、執行指紋、鎖定、說明、警告 |
| 正選樣本 | `role ≠ 預備` 或已替補上場的預備樣本（`isMainList`） |
| 預備樣本 | 尚未啟用的預備樣本 |
| 覆蓋統計 | 整體、各分層、各單位、各風險等級：母體／樣本筆數與金額、涵蓋率(%) |
| 替補紀錄 | 被替補樣本、替補樣本、原因、時間 |

樣本欄位：`# / 角色 / 鍵值 / 金額 / 分層 / 單位 / 風險 / 入選原因 / 狀態 / 替補關係 / 調閱期限 / 查核說明`。

### 7.3 列印報告
- 寫入 `#print-area` 後呼叫 `window.print()`；`@media print` 只顯示報告，A4、邊界 14mm。
- 內容：母體與參數表、母體／執行指紋、方法說明、覆蓋地圖（分層×單位，筆數）、正選（含替補上場者）與尚未啟用預備兩張表、製表／覆核／日期簽名欄。
- 未指定抽樣時（工具列、資料管理），使用目前母體的最近一次抽樣。

### 7.4 JSON 備份
```json
{
  "app": "佐證抽樣實驗室",
  "db": "bot-sample-lab",
  "version": 1,
  "exportedAt": "ISO",
  "data": {
    "populations": [], "popRows": [], "plans": [],
    "runs": [], "samples": [],
    "attachments": [{ "...": "...", "blob": "data:<mime>;base64,..." }]
  }
}
```
- 匯入檢查：`data` 必須存在、`db === 'bot-sample-lab'`、各 store 必須為陣列；任一不符即中止且不改動資料。
- **合併匯入**：同 id 覆蓋，其餘保留；**還原**：先清空六個資料 store 再匯入（同一交易）。
- `meta` 不納入備份。

---

## 8. 非功能需求

| 類別 | 要求 |
|---|---|
| 離線 | 雙擊 HTML 即可使用；執行期不得發出任何 http(s) 請求 |
| 安全 | 全面 `esc()`；CSV 公式注入防護；匯入 JSON 結構檢查；無 `eval`、無外部資源 |
| 效能 | 14 萬筆母體頁可開啟（測試環境約 14 秒）；`Math.max(...arr)` 一律改用 `maxOf/minOf` 迴圈避免爆堆疊；清單畫面最多渲染 100～300 列 |
| 防誤操作 | 對話框確定鍵處理中停用（防連點）；刪除、清除、還原、批次變更皆需確認 |
| 一致性 | 快速切頁時只保留最後一次導覽的畫面（render token） |
| 持久化 | 啟動時呼叫 `navigator.storage.persist()`；無痕模式或 IndexedDB 失敗顯示錯誤畫面 |
| RWD | 寬度 ≤ 820px 時左側選單改為上方橫排 |
| 時間顯示 | `toLocaleString('zh-TW', {hourCycle:'h23'})`，避免顯示 24:xx |
| 錯誤處理 | 全域 `error`／`unhandledrejection` 以 toast 顯示；頁面渲染錯誤顯示「發生錯誤」卡片與堆疊 |

---

## 9. 範例資料（`loadSample`）
- 以固定種子 `sample-demo-2026` 產生 480 筆授信案件（鍵值 `LN-115-0001`～`0480`）。
- 分層 4 類（企業授信、房貸、小額信貸、外匯融資，對數常態金額）；單位 8 間分行；風險高／中／低。
- 計畫 4 個：分層（n=40、m=8、每層最低 3）、MUS（n=30）、風險加權（n=25）、高額全查（門檻＝金額前 3% 分位）。
- 執行 2 次（分層、MUS），分層那次預先設定 18 筆已查核（含異常）、8 筆已取得、1 筆無法取得、1 筆已替補、5 筆逾期。

---

## 10. 驗收與測試

### 10.1 靜態檢查（`check_spa.py`）
- 無外部資源（src／href／@import／url()）、無網路 API、無 API key／Authorization 字樣、無佔位字。
- 內嵌 JS 以 `node --check` 通過語法檢查。

### 10.2 無頭瀏覽器實測
- 建立 IndexedDB → 載入範例 → 重新整理資料仍在（571 筆）→ 巡覽 8 頁無錯誤、無外部請求。

### 10.3 功能測試案例（V3 全數通過）
| 類別 | 案例 |
|---|---|
| 演算法 | 6 種方法：反轉輸入順序結果相同、換種子結果不同、筆數正確；MUS 大額自動全查；Neyman 配置總和 = n；屬性抽樣 n=59；0 元／權重 0 警告 |
| 密碼學 | SHA-256 與 hashlib 一致（空字串、ASCII、長中文） |
| 解析 | `parseNum` 17 種格式；實際匯入 `-$1,000`、`(NT$2,500)`、`3,000-`、`1.5E+3` |
| 匯出 | CSV 注入防護；XML 控制字元；底稿 5 張工作表；替補上場者列於正選並標示；母體清單匯出完整 64 碼指紋 |
| 流程 | 連點匯入只建 1 筆；新計畫執行頁顯示「尚未執行」；驗證重現 3 項通過；同種子提醒；鎖定後不可刪；輸入「刪除」才可刪鎖定母體；連鎖刪除歸零 |
| 追蹤 | 全選／取消全選；隱藏勾選列不被批次處理；替補期限 ≥ 今天；無法取得排除於涵蓋率 |
| 狀態 | 刪除唯一抽樣後停留原計畫；刪除兩次之一後選到剩餘那次；快速切頁停在最後一頁 |
| 安全 | 竄改備份（id 含 `"><img onerror>`）還原後開啟替補、編輯對話框不執行、不產生元素；非陣列 store 拒絕匯入且資料不變 |
| 其他 | 備份→清除→還原筆數一致；工具列列印未選抽樣時仍可產出；窄螢幕選單橫排；14 萬筆母體頁可開啟 |

---

## 11. 已知限制
1. 母體指紋只涵蓋鍵值與金額（為與 V1／V2 資料相容而保留）。
2. 大母體（十萬筆以上）開啟母體頁需十餘秒；明細清單只渲染前 100 列。
3. CSV 以 Excel 開啟時，前導 0 或超過 15 位的鍵值會被轉成數字，需保留原樣請用 Excel 匯出。
4. 附件存於瀏覽器，受瀏覽器配額限制；JSON 備份以 Base64 儲存，檔案約為原附件 1.37 倍。
5. 信賴水準、可容忍偏差率、可容忍誤述僅記錄，不做統計推論。
6. 單一使用者、單一瀏覽器；多人協作需以 JSON 合併匯入（同 id 會覆蓋）。

---

## 12. 版本紀錄
| 版本 | 內容 |
|---|---|
| V1 | 首版：六種抽樣方法、可重現種子、指紋、覆蓋地圖、佐證追蹤、Excel 底稿、列印報告、JSON 備份 |
| V2 | 修正：大母體 `Math.max` 爆堆疊、新計畫跳回舊結果、覆蓋地圖／追蹤誤報無抽樣、對話框連點、批次誤改隱藏列、快速切頁錯亂、涵蓋率不隨替補更新、工具列列印找不到抽樣。強化：CSV 公式注入、XML 控制字元、金額解析、鎖定母體刪除二次確認、抽不滿警告、替補期限、24 小時制、RWD |
| V3 | 修正：`-$1,000` 負數判為正、全選無法取消、刪除抽樣後跳到別的計畫、替補上場樣本列錯工作表、指紋匯出截斷。安全：屬性值全面 `esc()`、匯入 store 陣列檢查。改善：自然排序、替補關係欄、說明頁匯出注意事項 |

---

## 13. 擴充指引
- **新增抽樣方法**：在 `METHODS`、`METHOD_DESC` 加入定義，實作 `sXxx(rows, n, m, seed, ...)` 回傳 `{main, reserve, info, strata?}`，並於 `selectBy` 註冊；如要可作為 `keyplus` 的 baseMethod，同步加入白名單與編輯對話框選項。
- **新增 store／index**：修改 `SCHEMA` 並將 `APP.dbVersion` 加 1；若為業務資料同時加入 `DATA_STORES`（才會被備份、還原、清除）。
- **改版命名**：沿用 `dbName = bot-sample-lab`；檔名固定為 `佐證抽樣實驗室.html`，不加入版本號，版本只更新檔內的 `<title>`、標題與修正紀錄。
- **換色**：只需修改 `:root` 的 `--primary`、`--primary-d`、`--accent`、`--accent-l`；Excel 表頭色與熱力圖色分別在 `toExcelXML` 與 `heatmapSVG` 內。
