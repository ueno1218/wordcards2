// ====== 状態 ======
let allCards = [];
let deck = [];
let idx = 0;
let showBack = false;

let gradeFilter = null;          // null=未適用 / Set(A,B,C)=適用
let gradeFilterMode = false;     // 「評価で出題」ボタンで適用中か（タグ適用だけのときfalse）

let csvHeaders = [];
let csvRows = []; // 読み込んだ表データ（CSV/XLSX共通）

// 端末内に最終評価を保存（任意）
const GRADES_KEY = "wordcards_grades_v1";
let grades = loadGrades();

// 進捗保存（今回追加）
const STATE_KEY = "wordcards_state_v2";
let lastDatasetSig = "";           // 現在読み込んでいるデータのシグネチャ
let pendingRestore = null;         // 復元候補（ファイル読込後に適用）

let saveTimer = null;
const SAVE_DEBOUNCE_MS = 250;

// ====== UI ======
const el = (id) => document.getElementById(id);

const fileInput = el("csvFile");
const btnShuffle = el("btnShuffle");
const btnResetStats = el("btnResetStats");

const tagFilterInput = el("tagFilter");
const tagOptions = el("tagOptions");
const btnApplyFilter = el("btnApplyFilter");

const fA = el("fA");
const fB = el("fB");
const fC = el("fC");
const btnApplyGrade = el("btnApplyGrade");

const counter = el("counter");
const tagLabel = el("tagLabel");
const gradeLabel = el("gradeLabel");

const face = el("face");
const back = el("back");

const btnFlip = el("btnFlip");
const btnPrev = el("btnPrev");
const btnNext = el("btnNext");

const btnA = el("btnA");
const btnB = el("btnB");
const btnC = el("btnC");

const btnExport = el("btnExport");
const historyInput = el("historyFile");
const historyStatus = el("historyStatus");
const btnComment = el("btnComment");
const commentDialog = el("commentDialog");
const commentText = el("commentText");
const btnSaveComment = el("btnSaveComment");
const btnCancelComment = el("btnCancelComment");
const commentDisplay = el("commentDisplay");
const IDLE_MS = 5 * 60 * 1000;
let idleTimer = null;
let changedSinceExport = false;
let historyImportBusy = false;
const statsBox = el("stats");

// ====== 起動時：保存されている進捗を読み込んで「復元待ち」にしておく ======
window.addEventListener("load", () => {
  const s = loadState();
  if (s) pendingRestore = s;

  // タブが破棄される前に保存
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") saveStateNow();
  });
  window.addEventListener("pagehide", saveStateNow);
  // ページ非表示時にダウンロードを開始することはブラウザに拒否され得る。
  // 確実に保持できる端末内保存を優先する。
  for (const eventName of ["pointerdown", "keydown", "touchstart", "scroll"]) {
    window.addEventListener(eventName, resetIdleTimer, { passive: true });
  }
  resetIdleTimer();
});

// ====== ファイル読み込み（CSV/XLSX分岐） ======
fileInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  const name = (file.name || "").toLowerCase();

  try {
    if (name.endsWith(".xlsx")) {
      const { headers, rows } = await parseXlsxWithHeaders(file);
      applyLoadedRows(headers, rows);
    } else {
      // CSV（改行対応パーサ）
      const text = await file.text();
      const { headers, rows } = parseCsvWithHeaders(text);
      applyLoadedRows(headers, rows);
    }
  } catch (err) {
    console.error(err);
    alert("ファイルの読み込みに失敗しました。形式（列名やシート）をご確認ください。");
  }
});

// ====== XLSXパース（1枚目シート／1行目ヘッダ） ======
async function parseXlsxWithHeaders(file) {
  if (typeof XLSX === "undefined") {
    throw new Error("XLSX library not loaded.");
  }

  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });

  const sheetName = wb.SheetNames?.[0];
  if (!sheetName) return { headers: [], rows: [] };

  const ws = wb.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
  const headers = (matrix[0] || []).map(h => String(h ?? "").trim());
  const rows = matrix.slice(1).filter(r => r.some(v => String(v ?? "").trim() !== "")).map(r => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = r[i] ?? ""; });
    return obj;
  });
  return { headers, rows };
}

// ====== CSV/XLSX 共通の取り込み処理 ======
function applyLoadedRows(headers, rows) {
  if (!headers.includes("id") || !headers.includes("front") || !headers.includes("back")) {
    throw new Error("問題ファイルには id / front / back 列が必要です。");
  }
  csvHeaders = headers;
  csvRows = rows;
  // 前の問題ファイルの端末内評価を今回の読み込みへ持ち込まない。
  grades = {};
  saveGrades(grades);

  // データの簡易シグネチャ（復元一致判定用）
  lastDatasetSig = computeDatasetSignature(csvHeaders, csvRows);

  // 仕様：読み込み直後は（存在すれば）grade列をその値として採用
  allCards = csvRows.map((r, i) => {
    const id = normId(r.id, i);
    const g = normGrade(r.grade ?? "");
    return {
      id,
      front: String(r.front ?? "").trim(),
      back: String(r.back ?? "").trim(),
      tag: String(r.tag ?? "").trim(),
      grade: g,
      comment: String(r.comment ?? r["コメント"] ?? ""),
      _rowIndex: i
    };
  }).filter(c => c.front && c.back);

  rebuildTagOptions();

  // 問題ファイルのGradeが基準。履歴Excelの取り込み時だけ上書きする。
  changedSinceExport = false;
  if (historyStatus) historyStatus.textContent = "学習履歴Excel：未読み込み";
  // まず通常初期化
  gradeFilter = null;
  gradeFilterMode = false;
  idx = 0;
  showBack = false;

  // ★ 同一データなら表示位置・フィルタ・コメントのみ復元（Gradeは問題ファイル優先）
  if (pendingRestore && pendingRestore.datasetSig === lastDatasetSig) {
    restoreFromState(pendingRestore);
    pendingRestore = null;
    // restoreFromState 内で render まで行う
    scheduleSave(); // 復元直後に保存して整合させる
    return;
  }

  // 復元なし：通常開始
  pendingRestore = null;
  applyFilterAndStart();
  scheduleSave();
}

// ====== タグ候補生成（datalist更新） ======
function rebuildTagOptions() {
  tagOptions.innerHTML = "";

  const tags = new Set();
  for (const c of allCards) {
    const t = (c.tag ?? "").trim();
    if (t) tags.add(t);
  }

  const sorted = [...tags].sort((a, b) => a.localeCompare(b, "ja"));
  for (const t of sorted) {
    const opt = document.createElement("option");
    opt.value = t;
    tagOptions.appendChild(opt);
  }
}

// ====== フィルタ ======
btnApplyFilter.addEventListener("click", () => {
  gradeFilter = null;
  gradeFilterMode = false;
  applyFilterAndStart();
  scheduleSave();
});

btnApplyGrade.addEventListener("click", () => {
  gradeFilter = readGradeFilter();
  gradeFilterMode = true;
  applyFilterAndStart();
  scheduleSave();
});

function readGradeFilter() {
  const set = new Set();
  if (fA.checked) set.add("A");
  if (fB.checked) set.add("B");
  if (fC.checked) set.add("C");
  return set;
}

function applyFilterAndStart() {
  const t = (tagFilterInput.value || "").trim();

  let base = t ? allCards.filter(c => c.tag === t) : [...allCards];

  // 選択されたGradeに加え、未設定（空欄）の問題も出題する
  if (gradeFilter && gradeFilter.size > 0) {
    base = base.filter(c => gradeFilter.has(c.grade) || !c.grade);
  }

  deck = base;
  idx = 0;

  deck.length === 0 ? renderEmpty(t) : render();
}

// ====== カード操作 ======
btnFlip.addEventListener("click", () => {
  showBack = !showBack;
  render();
  scheduleSave();
});

btnPrev.addEventListener("click", () => {
  if (!deck.length) return;
  idx = (idx - 1 + deck.length) % deck.length;
  showBack = false;
  render();
  scheduleSave();
});

btnNext.addEventListener("click", () => {
  if (!deck.length) return;
  idx = (idx + 1) % deck.length;
  showBack = false;
  render();
  scheduleSave();
});

btnShuffle.addEventListener("click", () => {
  if (!deck.length) return;
  shuffle(deck);
  idx = 0;
  showBack = false;
  render();
  scheduleSave();
});

// ====== 評価(A/B/C)更新：更新だけ（裏表示・自動遷移はしない） ======
btnA.addEventListener("click", () => updateGradeOnly("A"));
btnB.addEventListener("click", () => updateGradeOnly("B"));
btnC.addEventListener("click", () => updateGradeOnly("C"));

function updateGradeOnly(g) {
  if (!deck.length) return;

  const c = deck[idx];

  // grade更新（最終更新が評価値）
  c.grade = g;

  // 書き出し用に、元の行データにもgradeを更新
  const r = csvRows[c._rowIndex];
  r.grade = g;

  // 端末内保存（任意）
  changedSinceExport = true;
  resetIdleTimer();
  grades[`${c.id}\u0001${c.tag}`] = g;
  saveGrades(grades);

  // A/B/Cを押したら、カードの裏面を表示する
  // すでに表示されている場合は、そのまま
  showBack = true;

  // 表示更新
  // ※カード位置 idx は変えない
  // ※自動で次へ進まない
  render();

  // 途中経過を保存
  scheduleSave();
}


// ====== 学習履歴リセット：現在読み込んでいる全カードのGradeを未設定に ======
btnResetStats.addEventListener("click", () => {
  if (!allCards.length) {
    alert("先に問題ファイルを読み込んでください。");
    return;
  }
  if (!confirm("現在読み込んでいる全カードのGrade（A/B/C）を未設定に戻します。\\n問題文・回答・タグは変更しません。よろしいですか？")) return;

  // フィルタ中でも、対象は現在のデッキだけでなく全問題。
  for (const c of allCards) {
    c.grade = "";
    if (csvRows[c._rowIndex]) csvRows[c._rowIndex].grade = "";
  }

  // 古いGradeが復活しないよう端末内の評価記録も消去する。
  grades = {};
  saveGrades(grades);
  pendingRestore = null;

  changedSinceExport = true;
  // フィルタや出題順を維持しつつ表示を更新。
  if (deck.length) render();
  else renderEmpty((tagFilterInput.value || "").trim());

  // 遅延保存ではなく直ちに保存し、再読み込みでもリセット状態を維持。
  if (saveTimer) {
    clearTimeout(saveTimer);
    saveTimer = null;
  }
  saveStateNow();
  resetIdleTimer();
  if (historyStatus) historyStatus.textContent = `学習履歴リセット：${allCards.length}件を未設定に変更（未書き出し）`;
  alert(`${allCards.length}件のGradeをリセットしました。\\n必要なら「学習履歴Excelに書き出し」で保存してください。`);
});

// ====== カードごとのコメント編集 ======
let editingCard = null;
btnComment.addEventListener("click", () => {
  if (!deck.length) return;
  editingCard = deck[idx];
  commentText.value = editingCard.comment ?? "";
  commentDialog.showModal();
  commentText.focus();
});
btnCancelComment.addEventListener("click", () => commentDialog.close());
btnSaveComment.addEventListener("click", () => {
  if (!editingCard) return;
  const comment = commentText.value;
  editingCard.comment = comment;
  csvRows[editingCard._rowIndex].comment = comment;
  changedSinceExport = true;
  commentDialog.close();
  editingCard = null;
  if (deck.length) render();
  saveStateNow();
  resetIdleTimer();
});
commentDialog.addEventListener("close", () => { editingCard = null; });

// ====== 学習履歴Excel：ID / タグ / grade / コメント ======
function cardKey(id, tag) {
  return JSON.stringify([String(id ?? "").trim(), String(tag ?? "").trim()]);
}

function historyFilename(date = new Date()) {
  const p = n => String(n).padStart(2, "0");
  return `LatestGrade_${date.getFullYear()}${p(date.getMonth()+1)}${p(date.getDate())}_${p(date.getHours())}${p(date.getMinutes())}${p(date.getSeconds())}.xlsx`;
}

function buildHistoryWorkbook() {
  if (typeof XLSX === "undefined") throw new Error("Excelライブラリが利用できません。");
  const data = [["ID", "タグ", "grade", "コメント"]];
  allCards.forEach(c => data.push([c.id, c.tag, c.grade, c.comment ?? ""]));
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = [{wch:22},{wch:24},{wch:12},{wch:60}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "LatestGrade");
  return wb;
}

function exportHistory(manual = true) {
  if (!allCards.length) {
    if (manual) alert("先に問題ファイルを読み込んでください。");
    return false;
  }
  try {
    const filename = historyFilename();
    XLSX.writeFile(buildHistoryWorkbook(), filename);
    changedSinceExport = false;
    saveStateNow();
    if (historyStatus) historyStatus.textContent = `書き出しを開始：${filename}`;
    return true;
  } catch (e) {
    console.error("履歴Excelの書き出しに失敗", e);
    if (manual) alert(`書き出しに失敗しました：${e.message}`);
    return false;
  }
}

btnExport.addEventListener("click", () => exportHistory(true));

historyInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  if (!allCards.length) {
    alert("先に左上から問題ファイルを読み込んでください。");
    historyInput.value = "";
    return;
  }
  if (historyImportBusy) return;
  historyImportBusy = true;
  try {
    const name = file.name.toLowerCase();
    let parsed;
    if (name.endsWith(".xlsx")) parsed = await parseXlsxWithHeaders(file);
    else if (name.endsWith(".csv")) parsed = parseCsvWithHeaders(await file.text());
    else throw new Error("XLSXまたはCSVファイルを指定してください。");
    const {headers, rows} = parsed;
    const idColumn = headers.find(h => ["ID", "id"].includes(h));
    const tagColumn = headers.find(h => ["タグ", "tag"].includes(h));
    const gradeColumn = headers.find(h => h.toLowerCase() === "grade");
    const commentColumn = headers.find(h => ["コメント", "comment"].includes(h.toLowerCase()));
    if (!idColumn || !tagColumn || !gradeColumn) {
      throw new Error("学習履歴ファイルには ID・タグ・grade の3列が必要です。");
    }
    const lookup = new Map();
    for (const r of rows) {
      const key = cardKey(r[idColumn], r[tagColumn]);
      if (lookup.has(key)) throw new Error(`学習履歴にIDとタグの重複があります：${r[idColumn]} / ${r[tagColumn]}`);
      const raw = String(r[gradeColumn] ?? "").trim();
      const g = normGrade(raw);
      if (raw && !g) throw new Error(`不正なGrade：${r[idColumn]} / ${r[tagColumn]} / ${raw}`);
      lookup.set(key, {grade: g, comment: commentColumn ? String(r[commentColumn] ?? "") : null});
    }
    const originalKeys = new Set(allCards.map(c => cardKey(c.id,c.tag)));
    let updated = 0;
    for (const c of allCards) {
      const key = cardKey(c.id,c.tag);
      if (!lookup.has(key)) continue;
      const item = lookup.get(key);
      const g = item.grade;
      c.grade = g;
      if (item.comment !== null) {
        c.comment = item.comment;
        csvRows[c._rowIndex].comment = item.comment;
      }
      csvRows[c._rowIndex].grade = g;
      grades[key] = g;
      updated++;
    }
    saveGrades(grades);
    const unmatched = [...lookup.keys()].filter(k => !originalKeys.has(k)).length;
    changedSinceExport = true;
    // 表示中のデッキの順序と現在位置は保つ。
    if (deck.length) render();
    else renderStats();
    scheduleSave();
    if (historyStatus) historyStatus.textContent = `学習履歴：${updated}件反映 / 対象外${unmatched}件`;
    alert(`学習履歴を読み込みました。\\n反映：${updated}件\\n一致しなかった履歴：${unmatched}件\\n履歴のない問題は元のGradeを維持します。`);
  } catch (e) {
    console.error(e);
    alert(`学習履歴の読み込みに失敗しました：${e.message}`);
  } finally {
    historyImportBusy = false;
    historyInput.value = "";
  }
});

// 無操作中、画面が表示されている間に書き出しを試行する。
// ブラウザが自動ダウンロードを拒否した場合でもlocalStorage保存は残る。
function resetIdleTimer() {
  if (idleTimer) clearTimeout(idleTimer);
  idleTimer = setTimeout(() => {
    saveStateNow();
    if (changedSinceExport && document.visibilityState === "visible") {
      exportHistory(false);
    }
  }, IDLE_MS);
}

// ====== 描画 ======
function render() {
  const c = deck[idx];

  counter.textContent = `${idx + 1} / ${deck.length}`;
  tagLabel.textContent = c.tag ? `#${c.tag}` : "";

  // 右上に grade 表示（tagの右側）
  gradeLabel.textContent = c.grade ? `Grade:${c.grade}` : "Grade:-";

  face.textContent = c.front;
  back.textContent = c.back;

  // invisible のみで表示制御（位置はズレない）
  back.classList.toggle("invisible", !showBack);
  commentDisplay.textContent = c.comment ? `【コメント】\\n${c.comment}` : "";
  commentDisplay.classList.toggle("hidden", !showBack || !c.comment);
  btnComment.textContent = c.comment ? "コメントを編集" : "コメント";

  btnFlip.textContent = showBack ? "表に戻す" : "裏を見る";

  renderStats();
}

function renderEmpty(tag) {
  counter.textContent = "0 / 0";
  tagLabel.textContent = tag ? `#${tag}` : "";
  gradeLabel.textContent = "";

  face.textContent = "該当するカードがありません";
  back.textContent = "";
  back.classList.add("invisible");
  commentDisplay.textContent = "";
  commentDisplay.classList.add("hidden");
  btnComment.textContent = "コメント";

  renderStats();
}

function renderStats() {
  if (!allCards.length) {
    statsBox.textContent = "CSV/XLSX未読込";
    return;
  }

  let A = 0, B = 0, C = 0, U = 0;
  for (const c of allCards) {
    if (c.grade === "A") A++;
    else if (c.grade === "B") B++;
    else if (c.grade === "C") C++;
    else U++;
  }

  const gf = (gradeFilter && gradeFilter.size > 0) ? ` / 出題:${[...gradeFilter].join(",")}` : "";
  statsBox.textContent = `現在評価：A ${A} / B ${B} / C ${C} / 未設定 ${U}${gf}`;
}

// ====== 進捗保存・復元（今回追加） ======
function scheduleSave() {
  if (!lastDatasetSig) return; // まだデータ未読込
  if (saveTimer) clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    saveTimer = null;
    saveStateNow();
  }, SAVE_DEBOUNCE_MS);
}

function saveStateNow() {
  if (!lastDatasetSig) return; // まだデータ未読込
  const state = {
    v: 2,
    savedAt: Date.now(),
    datasetSig: lastDatasetSig,
    // 復元時にID・タグごとに最新Gradeを反映する
    gradeEntries: allCards.map(c => [c.id, c.tag, c.grade]),
    commentEntries: allCards.map(c => [c.id, c.tag, c.comment ?? ""]),

    // フィルタ
    tagFilter: (tagFilterInput.value || "").trim(),
    gradeChecks: { A: !!fA.checked, B: !!fB.checked, C: !!fC.checked },
    gradeFilterMode: !!gradeFilterMode, // 「評価で出題」を押して適用中か

    // デッキ順・現在位置
    deckIds: Array.isArray(deck) ? deck.map(c => c.id) : [],
    idx: Number.isFinite(idx) ? idx : 0
  };

  try {
    localStorage.setItem(STATE_KEY, JSON.stringify(state));
  } catch (e) {
    // 容量不足などで保存できない場合がある
    console.warn("Failed to save state:", e);
  }
}

function loadState() {
  try {
    const raw = localStorage.getItem(STATE_KEY);
    if (!raw) return null;
    const s = JSON.parse(raw);
    if (!s || !s.datasetSig) return null;
    return s;
  } catch {
    return null;
  }
}

function restoreFromState(s) {
  if (Array.isArray(s.commentEntries)) {
    const comments = new Map(s.commentEntries.map(([id, tag, comment]) => [cardKey(id, tag), String(comment ?? "")]));
    for (const c of allCards) {
      const key = cardKey(c.id, c.tag);
      if (comments.has(key)) {
        c.comment = comments.get(key);
        csvRows[c._rowIndex].comment = c.comment;
      }
    }
  }
  // 問題ファイルを明示的に読み込んだ際は、そのgrade列を最優先する。
  // 保存済みのgradeEntriesで上書きしない（コメント・出題位置等は復元可）。
  // フィルタUIの復元
  tagFilterInput.value = s.tagFilter || "";
  const checks = s.gradeChecks || {};
  fA.checked = checks.A !== false; // 未設定ならtrue寄り（以前のデフォルト）
  fB.checked = checks.B !== false;
  fC.checked = checks.C !== false;

  // フィルタモードの復元
  gradeFilterMode = !!s.gradeFilterMode;
  if (gradeFilterMode) {
    gradeFilter = readGradeFilter();
  } else {
    gradeFilter = null;
  }

  // フィルタ適用して deck を作る
  const t = (tagFilterInput.value || "").trim();
  let base = t ? allCards.filter(c => c.tag === t) : [...allCards];
  if (gradeFilter && gradeFilter.size > 0) {
    base = base.filter(c => gradeFilter.has(c.grade) || !c.grade);
  }
  deck = base;

  // デッキ順の復元（保存されたdeckIds順に並び替え）
  const deckIds = Array.isArray(s.deckIds) ? s.deckIds : [];
  if (deckIds.length && deck.length) {
    const map = new Map(deck.map(c => [c.id, c]));
    const ordered = [];
    for (const id of deckIds) {
      const c = map.get(id);
      if (c) ordered.push(c);
    }
    // 新規/漏れがあれば後ろに付ける
    for (const c of deck) {
      if (!deckIds.includes(c.id)) ordered.push(c);
    }
    deck = ordered;
  }

  // idx復元
  const savedIdx = Number.isFinite(s.idx) ? s.idx : 0;
  idx = deck.length ? Math.max(0, Math.min(savedIdx, deck.length - 1)) : 0;

  showBack = false;
  deck.length === 0 ? renderEmpty(t) : render();
}

// ====== データ一致判定用の簡易シグネチャ ======
function computeDatasetSignature(headers, rows) {
  // 安価に「同じファイル内容か」を判定するためのシグネチャ
  // 先頭・末尾のid/front/back/tagと件数・ヘッダで作る（Gradeは除外）
  const h = Array.isArray(headers) ? headers.join("|") : "";
  const n = Array.isArray(rows) ? rows.length : 0;

  const pick = (r) => {
    if (!r) return "";
    const id = String(r.id ?? "").trim();
    const front = String(r.front ?? "").trim();
    const back = String(r.back ?? "").trim();
    const tag = String(r.tag ?? "").trim();
    return `${id}#${front}#${back}#${tag}`;
  };

  const first = n ? pick(rows[0]) : "";
  const mid = n ? pick(rows[Math.floor(n / 2)]) : "";
  const last = n ? pick(rows[n - 1]) : "";

  // 文字列を軽くハッシュ化
  const raw = `${h}||${n}||${first}||${mid}||${last}`;
  return simpleHash(raw);
}

function simpleHash(str) {
  // 軽量なハッシュ（FNV-1a風）
  let h = 2166136261;
  for (let i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return (h >>> 0).toString(16);
}

// ====== CSVパース（改行対応・ヘッダ保持） ======
function parseCsvWithHeaders(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  text = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];

    if (ch === '"') {
      if (inQuotes && next === '"') { cell += '"'; i++; }
      else inQuotes = !inQuotes;
    } else if (ch === "," && !inQuotes) {
      row.push(cell); cell = "";
    } else if (ch === "\n" && !inQuotes) {
      row.push(cell); rows.push(row);
      row = []; cell = "";
    } else {
      cell += ch;
    }
  }
  row.push(cell); rows.push(row);

  if (rows.length === 0) return { headers: [], rows: [] };

  const headers = rows[0].map(h => String(h ?? "").trim());
  const out = [];

  for (let i = 1; i < rows.length; i++) {
    if (rows[i].length === 1 && rows[i][0] === "") continue;
    const obj = {};
    for (let j = 0; j < headers.length; j++) obj[headers[j]] = rows[i][j] ?? "";
    out.push(obj);
  }
  return { headers, rows: out };
}

// ====== 正規化 ======
function normId(rawId, i) {
  const s = String(rawId ?? "").trim();
  return s || String(i + 1);
}
function normGrade(g) {
  const s = String(g ?? "").trim().toUpperCase();
  return (s === "A" || s === "B" || s === "C") ? s : "";
}

// ====== ユーティリティ ======
function shuffle(a) {
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
}
function rowsToCsv(rows) {
  const esc = (v) => {
    const s = String(v ?? "");
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  return rows.map(r => r.map(esc).join(",")).join("\n");
}
function downloadCsv(csvText, filename) {
  const blob = new Blob([csvText], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// ====== localStorage（任意） ======
function loadGrades() {
  try { return JSON.parse(localStorage.getItem(GRADES_KEY) || "{}"); }
  catch { return {}; }
}
function saveGrades(obj) {
  localStorage.setItem(GRADES_KEY, JSON.stringify(obj));
}

// ====== PWA登録（サービスワーカー） ======
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./sw.js").catch(() => {});
  });
}
