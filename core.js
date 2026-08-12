/* 工作台｜核心（共用）
   登入、Graph 讀寫、文件庫讀取、共用小工具。

   每個頁面必須在載入這個檔案之前自行宣告 REDIRECT_URI，
   因為 Entra 對重新導向 URI 是完全比對，兩個頁面的網址不同，
   必須各自註冊、各自宣告。寫在共用檔裡會讓其中一頁永遠登不進去。

   每個頁面必須自行定義 onDataLoaded()，資料讀完後由 loadData() 呼叫。 */

/* ================================================================
   一、設定區　—　所有識別碼集中在這裡，不散在各模組
   ================================================================ */
const MSAL_CONFIG = {
  auth: {
    clientId: 'bec4f97b-ca7e-4841-94e1-d514f27d71bb',
    authority: 'https://login.microsoftonline.com/c202d9ba-2e47-4aff-88e1-1fb71692a54b',
    redirectUri: REDIRECT_URI
  },
  cache: { cacheLocation: 'sessionStorage' }
};

const GRAPH_SITE_ID = 'texray.sharepoint.com,523489e7-e84e-4554-96f9-2a27018db21a,fc5dd947-527e-47f8-8762-02c81a22d02a';

// 新模組建好清單後，把清單 ID 填進來即可，其他地方不用改
const LISTS = {
  exhibitions:  '01765d4f-5cb9-4a5b-b04f-a9be79f5960e',
  requests:     '',
  todos:        '',
  annual:       '',
  projects:     '',
  projectItems: '',
  external:     '',
  categories:   '',
  people:       ''
};

// 展覽檔案放在同一個網站的預設文件庫，路徑固定為
//   Shared Documents/展覽/{年度}/{FolderKey}/{01_設計產出,02_現場照片,03_結案報告}
// 所以連結是「算」出來的，清單裡不再存四個連結欄位。
// 已於 2026-08-11 用 Graph 確認過實際路徑，不是推測。
const LIB_ROOT = 'https://texray.sharepoint.com/sites/internal-tools/Shared%20Documents';
const EX_FOLDER_ROOT = LIB_ROOT + '/' + encodeURIComponent('展覽');
const EX_SUBFOLDERS = { design: '01_設計產出', photo: '02_現場照片', report: '03_結案報告' };

// 先要可寫入的權限；若尚未經組織同意，自動退回唯讀，網頁仍可正常使用
const SCOPE_RW = ['https://graph.microsoft.com/Sites.ReadWrite.All'];
const SCOPE_RO = ['https://graph.microsoft.com/Sites.Read.All'];


/* ================================================================
   二、殼的共用函式　—　登入、讀取、寫入、回饋、導覽
      這一段是一次性的，第 N 個模組都用同一套
   ================================================================ */
let msalApp = null;
let activeAccount = null;
let canWrite = false;

// LINE／Teams 等內建瀏覽器常擋 window.open，用來決定要不要顯示「改用 Safari 開啟」的提示。
function isInAppBrowser() {
  const ua = navigator.userAgent || '';
  return /Line\/|FBAN|FBAV|Instagram|MicroMessenger|Teams|wv\)/i.test(ua);
}

// 靜默偵測有沒有可寫入權限。IT 核准後這裡會自動成功，不需要改程式或重新授權。
async function probeWrite() {
  try {
    await msalApp.acquireTokenSilent({ scopes: SCOPE_RW, account: activeAccount });
    canWrite = true;
  } catch (e) {
    canWrite = false;
  }
}

async function getToken() {
  const scopes = canWrite ? SCOPE_RW : SCOPE_RO;
  try {
    return (await msalApp.acquireTokenSilent({ scopes, account: activeAccount })).accessToken;
  } catch (e) {
    // 只用整頁轉向。不用彈出視窗——因為重新導向 URI 是根目錄，
    // 彈出視窗會把整套系統在裡面再跑一次，兩邊搶同一份授權碼。
    await msalApp.acquireTokenRedirect({ scopes });
    return new Promise(() => {});   // 頁面即將跳離
  }
}

function setModeBadge() {
  const b = document.getElementById('mode-badge');
  if (!b) return;
  b.textContent = canWrite ? '可寫入' : '唯讀';
  b.className = 'mode-badge ' + (canWrite ? 'mode-rw' : 'mode-ro');
}

function listUrl(key) {
  return 'https://graph.microsoft.com/v1.0/sites/' + GRAPH_SITE_ID + '/lists/' + LISTS[key];
}

// 讀任何一張清單
async function spGet(key) {
  if (!LISTS[key]) throw new Error('清單「' + key + '」尚未建立');
  const token = await getToken();
  const res = await fetch(listUrl(key) + '/items?expand=fields&$top=500',
                          { headers: { Authorization: 'Bearer ' + token } });
  if (!res.ok) throw new Error('Graph ' + res.status + '：' + (await res.text()).slice(0, 300));
  return (await res.json()).value;
}

// 改任何一張清單的任何欄位（權限下來後即可運作）
async function spPatch(key, itemId, fields) {
  if (!canWrite) { toast('目前是唯讀模式，寫入權限尚未核准', true); return null; }
  const token = await getToken();
  const res = await fetch(listUrl(key) + '/items/' + itemId + '/fields', {
    method: 'PATCH',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify(fields)
  });
  if (!res.ok) { toast('儲存失敗（' + res.status + '）', true); throw new Error(await res.text()); }
  toast('已儲存');
  return res.json();
}

// 新增一筆
async function spCreate(key, fields) {
  if (!canWrite) { toast('目前是唯讀模式，寫入權限尚未核准', true); return null; }
  const token = await getToken();
  const res = await fetch(listUrl(key) + '/items', {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ fields })
  });
  if (!res.ok) { toast('新增失敗（' + res.status + '）', true); throw new Error(await res.text()); }
  toast('已新增');
  return res.json();
}

// 軟刪除：各模組共用同一個欄位名稱
async function spSoftDelete(key, itemId) {
  return spPatch(key, itemId, { IsDeleted: true });
}

// ===== 文件庫（唯讀就夠，不需要等寫入權限）=====
// 已在正式網址上用網頁自己的 Sites.Read.All 權杖實測三種情況：
//   有內容 → 200 帶陣列 ／ 空資料夾 → 200 零筆 ／ 路徑不存在 → 404
// 三種都要分開處理，因為「空的」與「找不到」對使用者的意思完全不同。
function drivePath(segments) {
  return segments.map(encodeURIComponent).join('/');
}

const folderCache = {};   // 路徑 → { state, items }

async function driveList(segments) {
  const path = segments.join('/');
  if (folderCache[path]) return folderCache[path];
  const token = await getToken();
  const url = 'https://graph.microsoft.com/v1.0/sites/' + GRAPH_SITE_ID
            + '/drive/root:/' + drivePath(segments)
            + ':/children?$select=name,size,lastModifiedDateTime,webUrl,folder,file&$orderby=name';
  let result;
  try {
    const res = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    if (res.status === 404)      result = { state: 'missing', items: [] };
    else if (!res.ok)            result = { state: 'error', items: [], code: res.status };
    else {
      const items = (await res.json()).value.filter(x => !x.folder);
      result = { state: items.length ? 'ok' : 'empty', items };
    }
  } catch (e) {
    result = { state: 'error', items: [], code: e.message };
  }
  folderCache[path] = result;
  return result;
}

function fmtSize(b) {
  if (!b) return '';
  if (b < 1024) return b + ' B';
  if (b < 1024 * 1024) return Math.round(b / 1024) + ' KB';
  return (b / 1048576).toFixed(1) + ' MB';
}

function fmtStamp(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d)) return '';
  const p = n => String(n).padStart(2, '0');
  return d.getFullYear() + '/' + p(d.getMonth() + 1) + '/' + p(d.getDate()) + ' ' + p(d.getHours()) + ':' + p(d.getMinutes());
}

let toastTimer = null;
function toast(msg, isErr) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className = 'show' + (isErr ? ' err' : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.className = ''; }, isErr ? 4000 : 1800);
}

/* ================================================================
   五、啟動
   ================================================================ */
function showError(msg) {
  document.querySelectorAll('.module').forEach(m => m.classList.add('hidden'));
  document.getElementById('signin').classList.add('hidden');
  const loading = document.getElementById('loading');
  loading.classList.remove('hidden');
  loading.innerHTML = '載入失敗：' + msg
    + '<br><button class="back-btn" style="margin-top:12px;" onclick="location.reload()">重新載入</button>'
    + '<br><span class="empty-note" style="display:inline-block; margin-top:8px;">'
    + '若一直失敗，把這個畫面截圖下來，錯誤訊息就在上面那一行。</span>';
}

function showSignIn() {
  document.getElementById('loading').classList.add('hidden');
  document.querySelectorAll('.module').forEach(m => m.classList.add('hidden'));
  document.getElementById('signin').classList.remove('hidden');
  const hintEl = document.getElementById('signin-hint');
  if (hintEl) hintEl.innerHTML = isInAppBrowser()
    ? '你目前是在 LINE 或 Teams 的內建瀏覽器裡開啟。<br>'
      + '若登入卡住，請用選單選「用 Safari 開啟」或「用預設瀏覽器開啟」。'
    : '';
}

// 由使用者點按觸發，整頁跳去 Microsoft 登入再跳回來。
// 刻意不用彈出視窗：重新導向 URI 是根目錄，彈出視窗會讓整套系統在裡面再跑一次。
async function signIn() {
  const hint = document.getElementById('signin-hint');
  if (hint) hint.textContent = '正在前往 Microsoft 登入頁…';
  try {
    await msalApp.loginRedirect({ scopes: SCOPE_RO });
  } catch (e) {
    if (hint) hint.innerHTML = '登入失敗：' + (e.message || e);
  }
}

async function loadData() {
  const loading = document.getElementById('loading');
  loading.classList.remove('hidden');
  try {
    if (msalApp.setActiveAccount) msalApp.setActiveAccount(activeAccount);

    // 靜默確認寫入權限。核准前為唯讀，核准後自動變可寫入。
    await probeWrite();
    setModeBadge();

    loading.textContent = '資料載入中…';
    const raw = await spGet('exhibitions');
    // 軟刪除：骨幹規劃五之二第 4 條，各模組一律過濾掉 IsDeleted
    exhibitions = raw.map(it => normalize(it.fields || {}, it.id))
      .filter(e => e['展覽名稱'] && e['已刪除'] !== '是');

    loading.classList.add('hidden');
    document.getElementById('updated-label').textContent =
      '更新時間：' + new Date().toLocaleString('zh-TW');
    // 資料讀完之後做什麼由各頁面決定：工作台去儀表板，展覽檢視直接顯示展覽
    onDataLoaded();
  } catch (err) {
    showError(err.message || err);
    console.error(err);
  }
}

async function boot() {
  const loading = document.getElementById('loading');
  try {
    if (typeof msal === 'undefined') {
      throw new Error('登入函式庫（MSAL）未載入，可能是公司網路或防火牆擋住了 cdn.jsdelivr.net');
    }
    msalApp = new msal.PublicClientApplication(MSAL_CONFIG);
    if (msalApp.initialize) await msalApp.initialize();

    // 從整頁轉向登入回來時，先處理回應
    try {
      const r = await msalApp.handleRedirectPromise();
      if (r && r.account) activeAccount = r.account;
    } catch (e) {
      console.warn('handleRedirectPromise:', e && e.message);
      const h = document.getElementById('signin-hint');
      if (h) h.innerHTML = '上次的跳轉登入沒有完成：' + (e.message || e);
    }
    if (!activeAccount) activeAccount = msalApp.getAllAccounts()[0] || null;

    if (!activeAccount) { showSignIn(); return; }
    await loadData();
  } catch (err) {
    showError(err.message || err);
    console.error(err);
  }
}

boot();
