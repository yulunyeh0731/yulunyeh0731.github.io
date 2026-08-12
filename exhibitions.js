/* 工作台｜展覽模組（共用）
   兩個頁面共用：根目錄的工作台、/exhibitions/ 的展覽檢視。
   版面與繪製程式都只存在這一個檔案——這是抽出來的唯一理由。
   如果版面複製成兩份，改一邊忘記另一邊只是時間問題。

   相依：core.js（getToken／spGet／spPatch／canWrite／toast）。
   掛載：mountExhibitions(容器元素)。 */

const EX_MARKUP = `

      <div class="year-tabs" id="year-tabs"></div>

      <div id="view-main" class="hidden">
        <div class="card">
          <p class="section-title" id="stats-year-label">年度展覽進度</p>
          <div class="stats-grid">
            <div class="stat-box stat-total"><p>全年展覽數</p><p id="stat-total">-</p></div>
            <div class="stat-box stat-confirmed"><p>確定參加</p><p id="stat-confirmed">-</p></div>
            <div class="stat-box stat-evaluating"><p>評估中</p><p id="stat-evaluating">-</p></div>
            <div class="stat-box stat-done"><p>已結案</p><p id="stat-done">-</p></div>
          </div>
          <p id="next-ex" style="font-size:13px; margin:14px 0 0; padding-top:12px; border-top:1px solid var(--border); color:var(--text-secondary);"></p>
        </div>

        <div class="card">
          <p class="section-title">展覽一覽　·　點任一筆查看詳情</p>
          <div id="group-confirmed"></div>
          <div id="group-evaluating"></div>
          <div id="group-done"></div>
          <div id="group-skipped"></div>
        </div>
      </div>

      <div id="view-detail" class="hidden">
        <button class="back-btn" onclick="showMain()">← 返回總覽</button>
        <div class="card" style="margin-top:10px;">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:16px;">
            <span id="d-name" style="font-size:18px; font-weight:700;"></span>
            <span id="d-status" class="tag" style="margin-left:auto;"></span>
          </div>
          <div class="hero-grid">
            <div class="hero-box"><p><span class="hero-icon">📅</span>展期</p><p id="d-date"></p></div>
            <div class="hero-box"><p><span class="hero-icon">📍</span>地點</p><p id="d-place"></p></div>
            <div class="hero-box"><p><span class="hero-icon">🏷️</span>攤位</p><p id="d-booth"></p></div>
            <div class="hero-box"><p><span class="hero-icon">📐</span>攤位大小</p><p id="d-boothsize"></p></div>
          </div>
        </div>

        <div class="card">
          <div class="detail-tabs">
            <div class="detail-tab active" data-tab="basic" onclick="switchTab('basic')">基本資訊</div>
            <div class="detail-tab" data-tab="budget" onclick="switchTab('budget')">預算與補助</div>
            <div class="detail-tab" data-tab="files" onclick="switchTab('files')">文件與素材</div>
          </div>

          <div class="tab-panel" id="tab-basic">
            <p class="section-title" style="margin-top:0;">前置階段</p>
            <div id="d-stages" style="display:flex; gap:6px; margin-bottom:18px; flex-wrap:wrap;"></div>
            <div class="detail-row" style="margin-bottom:16px;">
              <div><span>展覽類型</span><span id="d-scope"></span></div>
              <div><span>業務參展人</span><span id="d-salesperson"></span></div>
            </div>
            <p class="section-title">備註說明</p>
            <p id="d-note" style="font-size:13px; margin:0; color:var(--text-secondary);"></p>
          </div>

          <div class="tab-panel hidden" id="tab-budget">
            <p class="section-title" style="margin-top:0;">費用</p>
            <div class="detail-row" style="margin-bottom:6px;">
              <div><span>攤位費用</span><span id="d-fee-booth"></span></div>
              <div><span>運費</span><span id="d-fee-freight"></span></div>
              <div><span>裝置費用</span><span id="d-fee-setup"></span></div>
              <div><span>合計</span><span id="d-fee-total"></span></div>
            </div>
            <p class="empty-note" style="margin:0 0 18px;">空白不等於 0，代表由別的單位支付或還沒確定，不計入合計。</p>
            <p class="section-title">補助</p>
            <div class="detail-row">
              <div><span>補助單位</span><span id="d-subsidy-org"></span></div>
              <div><span>核銷狀態</span><span id="d-subsidy-status"></span></div>
              <div><span>申請金額</span><span id="d-subsidy-apply"></span></div>
              <div><span>核准金額</span><span id="d-subsidy-approved"></span></div>
            </div>
          </div>

          <div class="tab-panel hidden" id="tab-files">
            <p class="section-title" style="margin-top:0;">設計產出</p>
            <div id="d-design"></div>
            <div id="d-design-files" style="margin-bottom:18px;"></div>
            <p class="section-title">現場照片</p>
            <div id="d-photos"></div>
            <div id="d-photo-files" style="margin-bottom:18px;"></div>
            <p class="section-title">結案報告</p>
            <div id="d-report"></div>
            <div id="d-report-files"></div>
            <div id="d-closing" style="display:flex; gap:6px; margin-top:10px; flex-wrap:wrap; align-items:center;"></div>
            <p class="empty-note" style="margin:18px 0 0;">檔案直接上傳到文件庫。在資料夾按「新增 OneDrive 捷徑」，之後從檔案總管拖進去就好。</p>
          </div>
        </div>
      </div>
    </div>
`;

function mountExhibitions(el) {
  el.innerHTML = EX_MARKUP;
}

/* ================================================================
   三、模組：展覽　—　繪製邏輯沿用 sp.html，未修改行為
   ================================================================ */
// 左邊是 SharePoint 內部名稱（建立後改不掉），右邊只是這支程式內部用的鍵名。
// 對照的是 2026-08-11 重建的 Exhibitions，26 欄。
const FIELD_MAP = {
  Title: '展覽名稱', Year: '年度', Scope: '展覽類型', Status: '參展狀態',
  StartDate: '展期開始', EndDate: '展期結束',
  City: '地點城市', Country: '國家',
  BoothNo: 'BoothNo', BoothSize: 'BoothSize',
  SalesPerson: '業務參展人',
  RegDone: '報名', PayDone: '繳費', BoothDone: '選位', PackDone: '備品打包',
  BoothFee: '攤位費用', FreightFee: '運費', SetupFee: '裝置費用',
  SubsidyOrg: '補助單位', SubsidyApplied: '補助申請金額',
  SubsidyApproved: '補助核准金額', SubsidyStatus: '補助核銷狀態',
  ClosingReportDone: '結案簡報完成',
  Remarks: '備註說明',
  FolderKey: '資料夾代號', IsDeleted: '已刪除'
};
const BOOL_FIELDS = ['RegDone', 'PayDone', 'BoothDone', 'PackDone', 'ClosingReportDone', 'IsDeleted'];

// 前置階段四顆。第四顆的字面依「展覽類型」而變：
// 國內＝自己布展，國外＝要出貨。同一個 PackDone 欄位，兩種說法。
const STAGE_FIELDS = { RegDone: '報名', PayDone: '繳費', BoothDone: '選位', PackDone: '備品打包' };
function stageLabel(en, ex) {
  if (en !== 'PackDone') return STAGE_FIELDS[en];
  return ex['展覽類型'] === '國內' ? '布展' : '運輸';
}

let exhibitions = [];
let currentYear = null;

function toSlashDate(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d)) return String(iso);
  const p = n => String(n).padStart(2, '0');
  return d.getUTCFullYear() + '/' + p(d.getUTCMonth() + 1) + '/' + p(d.getUTCDate());
}

function normalize(f, id) {
  const o = { _id: id };
  for (const en in FIELD_MAP) {
    const zh = FIELD_MAP[en];
    let v = f[en];
    if (v === undefined || v === null) { o[zh] = ''; continue; }
    if (en === 'StartDate' || en === 'EndDate') o[zh] = toSlashDate(v);
    else if (BOOL_FIELDS.indexOf(en) >= 0) o[zh] = (v === true) ? '是' : '否';
    else if (typeof v === 'object') o[zh] = v.Url || '';
    else o[zh] = String(v);
  }
  return o;
}

// 已結案＝已完成（綠），暫不參加＝不適用（灰）。
// 兩者原本都是灰，意思完全相反卻長得一樣。左側色槓與詳情頁的狀態標籤都讀這張表。
const STATUS_COLOR = {
  '確定參加': 'blue', '評估中': 'amber', '已結案': 'green', '暫不參加': 'gray'
};

function fmtMoney(v) {
  if (v === '' || v === undefined || v === null) return '-';
  const n = Number(v);
  if (isNaN(n)) return v;
  return 'NT$ ' + n.toLocaleString();
}

function parseDate(d) {
  if (!d) return null;
  const p = String(d).split('/');
  if (p.length !== 3) return null;
  const dt = new Date(+p[0], +p[1] - 1, +p[2]);
  return isNaN(dt) ? null : dt;
}

function daysInfo(ex) {
  const s = parseDate(ex['展期開始']), e = parseDate(ex['展期結束']) || s;
  if (!s) return { text: '', cls: '' };
  const today = new Date(); today.setHours(0, 0, 0, 0);
  if (e && today > e) return { text: '已結束', cls: 'cnt-past' };
  if (today >= s)     return { text: '進行中', cls: 'cnt-now' };
  const n = Math.round((s - today) / 86400000);
  return { text: '還有 ' + n + ' 天', cls: 'cnt-soon' };
}

function stageSetHtml(ex) {
  const st = ex['參展狀態'];
  if (st !== '確定參加' && st !== '評估中') return '';
  const pill = en => {
    const on = ex[STAGE_FIELDS[en]] === '是';
    return `<span class="stage ${on ? 'stage-on' : 'stage-off'}">${stageLabel(en, ex)}</span>`;
  };
  return `<span class="stage-set ex-right">${Object.keys(STAGE_FIELDS).map(pill).join('')}</span>`;
}

function shortDate(d) {
  if (!d) return '-';
  const parts = d.split('/');
  return parts.length === 3 ? `${parts[1]}/${parts[2]}` : d;
}

// 資料夾一定存在（建清單時一次建好），所以連結是算出來的，不是存出來的。
// 裡面有沒有檔案這支程式看不到——那要另外去讀文件庫，是下一階段的事。
function folderUrl(ex, key) {
  const y = ex['年度'], k = ex['資料夾代號'];
  if (!y || !k) return '';
  return EX_FOLDER_ROOT + '/' + encodeURIComponent(y)
    + '/' + encodeURIComponent(k)
    + '/' + encodeURIComponent(EX_SUBFOLDERS[key]);
}

function folderCard(ex, key, label, icon) {
  const url = folderUrl(ex, key);
  if (!url) return `<p class="empty-note">這一場還沒填資料夾代號，無法組出連結。</p>`;
  return `<a class="file-card" href="${url}" target="_blank" rel="noopener">
    <span>${icon}</span><span>${label}</span><span class="arrow">↗</span>
  </a>`;
}

// 檔案清單。用途是存查與抓取，所以每一筆都能直接點開，並顯示上傳時間與大小。
// 刻意不從檔名推斷內容——訂命名規則讓程式比對，人一沒照規則就會靜默判斷錯。
async function renderFolderFiles(ex, key, elId) {
  const el = document.getElementById(elId);
  if (!el) return;
  if (!ex['資料夾代號'] || !ex['年度']) { el.innerHTML = ''; return; }
  el.innerHTML = `<p class="empty-note">讀取中…</p>`;
  const r = await driveList(['展覽', ex['年度'], ex['資料夾代號'], EX_SUBFOLDERS[key]]);
  if (r.state === 'missing') {
    el.innerHTML = `<p class="empty-note">找不到這個資料夾。可能是資料夾代號與文件庫上的名稱不一致。</p>`;
    return;
  }
  if (r.state === 'error') {
    el.innerHTML = `<p class="empty-note">讀取失敗（${r.code}）。</p>`;
    return;
  }
  if (r.state === 'empty') {
    el.innerHTML = `<p class="empty-note">資料夾是空的，還沒有檔案。</p>`;
    return;
  }
  el.innerHTML = `<p class="empty-note" style="margin:0 0 6px;">共 ${r.items.length} 個檔案</p>`
    + r.items.map(f => `<a class="doc-row" href="${f.webUrl}" target="_blank" rel="noopener">
        <span class="doc-name">${f.name}</span>
        <span class="doc-meta">${fmtStamp(f.lastModifiedDateTime)}　${fmtSize(f.size)}</span>
      </a>`).join('');
}

function renderAllFolders(idx) {
  const ex = exhibitions[idx];
  renderFolderFiles(ex, 'design', 'd-design-files');
  renderFolderFiles(ex, 'photo',  'd-photo-files');
  renderFolderFiles(ex, 'report', 'd-report-files');
}

function byStartDate(a, b) {
  const da = parseDate(a['展期開始']), db = parseDate(b['展期開始']);
  if (!da && !db) return 0;
  if (!da) return 1;
  if (!db) return -1;
  return da - db;
}

// 狀態不在這裡顯示。它已經由分組標題（「確定參加（4）」）與左側色槓表達了兩次，
// 再掛一顆標籤是第三次——同一組四列的字一模一樣，掃過去零資訊量。
function exRowHtml(ex) {
  const idx = exhibitions.indexOf(ex);
  const color = STATUS_COLOR[ex['參展狀態']] || 'gray';
  const di = daysInfo(ex);
  const scope = ex['展覽類型'];
  const badge = scope
    ? `<span class="nature-badge ${scope === '國內' ? 'nb-dom' : 'nb-int'}">${scope}</span>`
    : '';
  // 攤位號緊貼倒數左邊。倒數有 76px 最小寬度並靠右，所以攤位號各列對齊。
  // 「攤位」二字不寫——M2501、H42+H46 這種格式不會被誤認成別的東西。
  const booth = ex['BoothNo'] ? `<span class="ex-booth ex-right">${ex['BoothNo']}</span>` : '';
  return `
    <div class="ex-row c-${color}" onclick="showDetail(${idx})">
      <div class="ex-line1">
        ${badge}
        <span class="ex-name">${ex['展覽名稱'] || ''}</span>
        ${booth}
        <span class="ex-count ${di.cls}"${booth ? '' : ' style="margin-left:auto;"'}>${di.text}</span>
      </div>
      <div class="ex-line2">
        <span class="ex-meta">${shortDate(ex['展期開始'])} - ${shortDate(ex['展期結束'])}${ex['地點城市'] ? '　·　' + ex['地點城市'] : ''}</span>
        ${stageSetHtml(ex)}
      </div>
    </div>`;
}

function renderYearTabs() {
  const years = [...new Set(exhibitions.map(e => e['年度']).filter(Boolean))].sort();
  currentYear = years[years.length - 1] || null;
  const wrap = document.getElementById('year-tabs');
  const allYears = [...new Set([...years, String(Number(currentYear) + 1)])];
  wrap.innerHTML = allYears.map(y => {
    const hasData = years.includes(y);
    return `<button class="year-tab ${y === currentYear ? 'active' : ''}" ${hasData ? `onclick="selectYear('${y}')"` : 'disabled'}>${y}${hasData ? '' : '（尚無資料）'}</button>`;
  }).join('');
  renderYearView();
}

function selectYear(y) {
  currentYear = y;
  document.querySelectorAll('.year-tab').forEach(b => b.classList.toggle('active', b.textContent.startsWith(y)));
  renderYearView();
}

function renderYearView() {
  const list = exhibitions.filter(e => e['年度'] === currentYear);
  document.getElementById('stats-year-label').textContent = `${currentYear} 年度展覽進度`;
  renderStats(list);
  renderGroupedList(list);
}

function nextExhibition(list) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  return list
    .filter(e => e['參展狀態'] === '確定參加')
    .filter(e => { const d = parseDate(e['展期結束']) || parseDate(e['展期開始']); return d && d >= today; })
    .sort(byStartDate)[0] || null;
}

function renderStats(list) {
  document.getElementById('stat-total').textContent = list.length;
  document.getElementById('stat-confirmed').textContent = list.filter(e => e['參展狀態'] === '確定參加').length;
  document.getElementById('stat-evaluating').textContent = list.filter(e => e['參展狀態'] === '評估中').length;
  document.getElementById('stat-done').textContent = list.filter(e => e['參展狀態'] === '已結案').length;

  const n = nextExhibition(list);
  const el = document.getElementById('next-ex');
  if (n) {
    const di = daysInfo(n);
    el.innerHTML = `下一場　<strong style="color:var(--text);">${n['展覽名稱']}</strong>`
      + `　·　${shortDate(n['展期開始'])} 開展　·　`
      + `<span class="${di.cls}" style="font-weight:600;">${di.text}</span>`
      + (n['地點城市'] ? `　·　${n['地點城市']}` : '');
  } else {
    el.textContent = '本年度已無未結束的確定參加場次。';
  }
}

function renderGroupedList(list) {
  list = list.slice().sort(byStartDate);
  const confirmed = list.filter(e => e['參展狀態'] === '確定參加');
  const evaluating = list.filter(e => e['參展狀態'] === '評估中');
  const done = list.filter(e => e['參展狀態'] === '已結案');
  const skipped = list.filter(e => e['參展狀態'] === '暫不參加');

  document.getElementById('group-confirmed').innerHTML = confirmed.length
    ? `<p class="group-label">確定參加（${confirmed.length}）</p>${confirmed.map(exRowHtml).join('')}` : '';
  document.getElementById('group-evaluating').innerHTML = evaluating.length
    ? `<p class="group-label">評估中（${evaluating.length}）</p>${evaluating.map(exRowHtml).join('')}` : '';
  document.getElementById('group-done').innerHTML = done.length
    ? `<p class="group-label">已結案（${done.length}）</p>${done.map(exRowHtml).join('')}` : '';
  document.getElementById('group-skipped').innerHTML = skipped.length
    ? `<p class="collapse-toggle" onclick="toggleSkipped()">▸ 暫不參加（${skipped.length}）— 點擊展開</p>
       <div id="skipped-list" class="collapsed">${skipped.map(exRowHtml).join('')}</div>` : '';
}

function toggleSkipped() {
  document.getElementById('skipped-list').classList.toggle('collapsed');
}

function switchTab(name) {
  document.querySelectorAll('.detail-tab').forEach(t => t.classList.toggle('active', t.dataset.tab === name));
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.toggle('hidden', p.id !== 'tab-' + name));
}

// 前置階段：規則 1（可逆的單一欄位）→ 點了就存，沒有按鈕
function renderStages(idx) {
  const ex = exhibitions[idx];
  const html = Object.keys(STAGE_FIELDS).map(en => {
    const on = ex[STAGE_FIELDS[en]] === '是';
    return `<span class="stage ${on ? 'stage-on' : 'stage-off'}"
      style="font-size:12.5px; padding:6px 12px; cursor:${canWrite ? 'pointer' : 'default'};"
      onclick="toggleStage(${idx}, '${en}')">${stageLabel(en, ex)}</span>`;
  }).join('');
  document.getElementById('d-stages').innerHTML = html
    + (canWrite ? '' : '<span class="empty-note" style="align-self:center; margin-left:6px;">（唯讀模式，權限核准後可直接點按）</span>');
}

// 結案簡報是展後的事，不放在前置階段那一排。同樣是規則 1，點了就存。
function renderClosing(idx) {
  const ex = exhibitions[idx];
  const on = ex['結案簡報完成'] === '是';
  document.getElementById('d-closing').innerHTML =
    `<span class="stage ${on ? 'stage-on' : 'stage-off'}"
       style="font-size:12.5px; padding:6px 12px; cursor:${canWrite ? 'pointer' : 'default'};"
       onclick="toggleStage(${idx}, 'ClosingReportDone')">結案簡報${on ? '已完成' : '未完成'}</span>`
    + (canWrite ? '' : '<span class="empty-note">（唯讀模式）</span>');
}

async function toggleStage(idx, en) {
  if (!canWrite) { toast('目前是唯讀模式，寫入權限尚未核准', true); return; }
  const ex = exhibitions[idx];
  const zh = FIELD_MAP[en];
  const next = ex[zh] !== '是';
  try {
    await spPatch('exhibitions', ex._id, { [en]: next });
    ex[zh] = next ? '是' : '否';
    renderStages(idx);
    renderClosing(idx);
    renderYearView();
    // 儀表板只存在於工作台那一頁，展覽檢視頁沒有它。共用檔不能假設它存在。
    if (typeof renderDashboard === 'function') renderDashboard();
  } catch (e) { console.error(e); }
}

function showDetail(idx) {
  const ex = exhibitions[idx];
  // 每次開啟都清掉資料夾快取重讀。你剛在 SharePoint 丟完檔案回來這一頁，
  // 最重要的就是看到它已經在裡面——三個請求的成本遠低於顯示過期內容。
  for (const k in folderCache) delete folderCache[k];
  switchTab('basic');
  renderStages(idx);
  document.getElementById('d-name').textContent = ex['展覽名稱'] || '';
  document.getElementById('d-status').outerHTML = `<span id="d-status" class="tag tag-${STATUS_COLOR[ex['參展狀態']] || 'gray'}" style="margin-left:auto;">${ex['參展狀態'] || '-'}</span>`;
  document.getElementById('d-date').textContent = `${shortDate(ex['展期開始'])} - ${shortDate(ex['展期結束'])}`;
  document.getElementById('d-place').textContent = `${ex['地點城市'] || ''}`;
  document.getElementById('d-booth').textContent = ex['BoothNo'] || '未提供';
  document.getElementById('d-boothsize').textContent = ex['BoothSize'] || '-';
  document.getElementById('d-scope').textContent = ex['展覽類型'] || '-';
  document.getElementById('d-salesperson').textContent = ex['業務參展人'] || '-';

  // 空白與 0 是兩件事：空白代表由別的單位付或還沒確定，0 代表真的不用錢。
  const feeVal = k => {
    const raw = ex[k];
    if (raw === '' || raw === undefined || raw === null) return null;
    const n = Number(raw);
    return isNaN(n) ? null : n;
  };
  const fees = ['攤位費用', '運費', '裝置費用'].map(feeVal);
  document.getElementById('d-fee-booth').textContent = fmtMoney(ex['攤位費用']);
  document.getElementById('d-fee-freight').textContent = fmtMoney(ex['運費']);
  document.getElementById('d-fee-setup').textContent = fmtMoney(ex['裝置費用']);
  const filled = fees.filter(v => v !== null);
  const blanks = 3 - filled.length;
  document.getElementById('d-fee-total').innerHTML = filled.length
    ? fmtMoney(filled.reduce((a, b) => a + b, 0))
      + (blanks ? ` <span class="empty-note">（${blanks} 欄空白）</span>` : '')
    : '-';

  document.getElementById('d-subsidy-org').textContent = ex['補助單位'] || '（無補助）';
  document.getElementById('d-subsidy-status').textContent = ex['補助核銷狀態'] || '-';
  document.getElementById('d-subsidy-apply').textContent = fmtMoney(ex['補助申請金額']);
  document.getElementById('d-subsidy-approved').textContent = fmtMoney(ex['補助核准金額']);
  document.getElementById('d-note').textContent = ex['備註說明'] || '（無）';

  document.getElementById('d-design').innerHTML = folderCard(ex, 'design', '01_設計產出', '📁');
  document.getElementById('d-photos').innerHTML = folderCard(ex, 'photo', '02_現場照片', '📁');
  document.getElementById('d-report').innerHTML = folderCard(ex, 'report', '03_結案報告', '📁');
  renderClosing(idx);
  renderAllFolders(idx);

  document.getElementById('view-main').classList.add('hidden');
  document.getElementById('view-detail').classList.remove('hidden');
  window.scrollTo(0, 0);
}

function showMain() {
  document.getElementById('view-detail').classList.add('hidden');
  document.getElementById('view-main').classList.remove('hidden');
}
