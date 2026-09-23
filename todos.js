/* 工作台｜待辦事項模組
   只掛在根目錄的工作台，不掛在公開的 /exhibitions/。
   這是部門內部的交辦資料，公開頁刻意不讀它——不是不顯示，是不發請求。

   相依：core.js（spGet／spPatch／spCreate／spSoftDelete／canWrite／toast
                  ／normalizeFields／parseDate／toSlashDate／shortDate）
   掛載：mountTodos(容器元素)

   命名一律加 todo 前綴。exhibitions.js 已經佔用了 showMain／showDetail／switchTab
   這些名字，兩邊同名會互相蓋掉，而且不會報錯，只會行為變得莫名其妙。 */

/* ================================================================
   一、設定
   ================================================================ */
// 左邊是 SharePoint 內部名稱（建立後改不掉），右邊是程式內部用的鍵名。
// 對照的是 2026-09-23 建立的 Todos，八欄（Title 是內建的）。
const TODO_FIELD_MAP = {
  Title: '工作名稱', Detail: '內容說明', Dept: '部門',
  Owner: '負責人', Assigner: '交辦人',
  Status: '狀態', DueDate: '到期日', CompletedDate: '完成日期',
  IsDeleted: '已刪除'
};
const TODO_BOOLS  = ['IsDeleted'];
const TODO_DATES  = ['DueDate', 'CompletedDate'];
const TODO_PEOPLE = ['Owner', 'Assigner'];

// 人員欄位存的是網站使用者的內部編號，不是 email。
// 這三個編號是 2026-09-23 用 /_api/web/siteusers 查出來的實際值，不是推測。
// 之後有新同仁加入，要用同一個方式查出編號再加進來。
const TODO_PEOPLE_LIST = [
  { id: '6',  name: 'Yulun' },
  { id: '10', name: 'Eric' },
  { id: '11', name: 'Karen' }
];
function todoPersonName(id) {
  if (!id) return '';
  const p = TODO_PEOPLE_LIST.find(x => x.id === String(id));
  return p ? p.name : '（未知）';
}

const TODO_STATUSES = ['未開始', '進行中', '暫緩', '已完成'];
const TODO_STATUS_COLOR = { '未開始': 'gray', '進行中': 'blue', '暫緩': 'amber', '已完成': 'green' };
const TODO_DEPTS = ['企劃', '總務'];

let todos = [];
let todoFilter = { owner: '', status: '', dept: '' };

/* ================================================================
   二、讀取
   ================================================================ */
DATA_LOADERS.todos = async function () {
  const raw = await spGet('todos');
  todos = raw
    .map(it => normalizeFields(it.fields || {}, it.id, TODO_FIELD_MAP,
          { bools: TODO_BOOLS, dates: TODO_DATES, people: TODO_PEOPLE }))
    .filter(t => t['工作名稱'] && t['已刪除'] !== '是');
};

/* ================================================================
   三、版面
   ================================================================ */
const TODO_MARKUP = `
  <div id="todo-view-main">

    <div class="card">
      <p class="section-title" style="margin-top:0;">待辦進度</p>
      <div class="stats-grid">
        <div class="stat-box stat-total"><p>未開始</p><p id="todo-stat-new">-</p></div>
        <div class="stat-box stat-confirmed"><p>進行中</p><p id="todo-stat-doing">-</p></div>
        <div class="stat-box stat-evaluating"><p>暫緩</p><p id="todo-stat-hold">-</p></div>
        <div class="stat-box stat-done"><p>已完成</p><p id="todo-stat-done">-</p></div>
      </div>
      <p id="todo-overdue" style="font-size:13px; margin:14px 0 0; padding-top:12px;
         border-top:1px solid var(--border); color:var(--text-secondary);"></p>
    </div>

    <div class="card">
      <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
        <p class="section-title" style="margin:0;">新增待辦</p>
        <button id="todo-new-btn" class="back-btn" style="margin-left:auto;"
                onclick="todoToggleForm()">＋ 新增</button>
      </div>
      <div id="todo-form" class="hidden" style="margin-top:14px;">
        <div class="todo-field"><label>工作名稱</label>
          <input id="tf-title" type="text" placeholder="例：整理秋季型錄封面稿"></div>
        <div class="todo-field"><label>內容說明</label>
          <textarea id="tf-detail" rows="3" placeholder="可留空"></textarea></div>
        <div class="todo-field-row">
          <div class="todo-field"><label>部門</label>
            <select id="tf-dept">${TODO_DEPTS.map(d => `<option value="${d}">${d}</option>`).join('')}</select></div>
          <div class="todo-field"><label>到期日</label>
            <input id="tf-due" type="date"></div>
        </div>
        <div class="todo-field-row">
          <div class="todo-field"><label>負責人</label>
            <select id="tf-owner">${TODO_PEOPLE_LIST.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}</select></div>
          <div class="todo-field"><label>交辦人</label>
            <select id="tf-assigner">${TODO_PEOPLE_LIST.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}</select></div>
        </div>
        <div style="display:flex; gap:8px; align-items:center; margin-top:12px;">
          <button class="back-btn" onclick="todoCreate()">新增</button>
          <button class="back-btn" onclick="todoToggleForm()">取消</button>
          <span id="todo-form-note" class="empty-note"></span>
        </div>
      </div>
    </div>

    <div class="card">
      <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin-bottom:12px;">
        <p class="section-title" style="margin:0;">全部待辦　·　點任一筆查看與修改</p>
      </div>
      <div class="todo-filters">
        <select id="todo-f-owner" onchange="todoSetFilter()">
          <option value="">全部負責人</option>
          ${TODO_PEOPLE_LIST.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}
        </select>
        <select id="todo-f-status" onchange="todoSetFilter()">
          <option value="">全部狀態</option>
          ${TODO_STATUSES.map(s => `<option value="${s}">${s}</option>`).join('')}
        </select>
        <select id="todo-f-dept" onchange="todoSetFilter()">
          <option value="">全部部門</option>
          ${TODO_DEPTS.map(d => `<option value="${d}">${d}</option>`).join('')}
        </select>
        <button class="back-btn" onclick="todoClearFilter()">清除</button>
      </div>
      <div id="todo-list"></div>
    </div>

  </div>

  <div id="todo-view-detail" class="hidden">
    <button class="back-btn" onclick="todoShowMain()">← 返回清單</button>
    <div class="card" style="margin-top:10px;">
      <div style="display:flex; align-items:center; gap:10px; margin-bottom:16px;">
        <span id="td-name" style="font-size:18px; font-weight:700;"></span>
        <span id="td-status-tag" class="tag" style="margin-left:auto;"></span>
      </div>
      <div class="hero-grid">
        <div class="hero-box"><p><span class="hero-icon">🏷️</span>部門</p><p id="td-dept"></p></div>
        <div class="hero-box"><p><span class="hero-icon">📅</span>到期日</p><p id="td-due"></p></div>
        <div class="hero-box"><p><span class="hero-icon">🙋</span>負責人</p><p id="td-owner"></p></div>
        <div class="hero-box"><p><span class="hero-icon">📮</span>交辦人</p><p id="td-assigner"></p></div>
      </div>
    </div>

    <div class="card">
      <p class="section-title" style="margin-top:0;">狀態</p>
      <div id="td-status-pills" style="display:flex; gap:6px; flex-wrap:wrap; margin-bottom:18px;"></div>

      <p class="section-title">指派與改期</p>
      <div class="todo-field-row">
        <div class="todo-field"><label>負責人</label>
          <select id="td-owner-sel" onchange="todoAssign()"></select></div>
        <div class="todo-field"><label>到期日</label>
          <input id="td-due-input" type="date" onchange="todoReschedule()"></div>
      </div>
      <p id="td-write-note" class="empty-note" style="margin:10px 0 0;"></p>

      <p class="section-title">內容說明</p>
      <p id="td-detail" style="font-size:13px; margin:0 0 18px; color:var(--text-secondary); white-space:pre-wrap;"></p>

      <p class="section-title">完成日期</p>
      <p id="td-completed" style="font-size:13px; margin:0 0 18px; color:var(--text-secondary);"></p>

      <div style="padding-top:14px; border-top:1px solid var(--border);">
        <button class="back-btn" onclick="todoDelete()">刪除這一筆</button>
        <span class="empty-note" style="margin-left:8px;">軟刪除，資料還在，可在回收桶還原。</span>
      </div>
    </div>
  </div>
`;

// 這幾個樣式只有待辦在用，放在模組裡比塞進 app.css 好——
// app.css 是兩個頁面共用的，公開頁不需要這些規則。
const TODO_STYLE = `
<style>
.todo-field { margin-bottom:10px; }
.todo-field label { display:block; font-size:12px; color:var(--label); margin-bottom:4px; }
.todo-field input, .todo-field select, .todo-field textarea {
  width:100%; box-sizing:border-box; font-family:inherit; font-size:13px;
  padding:8px 10px; border:1px solid var(--border); border-radius:8px;
  background:var(--card); color:var(--text);
}
.todo-field textarea { resize:vertical; }
.todo-field-row { display:flex; gap:10px; flex-wrap:wrap; }
.todo-field-row .todo-field { flex:1 1 160px; }
.todo-filters { display:flex; gap:8px; flex-wrap:wrap; margin-bottom:14px; }
.todo-filters select {
  font-family:inherit; font-size:12.5px; padding:6px 10px;
  border:1px solid var(--border); border-radius:8px; background:var(--card); color:var(--text);
}
.todo-overdue { color:var(--red); font-weight:600; }
</style>
`;

function mountTodos(el) {
  el.innerHTML = TODO_STYLE + TODO_MARKUP;
}

/* ================================================================
   四、繪製
   ================================================================ */
// 到期日的倒數。沒有到期日就不顯示——硬要顯示「未設定」會讓每一列都多一段雜訊。
function todoDueInfo(t) {
  const d = parseDate(t['到期日']);
  if (!d) return { text: '', cls: '', style: '' };
  if (t['狀態'] === '已完成') return { text: '已完成', cls: 'cnt-past', style: '' };
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const n = Math.round((d - today) / 86400000);
  if (n < 0)  return { text: '逾期 ' + (-n) + ' 天', cls: '', style: 'color:var(--red); font-weight:600;' };
  if (n === 0) return { text: '今天到期', cls: 'cnt-now', style: '' };
  return { text: '還有 ' + n + ' 天', cls: 'cnt-soon', style: '' };
}

function todoFiltered() {
  return todos.filter(t =>
    (!todoFilter.owner  || t['負責人'] === todoFilter.owner) &&
    (!todoFilter.status || t['狀態']   === todoFilter.status) &&
    (!todoFilter.dept   || t['部門']   === todoFilter.dept));
}

// 排序：未完成的在前，其中到期日近的在前，沒有到期日的排最後。
function todoSort(a, b) {
  const doneA = a['狀態'] === '已完成', doneB = b['狀態'] === '已完成';
  if (doneA !== doneB) return doneA ? 1 : -1;
  const da = parseDate(a['到期日']), db = parseDate(b['到期日']);
  if (!da && !db) return 0;
  if (!da) return 1;
  if (!db) return -1;
  return da - db;
}

function todoRowHtml(t) {
  const idx = todos.indexOf(t);
  const color = TODO_STATUS_COLOR[t['狀態']] || 'gray';
  const di = todoDueInfo(t);
  const dept = t['部門']
    ? `<span class="nature-badge ${t['部門'] === '企劃' ? 'nb-dom' : 'nb-int'}">${t['部門']}</span>`
    : '';
  const who = [
    t['負責人']  ? todoPersonName(t['負責人']) : '',
    t['交辦人'] && t['交辦人'] !== t['負責人'] ? '← ' + todoPersonName(t['交辦人']) : ''
  ].filter(Boolean).join('　');
  return `
    <div class="ex-row c-${color}" onclick="todoShowDetail(${idx})">
      <div class="ex-line1">
        ${dept}
        <span class="ex-name">${t['工作名稱']}</span>
        <span class="ex-count ${di.cls} ex-right" style="${di.style}">${di.text}</span>
      </div>
      <div class="ex-line2">
        <span class="ex-meta">${who}${t['到期日'] ? '　·　' + shortDate(t['到期日']) : ''}</span>
        <span class="stage-set ex-right"><span class="tag tag-${color}">${t['狀態'] || '未開始'}</span></span>
      </div>
    </div>`;
}

function todoRenderList() {
  const list = todoFiltered().slice().sort(todoSort);
  const el = document.getElementById('todo-list');
  if (!el) return;
  el.innerHTML = list.length
    ? list.map(todoRowHtml).join('')
    : `<p class="empty-note">${todos.length ? '這組篩選沒有符合的項目。' : '還沒有待辦事項。按上面的「新增」建立第一筆。'}</p>`;
}

function todoRenderStats() {
  const c = s => todos.filter(t => t['狀態'] === s).length;
  document.getElementById('todo-stat-new').textContent   = c('未開始');
  document.getElementById('todo-stat-doing').textContent = c('進行中');
  document.getElementById('todo-stat-hold').textContent  = c('暫緩');
  document.getElementById('todo-stat-done').textContent  = c('已完成');

  const overdue = todoOverdue();
  const el = document.getElementById('todo-overdue');
  if (overdue.length) {
    el.innerHTML = `<span class="todo-overdue">逾期 ${overdue.length} 件</span>　·　`
      + overdue.slice(0, 3).map(t => t['工作名稱']).join('、')
      + (overdue.length > 3 ? ' 等' : '');
  } else {
    el.textContent = todos.length ? '沒有逾期的項目。' : '';
  }
}

function todoOverdue() {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  return todos.filter(t => {
    if (t['狀態'] === '已完成') return false;
    const d = parseDate(t['到期日']);
    return d && d < today;
  }).sort(todoSort);
}

function todoRender() {
  todoRenderStats();
  todoRenderList();
  const btn = document.getElementById('todo-new-btn');
  if (btn) {
    btn.disabled = !canWrite;
    btn.style.opacity = canWrite ? '' : '0.45';
    btn.style.cursor  = canWrite ? 'pointer' : 'default';
    btn.title = canWrite ? '' : '寫入權限核准後可用';
  }
}

function todoSetFilter() {
  todoFilter.owner  = document.getElementById('todo-f-owner').value;
  todoFilter.status = document.getElementById('todo-f-status').value;
  todoFilter.dept   = document.getElementById('todo-f-dept').value;
  todoRenderList();
}

function todoClearFilter() {
  todoFilter = { owner: '', status: '', dept: '' };
  ['todo-f-owner', 'todo-f-status', 'todo-f-dept'].forEach(id => document.getElementById(id).value = '');
  todoRenderList();
}

/* ================================================================
   五、詳情與寫入
   ================================================================ */
let todoCurrent = -1;

function todoShowMain() {
  document.getElementById('todo-view-detail').classList.add('hidden');
  document.getElementById('todo-view-main').classList.remove('hidden');
}

function todoShowDetail(idx) {
  todoCurrent = idx;
  const t = todos[idx];
  const color = TODO_STATUS_COLOR[t['狀態']] || 'gray';

  document.getElementById('td-name').textContent = t['工作名稱'];
  document.getElementById('td-status-tag').outerHTML =
    `<span id="td-status-tag" class="tag tag-${color}" style="margin-left:auto;">${t['狀態'] || '未開始'}</span>`;
  document.getElementById('td-dept').textContent     = t['部門'] || '-';
  document.getElementById('td-due').textContent      = t['到期日'] || '未設定';
  document.getElementById('td-owner').textContent    = todoPersonName(t['負責人']) || '-';
  document.getElementById('td-assigner').textContent = todoPersonName(t['交辦人']) || '-';
  document.getElementById('td-detail').textContent   = t['內容說明'] || '（無）';
  document.getElementById('td-completed').textContent = t['完成日期'] || '（尚未完成）';

  todoRenderStatusPills();

  const sel = document.getElementById('td-owner-sel');
  sel.innerHTML = '<option value="">（未指派）</option>'
    + TODO_PEOPLE_LIST.map(p => `<option value="${p.id}">${p.name}</option>`).join('');
  sel.value = t['負責人'] || '';
  sel.disabled = !canWrite;

  const due = document.getElementById('td-due-input');
  due.value = t['到期日'] ? t['到期日'].replace(/\//g, '-') : '';
  due.disabled = !canWrite;

  document.getElementById('td-write-note').textContent =
    canWrite ? '' : '目前是唯讀模式，寫入權限核准後即可直接修改。';

  document.getElementById('todo-view-main').classList.add('hidden');
  document.getElementById('todo-view-detail').classList.remove('hidden');
  window.scrollTo(0, 0);
}

// 狀態：可逆的單一欄位 → 規則 1，點了就存，沒有按鈕。
// 完成日期由程式順手寫，不要求使用者填——要打字的欄位三週後一定是空的。
function todoRenderStatusPills() {
  const t = todos[todoCurrent];
  document.getElementById('td-status-pills').innerHTML =
    TODO_STATUSES.map(s => {
      const on = (t['狀態'] || '未開始') === s;
      return `<span class="stage ${on ? 'stage-on' : 'stage-off'}"
        style="font-size:12.5px; padding:6px 12px; cursor:${canWrite ? 'pointer' : 'default'};"
        onclick="todoSetStatus('${s}')">${s}</span>`;
    }).join('')
    + (canWrite ? '' : '<span class="empty-note" style="align-self:center; margin-left:6px;">（唯讀模式）</span>');
}

function todoTodayIso() {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + 'T00:00:00Z';
}

async function todoSetStatus(s) {
  if (!canWrite) { toast('目前是唯讀模式，寫入權限尚未核准', true); return; }
  const t = todos[todoCurrent];
  if ((t['狀態'] || '未開始') === s) return;
  const fields = { Status: s };
  // 改成已完成就記下今天；從已完成改回去就清掉，否則會留下一個假的完成日期。
  if (s === '已完成')            fields.CompletedDate = todoTodayIso();
  else if (t['狀態'] === '已完成') fields.CompletedDate = null;
  try {
    await spPatch('todos', t._id, fields);
    t['狀態'] = s;
    if (s === '已完成') t['完成日期'] = toSlashDate(fields.CompletedDate);
    else if (fields.CompletedDate === null) t['完成日期'] = '';
    todoShowDetail(todoCurrent);
    todoRender();
    if (typeof renderDashboard === 'function') renderDashboard();
  } catch (e) { console.error(e); }
}

async function todoAssign() {
  if (!canWrite) { toast('目前是唯讀模式，寫入權限尚未核准', true); return; }
  const t = todos[todoCurrent];
  const v = document.getElementById('td-owner-sel').value;
  try {
    await spPatch('todos', t._id, { OwnerLookupId: v === '' ? null : v });
    t['負責人'] = v;
    document.getElementById('td-owner').textContent = todoPersonName(v) || '-';
    todoRender();
  } catch (e) { console.error(e); }
}

async function todoReschedule() {
  if (!canWrite) { toast('目前是唯讀模式，寫入權限尚未核准', true); return; }
  const t = todos[todoCurrent];
  const v = document.getElementById('td-due-input').value;   // yyyy-mm-dd
  try {
    await spPatch('todos', t._id, { DueDate: v ? v + 'T00:00:00Z' : null });
    t['到期日'] = v ? v.replace(/-/g, '/') : '';
    document.getElementById('td-due').textContent = t['到期日'] || '未設定';
    todoRender();
    if (typeof renderDashboard === 'function') renderDashboard();
  } catch (e) { console.error(e); }
}

// 刪除：規則 3（不可逆感受）→ 一定要按，而且先問一次。
// 實際上是軟刪除，改 IsDeleted，資料還在。
async function todoDelete() {
  if (!canWrite) { toast('目前是唯讀模式，寫入權限尚未核准', true); return; }
  const t = todos[todoCurrent];
  if (!confirm('確定刪除「' + t['工作名稱'] + '」？\n這是軟刪除，資料還在，可以還原。')) return;
  try {
    await spSoftDelete('todos', t._id);
    todos.splice(todoCurrent, 1);
    todoShowMain();
    todoRender();
    if (typeof renderDashboard === 'function') renderDashboard();
  } catch (e) { console.error(e); }
}

function todoToggleForm() {
  if (!canWrite) { toast('目前是唯讀模式，寫入權限尚未核准', true); return; }
  document.getElementById('todo-form').classList.toggle('hidden');
  document.getElementById('todo-form-note').textContent = '';
}

async function todoCreate() {
  if (!canWrite) { toast('目前是唯讀模式，寫入權限尚未核准', true); return; }
  const title = document.getElementById('tf-title').value.trim();
  if (!title) {
    document.getElementById('todo-form-note').textContent = '工作名稱不能空白。';
    return;
  }
  const due = document.getElementById('tf-due').value;
  const fields = {
    Title:  title,
    Detail: document.getElementById('tf-detail').value.trim(),
    Dept:   document.getElementById('tf-dept').value,
    Status: '未開始',
    OwnerLookupId:    document.getElementById('tf-owner').value,
    AssignerLookupId: document.getElementById('tf-assigner').value,
    IsDeleted: false
  };
  if (due) fields.DueDate = due + 'T00:00:00Z';
  try {
    await spCreate('todos', fields);
    // 重讀而不是自己拼一筆假的：新增之後清單裡那一筆的 id 只有伺服器知道，
    // 拼假的會讓接下來的修改打到不存在的項目。
    await DATA_LOADERS.todos();
    ['tf-title', 'tf-detail', 'tf-due'].forEach(id => document.getElementById(id).value = '');
    document.getElementById('todo-form').classList.add('hidden');
    todoRender();
    if (typeof renderDashboard === 'function') renderDashboard();
  } catch (e) {
    document.getElementById('todo-form-note').textContent = '新增失敗，詳細訊息在主控台。';
    console.error(e);
  }
}
