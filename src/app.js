import "./style.css";
import { createIcons, ArrowUpRight, ChevronLeft, ChevronRight, CircleAlert, FileCheck2, FileSpreadsheet, FileStack, LoaderCircle, Search, X } from "lucide";
import { displayName, levels } from "./engine.js";
import { resultTable, selectRows } from "./ledger.js";
import { esc } from "./markup.js";
import { revisionDetails, revisionPanel } from "./revision-view.js";
import sampleURL from "../outputs/01a1e9d3-62a5-7b33-a92a-ac33e674e531/证书汇总输入样例.xlsx?url";

const iconSet = { ArrowUpRight, ChevronLeft, ChevronRight, CircleAlert, FileCheck2, FileSpreadsheet, FileStack, LoaderCircle, Search, X };
const paint = () =>
  createIcons({
    icons: iconSet,
    attrs: { "aria-hidden": "true", "stroke-width": "1.75" },
  });
const i = (name) => '<i data-lucide="' + name + '"></i>';
const main = document.querySelector("#main"),
  dialog = document.querySelector("#dialog"),
  input = document.querySelector("#file");
const titles = {
  summary: "证书汇总", revision: "核对新版工作簿",
  ledger: "数据台账", issue: "发放管理", depart: "部门导出",
};
const subtitles = {
  summary: "合并历史证书和考试成绩，将结果写入 00 表。",
  revision: "先核对来源记录与证书的变化，再下载新版。",
  ledger: "查看当前工作簿的证书、考试成绩与认证规则。",
  issue: "按员工工号批量标记证书的发放状态。",
  depart: "按员工最新所属部门，导出需要的工作表。",
};
const labels = ["汇总认证证书", "历史证书", "考试成绩", "认证规则"];
const ledger = { code: "00", tables: new Map(), views: new Map() };
const revision = {
  previous: null, latest: null, report: null, scope: false, issued: false, statusSource: "",
};
const revisionInputs = {
  previous: document.querySelector("#revision-previous"),
  latest: document.querySelector("#revision-latest"),
};
let revisionDetail;

const state = {
  tab: "summary",
  file: "",
  meta: null,
  result: null,
  busy: "",
  error: "",
  dirty: false,
  query: "",
  level: "",
  page: 0,
  ids: "",
  status: "已发放",
  dept: "",
  codes: new Set(["00", "01", "02"]),
};
const worker = new Worker(new URL("./worker.js", import.meta.url), {
  type: "module",
});
let pending, toastTimer, confirmation;
worker.onmessage = ({ data }) => {
  const p = pending;
  pending = null;
  if (p) data.error ? p.reject(new Error(data.error)) : p.resolve(data.value);
};
worker.onerror = () => {
  pending?.reject(new Error("工作簿处理模块启动失败，请刷新后重试。"));
  pending = null;
};
function rpc(action, args = {}, transfers = []) {
  return new Promise((resolve, reject) => {
    pending = { resolve, reject };
    worker.postMessage({ action, ...args }, transfers);
  });
}
const localStamp = () => {
  const d = new Date(),
    pad = (n) => String(n).padStart(2, "0");
  return (
    d.getFullYear() +
    "-" +
    pad(d.getMonth() + 1) +
    "-" +
    pad(d.getDate()) +
    " " +
    [d.getHours(), d.getMinutes(), d.getSeconds()].map(pad).join(":")
  );
};
function toast(msg) {
  const box = document.querySelector("#notice");
  box.textContent = msg;
  box.classList.add("visible");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => box.classList.remove("visible"), 4500);
}
function openDialog(title, body, className = "") {
  dialog.className = className;
  dialog.innerHTML =
    '<div class="dialog-head"><h2 id="dialog-title" tabindex="-1">' +
    esc(title) +
    '</h2><button class="icon-btn" aria-label="关闭" data-act="close">' +
    i("x") +
    "</button></div>" +
    body;
  paint();
  if (!dialog.open) dialog.showModal();
}
function button(action, label, icon, kind = "secondary", disabled = false) {
  return (
    '<button class="btn ' +
    kind +
    '" data-act="' +
    action +
    '" ' +
    (disabled || state.busy ? "disabled" : "") +
    ">" +
    (["tiny", "sample-link"].includes(kind) ? i(icon) : "") +
    esc(label) +
    "</button>"
  );
}
function pageHeader() {
  let tools = "";
  if (state.tab === "summary") {
    const sample = state.meta ? button("sample", "试用样例", "arrow-up-right", "sample-link") : "";
    tools = `<div class="page-actions">${sample}${button("to-revision", "换版核对", "")}</div>`;
  } else if (state.tab === "revision") {
    tools = button("to-summary", "返回证书汇总", "", "secondary");
  }
  return (
    '<div class="page-head"><div><h1>' +
    titles[state.tab] +
    "</h1><p>" +
    esc(state.tab === "ledger" && state.file ? state.file : subtitles[state.tab]) +
    "</p></div>" +
    tools +
    "</div>"
  );
}
function importPanel() {
  const loaded = Boolean(state.meta);
  const name = loaded ? esc(state.file) : "导入工作簿";
  const lead = loaded
    ? "工作簿已就绪，可以开始汇总。"
    : "拖入客户提供的 .xlsx 文件，或从电脑选择。";
  const action = loaded
    ? button("run", "开始汇总", "play", "primary") +
      button("choose", "更换文件", "folder-open", "plain")
    : button("choose", "选择工作簿", "folder-open", "primary") +
      button("sample", "试用样例", "arrow-up-right", "sample-link");
  return (
    '<section class="import">' +
    '<div class="file-mark">' + i("file-spreadsheet") + "</div>" +
    '<div class="import-copy"><h2>' +
    name +
    "</h2><p>" +
    lead +
    '</p><div class="file-actions">' +
    action +
    "</div></div></section>"
  );
}
function sheetPanel() {
  return (
    '<section class="sheets"><h2>工作簿包含</h2><ul>' +
    [1, 2, 3, 0]
      .map((j) => {
        const tag = state.meta
          ? state.meta.counts[j] + " 条"
          : ["输出", "来源", "来源", "规则"][j];
        return (
          '<li><span class="sheet-id">0' +
          j +
          "</span><b>" +
          labels[j] +
          "</b><small>" +
          tag +
          "</small></li>"
        );
      })
      .join("") +
    "</ul></section>"
  );
}
function currentFile() {
  return (
    '<section class="panel current-file" id="drop"><span class="current-icon">' +
    i("file-check-2") +
    '</span><div class="current-name"><b>' +
    esc(state.file) +
    "</b><small>历史证书 " +
    state.meta.counts[1] +
    " 条 · 考试成绩 " +
    state.meta.counts[2] +
    " 条 · 认证规则 " +
    state.meta.counts[3] +
    ' 条</small></div><div class="file-actions">' +
    button("to-ledger", "查看数据台账", "") +
    button("run", "重新汇总", "rotate-ccw", "plain") +
    button("choose", "更换文件", "folder-open", "plain") +
    "</div></section>"
  );
}
const filtered = () =>
  (state.result?.records ?? []).filter(
    (r) =>
      (!state.level || levels[r.rank] === state.level) &&
      (!state.query ||
        [r.name, r.id, r.lang, r.dept, r.team, displayName(r)].some((v) =>
          v.toLowerCase().includes(state.query.trim().toLowerCase()),
        )),
  );
const pageSize = 20;
function ledgerView() {
  if (!ledger.views.has(ledger.code))
    ledger.views.set(ledger.code, { query: "", filters: new Map(), sort: null, page: 0 });
  return ledger.views.get(ledger.code);
}
function ledgerHeader(name, column) {
  const sort = ledgerView().sort;
  const direction = sort?.column === column ? sort.direction : 0;
  const ariaSort = direction ? `aria-sort="${direction === 1 ? "ascending" : "descending"}"` : "";
  const arrow = direction === 1 ? "↑" : direction === -1 ? "↓" : "↕";
  const label = name.split(/\s*\/\s*(?=[A-Za-z])/)[0];
  return `<th scope="col" ${ariaSort}>
    <button class="ledger-sort" data-ledger-sort="${column}"
      title="${esc(name)}" aria-label="按${esc(name)}排序">
      <span>${esc(label)}</span><span class="sort-arrow" aria-hidden="true">${arrow}</span>
    </button></th>`;
}
function ledgerColumnWidth(name) {
  if (/^(证书$|科目)|证书名称|考试名称/.test(name)) return 240;
  if (/时间|日期/.test(name)) return 148;
  if (/部门/.test(name)) return 145;
  if (/姓名|工号/.test(name)) return 98;
  return 130;
}
function ledgerPanel() {
  if (!state.meta)
    return '<section class="panel empty"><div class="empty-icon">' +
      i("file-spreadsheet") +
      "</div><h2>还没有工作簿</h2><p>导入工作簿后，即可逐表查看、筛选和排序。</p>" +
      button("to-summary", "前往导入", "", "primary") + "</section>";
  const sheets = labels.map((label, index) => {
    const code = "0" + index;
    return `<button data-ledger-sheet="${code}" aria-pressed="${ledger.code === code}"
      ${state.busy ? "disabled" : ""}>${code} ${label}</button>`;
  }).join("");
  const table = ledger.tables.get(ledger.code);
  const sheetTabs = `<div class="ledger-sheets" role="group" aria-label="工作表">${sheets}</div>`;
  if (!table)
    return `<section class="ledger-panel">${sheetTabs}
      ${state.busy ? "" : button("ledger-retry", "重新读取", "")}</section>`;
  const view = ledgerView();
  const filters = table.columns.map((label, index) =>
    `<td><input id="ledger-filter-${index}" data-ledger-filter="${index}" type="search"
      autocomplete="off" placeholder="筛选" aria-label="筛选 ${esc(label)}"
      value="${esc(view.filters.get(index))}"></td>`,
  ).join("");
  const source = ledger.code === "00" && state.result ? "本次汇总" : "导入原表";
  const widths = [70, ...table.columns.map(ledgerColumnWidth)];
  const tableWidth = widths.reduce((total, width) => total + width, 0);
  return `<section class="ledger-panel">${sheetTabs}
    <div class="ledger-tools">
      <label class="search">${i("search")}<input type="search" id="ledger-search"
        aria-label="搜索当前表" placeholder="搜索当前表" value="${esc(view.query)}"></label>
      <button class="btn plain" id="ledger-clear" data-act="ledger-clear">清除筛选</button>
      <span class="ledger-source">${source} · ${table.columns.length} 列</span>
    </div>
    <div class="table-wrap ledger-scroll" role="region" tabindex="0" aria-label="工作表数据，可横向滚动">
      <table class="ledger-table" style="width: ${tableWidth}px">
        <colgroup>${widths.map((width) => `<col style="width: ${width}px">`).join("")}</colgroup>
        <thead>
          <tr id="ledger-head">${ledgerHeader("行号", -1)}${table.columns.map(ledgerHeader).join("")}</tr>
          <tr class="ledger-filters"><td></td>${filters}</tr>
        </thead>
        <tbody id="ledger-rows"></tbody>
      </table>
    </div>
    <p id="ledger-empty" class="empty-text" hidden></p>
    <div class="pager ledger-pager">
      <span id="ledger-count" aria-live="polite"></span>
      <div>
        ${button("ledger-prev", "上一页", "chevron-left", "tiny")}
        <span id="ledger-page"></span>
        ${button("ledger-next", "下一页", "chevron-right", "tiny")}
      </div>
    </div>
  </section>`;
}
function updateLedgerRows() {
  const body = document.querySelector("#ledger-rows");
  if (!body) return;
  const table = ledger.tables.get(ledger.code);
  const view = ledgerView();
  const rows = selectRows(table.rows, view);
  const pages = Math.max(1, Math.ceil(rows.length / pageSize));
  view.page = Math.min(view.page, pages - 1);
  body.innerHTML = rows.slice(view.page * pageSize, (view.page + 1) * pageSize)
    .map(({ line, values }) => `<tr><th scope="row" class="mono">${line}</th>${values
      .map((value) => `<td>${esc(value)}</td>`).join("")}</tr>`).join("");
  const empty = document.querySelector("#ledger-empty");
  empty.hidden = rows.length !== 0;
  empty.textContent = table.rows.length
    ? "没有符合筛选条件的记录。"
    : ledger.code === "00" && !state.result
      ? "原 00 表为空。完成汇总后可在这里查看生成的证书。"
      : "当前表没有记录。";
  document.querySelector("#ledger-count").textContent = rows.length === table.rows.length
    ? `共 ${rows.length} 条记录`
    : `匹配 ${rows.length} 条，共 ${table.rows.length} 条记录`;
  document.querySelector("#ledger-page").textContent = `${view.page + 1} / ${pages}`;
  document.querySelector('[data-act="ledger-prev"]').disabled = view.page === 0;
  document.querySelector('[data-act="ledger-next"]').disabled = view.page + 1 >= pages;
  document.querySelector("#ledger-clear").disabled = !view.query && view.filters.size === 0;
}
async function openLedger(code = ledger.code) {
  if (state.busy) return;
  state.tab = "ledger";
  ledger.code = code;
  if (state.meta && !ledger.tables.has(code)) {
    await task("正在读取工作表…", async () => {
      ledger.tables.set(code, await rpc("table", { code }));
    });
  } else {
    render();
  }
}
function orderLedger(button) {
  const column = Number(button.dataset.ledgerSort);
  const view = ledgerView();
  const direction = view.sort?.column === column ? -view.sort.direction : 1;
  view.sort = { column, direction };
  view.page = 0;
  const table = ledger.tables.get(ledger.code);
  document.querySelector("#ledger-head").innerHTML =
    ledgerHeader("行号", -1) + table.columns.map(ledgerHeader).join("");
  document.querySelector(`[data-ledger-sort="${column}"]`).focus({ preventScroll: true });
  updateLedgerRows();
}
const tableHead =
  '<thead><tr><th>员工</th><th>认证级别</th><th>语言类别</th><th>获得日期</th><th>所属部门</th><th>发放状态</th><th><span class="sr">详情</span></th></tr></thead>';
function tableContent() {
  const rows = filtered(),
    totalPages = Math.max(1, Math.ceil(rows.length / pageSize));
  state.page = Math.min(state.page, totalPages - 1);
  const slice = rows.slice(state.page * pageSize, (state.page + 1) * pageSize);
  let table = '<div class="table-wrap"><table>' + tableHead + "<tbody>";
  table +=
    slice
      .map(
        (r) =>
          "<tr><td><b>" +
          esc(r.name) +
          '</b><small class="mono">' +
          esc(r.id) +
          '</small></td><td><span class="level level' +
          r.rank +
          '">' +
          levels[r.rank] +
          "</span></td><td>" +
          esc(r.lang) +
          '</td><td class="mono">' +
          esc(r.date) +
          "</td><td>" +
          esc(r.team || r.dept || "—") +
          "</td><td>" +
          esc(r.status || "待标记") +
          '</td><td><button class="detail-link" aria-label="查看 ' +
          esc(r.name) +
          " " +
          esc(r.lang) +
          " " +
          levels[r.rank] +
          ' 详情" data-detail="' +
          rowIndex.get(r) +
          '">详情</button></td></tr>',
      )
      .join("") + "</tbody></table></div>";
  if (!slice.length)
    table +=
      '<p class="empty-text">' +
      (state.query || state.level
        ? "没有匹配的证书，试试其他工号或级别。"
        : "本次汇总没有符合条件的证书。") +
      "</p>";
  return (
    table +
    '<div class="pager"><span>共 ' +
    rows.length +
    " 条证书</span><div>" +
    button("prev", "上一页", "chevron-left", "tiny", !state.page) +
    "<span>" +
    (state.page + 1) +
    " / " +
    totalPages +
    "</span>" +
    button(
      "next",
      "下一页",
      "chevron-right",
      "tiny",
      state.page + 1 >= totalPages,
    ) +
    "</div></div>"
  );
}
let rowIndex = new Map();
const changeLabels = { added: "新增", removed: "移出", dateChanged: "日期调整" };
let changePage = 0;
function changeSummary(changes) {
  if (!changes) return "";
  if (!changes.items.length)
    return '<p class="change-summary unchanged">对照原 00 表，证书无新增或移出，获得日期无变化。</p>';
  const totals = Object.entries(changeLabels)
    .map(([key, label]) => label + " " + changes.counts[key])
    .join(" · ");
  const issued = changes.issuedCount ? "，其中 " + changes.issuedCount + " 张原已发放" : "";
  return '<div class="change-summary"><div><b>与原 00 表对照</b><span>' + totals + issued +
    '</span></div>' + button("changes", "查看变更", "arrow-up-right", "sample-link") + "</div>";
}
function changeItem(item) {
  let dates;
  if (item.kind === "dateChanged") dates = item.beforeDate + " → " + item.afterDate;
  else if (item.kind === "removed") dates = "原获得日期 " + item.beforeDate;
  else dates = "本次获得日期 " + item.afterDate;
  const location = item.sourceRow
    ? "<span>原 00 表第 " + esc(item.sourceRow) + " 行</span>"
    : "";
  return `<li><div><b>${esc([item.name, item.id].filter(Boolean).join("  ·  "))}</b>
    <span>${esc(item.lang)}  ·  ${levels[item.rank]}${item.wasIssued ? "  ·  原已发放" : ""}</span>
    <small><span>${esc(dates)}</span>${location}</small></div><strong>${changeLabels[item.kind]}</strong></li>`;
}
function updateChangeList() {
  const { items, issuedCount } = state.result.changes;
  const count = dialog.querySelector("#issued-only").checked ? issuedCount : items.length;
  // The worker sorts previously issued certificates first.
  const pages = Math.ceil(count / pageSize);
  const first = changePage * pageSize;
  const slice = items.slice(first, Math.min(first + pageSize, count));
  dialog.querySelector("#change-count").textContent = count + " 条变更";
  dialog.querySelector(".change-list").innerHTML = slice.map(changeItem).join("");
  const pager = dialog.querySelector(".pager");
  pager.hidden = pages < 2;
  pager.querySelector("span").textContent = (changePage + 1) + " / " + pages;
  pager.querySelector('[data-act="change-prev"]').disabled = changePage === 0;
  pager.querySelector('[data-act="change-next"]').disabled = changePage + 1 >= pages;
}
function showChanges() {
  const { issuedCount } = state.result.changes;
  changePage = 0;
  openDialog("重算变更",
    `<p class="change-intro">对照导入时的 00 表，原已发放的证书优先显示。</p>
    <div class="change-tools"><label><input id="issued-only" type="checkbox" ${issuedCount ? "" : "disabled"}>
    只看原已发放（${issuedCount}）</label><small id="change-count" aria-live="polite"></small></div>
    <ul class="change-list"></ul><div class="pager" hidden>
    ${button("change-prev", "上一页", "chevron-left", "tiny")}<span aria-live="polite"></span>
    ${button("change-next", "下一页", "chevron-right", "tiny")}</div>`, "changes-dialog");
  updateChangeList();
}
function turnChangePage(step) {
  changePage += step;
  updateChangeList();
  dialog.scrollTop = 0;
  dialog.querySelector("#dialog-title").focus({ preventScroll: true });
}
function clearRevisionReview() {
  Object.assign(revision, { report: null, scope: false, issued: false, statusSource: "" });
}
function updateRevisionDetails() {
  const { data, filter, page } = revisionDetail;
  const items = filter ? data.items.filter((item) => item.kind === filter) : data.items;
  const pages = Math.max(1, Math.ceil(items.length / pageSize));
  const start = page * pageSize;
  dialog.querySelector(".revision-items").innerHTML = items.slice(start, start + pageSize).map(data.render).join("");
  dialog.querySelector("#revision-detail-count").textContent = `共 ${items.length} 项 · ${page + 1} / ${pages}`;
  dialog.querySelector('[data-act="revision-detail-prev"]').disabled = page === 0;
  dialog.querySelector('[data-act="revision-detail-next"]').disabled = page + 1 >= pages;
}
function showRevisionDetails(section) {
  const data = revisionDetails(revision.report, section);
  revisionDetail = { data, filter: "", page: 0 };
  const kinds = data.labels && new Set(data.items.map((item) => item.kind));
  const filter = kinds?.size > 1
    ? `<label class="revision-detail-filter">变更类型
        <select id="revision-detail-filter"><option value="">全部</option>
          ${[...kinds].map((kind) => `<option value="${kind}">${data.labels[kind]}</option>`).join("")}
        </select></label>`
    : "";
  openDialog(data.title, `${filter}<ul class="revision-items"></ul>
    <div class="pager"><span id="revision-detail-count" aria-live="polite"></span><div>
    ${button("revision-detail-prev", "上一页", "chevron-left", "tiny")}
    ${button("revision-detail-next", "下一页", "chevron-right", "tiny")}</div></div>`,
  "revision-dialog");
  updateRevisionDetails();
}
function turnRevisionPage(step) {
  revisionDetail.page += step;
  updateRevisionDetails();
  dialog.scrollTop = 0;
  dialog.querySelector("#dialog-title").focus({ preventScroll: true });
}
function resultPanel() {
  const { records, changes, unmatched } = state.result;
  const employeeCount = new Set(records.map((record) => record.id)).size;
  const issuedCount = records.filter((record) => record.status === "已发放").length;
  const review = state.tab === "summary" ? changeSummary(changes) : "";
  const warning = unmatched
    ? `<p class="warning">${i("circle-alert")}有 ${unmatched} 条通过记录未匹配 03 中的考试名称。</p>`
    : "";
  const levelOptions = levels
    .map((level) => `<option ${state.level === level ? "selected" : ""}>${level}</option>`)
    .join("");
  return `<section class="results">
    <div class="panel-head">
      <div class="result-heading"><h2>汇总结果</h2><span class="sheet-chip">00 表</span></div>
      ${button("download", "下载工作簿", "download", "primary")}
    </div>
    ${review}${warning}
    <div class="filters">
      <label class="search">${i("search")}
        <input id="search" aria-label="搜索证书" placeholder="搜索工号、姓名或语言" value="${esc(state.query)}">
      </label>
      <select id="level" aria-label="筛选证书级别"><option value="">全部级别</option>${levelOptions}</select>
      <div class="stats"><span>${records.length} 张证书</span><span>${employeeCount} 位员工</span><span>${issuedCount} 张已发放</span></div>
    </div>
    <div id="table">${tableContent()}</div>
  </section>`;
}
function prerequisite() {
  return (
    '<section class="panel empty"><div class="empty-icon">' +
    i("file-stack") +
    "</div><strong>先生成证书清单</strong><p>导入工作簿并完成汇总后，即可" +
    (state.tab === "issue" ? "批量标记发放状态。" : "选择部门和工作表。") +
    "</p>" +
    button("to-summary", "前往证书汇总", "arrow-right", "primary") +
    "</section>"
  );
}
function matchIds() {
  const ids = new Set(
    state.ids
      .trim()
      .split(/[\s,，;；]+/)
      .filter(Boolean),
  );
  const matched = state.result.records.filter((r) => ids.has(r.id)),
    employees = new Set(matched.map((r) => r.id));
  return {
    matched,
    ids: [...employees],
    missing: [...ids].filter((id) => !employees.has(id)),
  };
}
function matchPreview() {
  const { matched, ids, missing } = matchIds();
  return (
    '<p class="match">已匹配 <b>' +
    ids.length +
    "</b> 位员工，共 <b>" +
    matched.length +
    "</b> 张证书</p>" +
    (missing.length
      ? '<p class="warning">未匹配工号（' +
        missing.length +
        "）：" +
        esc(missing.slice(0, 20).join("、")) +
        (missing.length > 20 ? "…" : "") +
        "</p>"
      : "") +
    button(
      "mark",
      "标记为" + state.status,
      "check-check",
      "primary",
      !matched.length,
    )
  );
}
function issuePanel() {
  if (!state.result) return prerequisite();
  return (
    '<section class="panel form-panel">' +
    "<h2>批量标记</h2>" +
    '<div class="form-grid"><label>员工工号<textarea id="ids" placeholder="每行输入一个工号，保留前导零" rows="5">' +
    esc(state.ids) +
    '</textarea></label><div><label>发放状态<select id="status">' +
    ["已发放", "未发放"]
      .map(
        (s) =>
          "<option " +
          (state.status === s ? "selected" : "") +
          ">" +
          s +
          "</option>",
      )
      .join("") +
    '</select></label><div id="match" aria-live="polite">' +
    matchPreview() +
    "</div></div></div></section>" +
    resultPanel()
  );
}
function departPanel() {
  if (!state.result) return prerequisite();
  const groups = state.result.departments,
    group = groups.find((d) => d.key === state.dept);
  return (
    '<section class="panel form-panel">' +
    "<h2>导出范围</h2>" +
    '<div class="form-grid"><div><label>所属部门<select id="dept"><option value="">请选择部门</option>' +
    groups
      .map(
        (d) =>
          '<option value="' +
          esc(d.key) +
          '" ' +
          (state.dept === d.key ? "selected" : "") +
          ">" +
          esc(d.label) +
          "</option>",
      )
      .join("") +
    '</select></label><p class="hint">按员工最新部门归组，一层部门为空时按 0 层部门导出。</p></div><fieldset><legend>导出的工作表</legend>' +
    labels
      .slice(0, 3)
      .map(
        (l, j) =>
          '<label class="check-row"><input type="checkbox" value="0' +
          j +
          '" name="sheet" ' +
          (state.codes.has("0" + j) ? "checked" : "") +
          "><span>0" +
          j +
          " · " +
          l +
          "</span><small>" +
          (group ? group.counts[j] + " 条" : "—") +
          "</small></label>",
      )
      .join("") +
    '<div class="form-foot">' +
    button(
      "dept-export",
      "下载部门工作簿",
      "download",
      "primary",
      !group || !state.codes.size,
    ) +
    "</div></fieldset></div></section>"
  );
}
function render() {
  const focused = document.activeElement;
  const focusId = main.contains(focused) ? focused.id : "";
  const focusSheet = focused?.name === "sheet" ? focused.value : "";
  document.title = titles[state.tab] + " · 证书管理台";

  document.querySelectorAll("[data-tab]").forEach((b) => {
    const selected = b.dataset.tab === state.tab;
    b.classList.toggle("active", selected || (state.tab === "revision" && b.dataset.tab === "summary"));
    if (b.dataset.tab === "ledger") b.disabled = Boolean(state.busy);
    if (selected) b.setAttribute("aria-current", "page");
    else b.removeAttribute("aria-current");
  });
  const alerts =
    (state.error
      ? '<div class="error" role="alert">' +
        i("circle-alert") +
        "<span>" +
        esc(state.error) +
        '</span><button class="icon-btn" aria-label="关闭提示" data-act="dismiss">' +
        i("x") +
        "</button></div>"
      : "") +
    (state.busy
      ? '<div class="busy" role="status">' +
        i("loader-circle") +
        esc(state.busy) +
        "</div>"
      : "");
  const overview = state.result
    ? currentFile()
    : '<div class="import-grid" id="drop">' + importPanel() + sheetPanel() + "</div>";
  main.innerHTML =
    pageHeader() +
    alerts +
    (state.tab === "summary"
      ? overview + (state.result ? resultPanel() : "")
      : state.tab === "revision"
        ? revisionPanel(revision, button, state.busy)
        : state.tab === "ledger"
          ? ledgerPanel()
          : state.tab === "issue"
            ? issuePanel()
            : departPanel()) +
    '<footer class="privacy">工作簿仅在当前浏览器中处理。</footer>';
  updateLedgerRows();
  paint();
  if (focusId) document.getElementById(focusId)?.focus({ preventScroll: true });
  if (focusSheet)
    main
      .querySelector(
        'input[name="sheet"][value="' + CSS.escape(focusSheet) + '"]',
      )
      ?.focus({ preventScroll: true });
}
async function task(label, fn) {
  if (state.busy) return;
  state.busy = label;
  state.error = "";
  render();
  try {
    await fn();
  } catch (e) {
    state.error = e.message || "操作失败，请重试。";
  } finally {
    state.busy = "";
    render();
  }
}
function showResult(value) {
  if (!state.result) ledger.views.delete("00");
  ledger.tables.set("00", resultTable(value.records));
  state.result = value;
  rowIndex = new Map(value.records.map((r, j) => [r, j]));
}
function confirmReplace(action) {
  if (!state.dirty) return action();
  confirmation = action;
  openDialog(
    "更换当前工作簿？",
    '<p>当前工作簿有尚未下载的汇总或发放记录。请确认是否放弃这些更改。</p><div class="dialog-actions">' +
      button("close", "保留当前工作簿", "x") +
      button("confirm", "放弃并更换", "folder-open", "primary") +
      "</div>",
  );
}
async function loadFile(file, sample = false) {
  return task("正在读取工作簿…", async () => {
    if (sample) {
      const response = await fetch(sampleURL);
      if (!response.ok) throw new Error("样例读取失败，请重试。");
      file = new File([await response.blob()], "证书汇总输入样例.xlsx");
    }
    if (!file.name.toLowerCase().endsWith(".xlsx"))
      throw new Error("请选择 .xlsx 格式的 Excel 工作簿。");
    const bytes = await file.arrayBuffer();
    const meta = await rpc("open", { bytes }, [bytes]);
    ledger.code = "00";
    ledger.tables.clear();
    ledger.views.clear();
    Object.assign(state, {
      file: file.name,
      meta,
      result: null,
      dirty: false,
      query: "",
      level: "",
      page: 0,
      ids: "",
      dept: "",
      codes: new Set(["00", "01", "02"]),
      tab: "summary",
    });
    toast(sample ? "样例已载入，可以开始汇总。" : "工作簿已载入。");
  });
}
function download(bytes, name) {
  const url = URL.createObjectURL(
    new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
  );
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}
function detail(index) {
  const r = state.result.records[index];
  openDialog(
    "证书详情",
    '<div class="detail-name">' +
      esc(displayName(r)) +
      "</div><dl>" +
      [
        ["持证员工", r.name + " · " + r.id],
        ["所属部门", [r.dept, r.team].filter(Boolean).join(" / ")],
        ["获得日期", r.date],
        ["有效期", "永久"],
        ["本人最高级别", r.highest ? "是" : "否"],
        ["发放状态", r.status || "待标记"],
        ["生成时间", r.generated],
      ]
        .map(
          ([key, v]) =>
            "<div><dt>" + key + "</dt><dd>" + esc(v || "—") + "</dd></div>",
        )
        .join("") +
      '</dl><h3>授予依据</h3><ul class="trace">' +
      r.trace
        .map(
          (item) =>
            "<li><small>" +
            item.sheet +
            " 表 · 第 " +
            item.row +
            " 行 · " +
            esc(item.date) +
            "</small><span>" +
            esc(item.title) +
            "</span></li>",
        )
        .join("") +
      "</ul>",
  );
}
const actions = {
  changes: showChanges,
  "change-prev": () => turnChangePage(-1),
  "change-next": () => turnChangePage(1),
  guide: () =>
    openDialog(
      "使用指南",
      '<ol class="instructions"><li><b>导入完整工作簿</b><p>准备名称以 00、01、02、03 开头的四张表。首次汇总的 00 可完全空白。</p></li><li><b>汇总并核对</b><p>点击「开始汇总」。每行的「详情」可查看授予依据。已有 00 记录时，可核对本次新增、移出和获得日期调整的证书。</p></li><li><b>标记并下载</b><p>在「发放管理」粘贴工号，核对匹配数量后标记。下载完整工作簿保存结果，后续可再次导入。</p></li><li><b>按部门导出</b><p>选择所属部门和工作表，下载该部门的记录。</p></li></ol>',
    ),
  close: () => dialog.close(),
  confirm: () => {
    const go = confirmation;
    confirmation = null;
    dialog.close();
    go?.();
  },
  dismiss: () => {
    state.error = "";
    render();
  },
  choose: () => input.click(),
  sample: () => confirmReplace(() => loadFile(null, true)),
  "to-summary": () => {
    state.tab = "summary";
    render();
  },
  "to-revision": () => {
    state.tab = "revision";
    state.error = "";
    render();
    main.focus({ preventScroll: true });
    window.scrollTo(0, 0);
  },
  "revision-pick-previous": () => revisionInputs.previous.click(),
  "revision-pick-latest": () => revisionInputs.latest.click(),
  "revision-review": () =>
    task("正在比对来源并试算证书…", async () => {
      clearRevisionReview();
      const previous = await revision.previous.arrayBuffer();
      const latest = await revision.latest.arrayBuffer();
      revision.report = await rpc("revision-review", { previous, latest, time: localStamp() }, [previous, latest]);
    }),
  "revision-export": () =>
    task("正在生成新版工作簿…", async () => {
      const { scope, issued, statusSource } = revision;
      download(await rpc("revision-export", { scope, issued, statusSource }),
        revision.latest.name.replace(/(?:_汇总)?\.xlsx$/i, "_汇总.xlsx"));
      toast("新版汇总工作簿已生成。");
    }),
  "revision-detail-history": () => showRevisionDetails("history"),
  "revision-detail-exams": () => showRevisionDetails("exams"),
  "revision-detail-rules": () => showRevisionDetails("rules"),
  "revision-detail-awards": () => showRevisionDetails("awards"),
  "revision-detail-conflicts": () => showRevisionDetails("conflicts"),
  "revision-detail-prev": () => turnRevisionPage(-1),
  "revision-detail-next": () => turnRevisionPage(1),
  "to-ledger": async () => {
    window.scrollTo(0, 0);
    await openLedger("00");
    if (state.tab === "ledger") main.focus({ preventScroll: true });
  },
  "ledger-retry": () => openLedger(),
  "ledger-clear": () => {
    const view = ledgerView();
    view.query = "";
    view.filters.clear();
    view.page = 0;
    main.querySelectorAll("#ledger-search, [data-ledger-filter]").forEach((field) => {
      field.value = "";
    });
    main.querySelector("#ledger-search").focus();
    updateLedgerRows();
  },
  "ledger-prev": () => {
    ledgerView().page--;
    updateLedgerRows();
  },
  "ledger-next": () => {
    ledgerView().page++;
    updateLedgerRows();
  },
  run: () =>
    task("正在匹配考试规则并合并历史证书…", async () => {
      showResult(await rpc("run", { time: localStamp() }));
      state.dirty = true;
      state.page = 0;
      toast("已完成汇总，共 " + state.result.records.length + " 条证书。");
    }),
  download: () =>
    task("正在生成完整工作簿…", async () => {
      download(
        await rpc("export"),
        state.file.replace(/(?:_汇总)?\.xlsx$/i, "_汇总.xlsx"),
      );
      state.dirty = false;
      toast("汇总工作簿已生成。");
    }),
  mark: () =>
    task("正在更新发放状态…", async () => {
      const m = matchIds();
      showResult(await rpc("mark", { ids: m.ids, status: state.status }));
      state.dirty = true;
      toast("已将 " + m.matched.length + " 张证书标记为" + state.status + "。");
    }),
  "dept-export": () =>
    task("正在生成部门工作簿…", async () => {
      const group = state.result.departments.find((d) => d.key === state.dept);
      download(
        await rpc("export", {
          codes: [...state.codes],
          department: state.dept,
        }),
        group.label.replace(/[\\/:*?"<>|]/g, "-") + "_证书明细.xlsx",
      );
      toast("部门工作簿已生成。");
    }),
  prev: () => {
    state.page--;
    render();
  },
  next: () => {
    state.page++;
    render();
  },
};
document.addEventListener("click", async (e) => {
  if (e.target.closest("[data-home]")) {
    e.preventDefault();
    state.tab = "summary";
    render();
    main.focus();
    return;
  }
  const btn = e.target.closest("button");
  if (!btn || btn.disabled) return;
  if (btn.dataset.ledgerSheet) {
    const code = btn.dataset.ledgerSheet;
    await openLedger(code);
    if (state.tab === "ledger" && ledger.code === code)
      main.querySelector(`[data-ledger-sheet="${code}"]`)?.focus({ preventScroll: true });
  } else if (btn.hasAttribute("data-ledger-sort")) {
    orderLedger(btn);
  } else if (btn.dataset.tab === "ledger") {
    window.scrollTo(0, 0);
    await openLedger();
    if (state.tab === "ledger") main.focus({ preventScroll: true });
  } else if (btn.dataset.tab) {
    state.tab = btn.dataset.tab;
    render();
    main.focus({ preventScroll: true });
    window.scrollTo(0, 0);
  } else if (btn.hasAttribute("data-detail"))
    detail(Number(btn.dataset.detail));
  else actions[btn.dataset.act]?.();
});
dialog.addEventListener("change", (e) => {
  if (e.target.id === "issued-only") {
    changePage = 0;
    updateChangeList();
  }
  if (e.target.id === "revision-detail-filter") {
    revisionDetail.filter = e.target.value;
    revisionDetail.page = 0;
    updateRevisionDetails();
  }
});
dialog.addEventListener("click", (e) => {
  if (e.target === dialog) {
    const b = dialog.getBoundingClientRect();
    if (
      e.clientX < b.left ||
      e.clientX > b.right ||
      e.clientY < b.top ||
      e.clientY > b.bottom
    )
      dialog.close();
  }
});
input.addEventListener("change", () => {
  const f = input.files[0];
  input.value = "";
  if (f) confirmReplace(() => loadFile(f));
});
for (const [role, picker] of Object.entries(revisionInputs)) {
  picker.addEventListener("change", () => {
    const file = picker.files[0];
    picker.value = "";
    if (!file || state.busy) return;
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      state.error = "请选择 .xlsx 格式的 Excel 工作簿。";
    } else {
      revision[role] = file;
      clearRevisionReview();
      state.error = "";
    }
    render();
  });
}
main.addEventListener("input", (e) => {
  if (e.target.id === "ledger-search" || e.target.hasAttribute("data-ledger-filter")) {
    const { value } = e.target;
    const view = ledgerView();
    if (e.target.id === "ledger-search") view.query = value;
    else {
      const column = Number(e.target.dataset.ledgerFilter);
      if (value) view.filters.set(column, value);
      else view.filters.delete(column);
    }
    view.page = 0;
    updateLedgerRows();
    return;
  }
  if (e.target.id === "search") {
    state.query = e.target.value;
    state.page = 0;
    document.querySelector("#table").innerHTML = tableContent();
    paint();
  }
  if (e.target.id === "ids") {
    state.ids = e.target.value;
    document.querySelector("#match").innerHTML = matchPreview();
    paint();
  }
});
main.addEventListener("change", (e) => {
  const el = e.target;
  if (el.id === "level") {
    state.level = el.value;
    state.page = 0;
  } else if (el.id === "dept") state.dept = el.value;
  else if (el.id === "status") state.status = el.value;
  else if (el.id === "revision-scope") revision.scope = el.checked;
  else if (el.id === "revision-issued") revision.issued = el.checked;
  else if (el.name === "revision-status") revision.statusSource = el.value;
  else if (el.name === "sheet") {
    if (el.checked) state.codes.add(el.value);
    else state.codes.delete(el.value);
  } else return;
  render();
});
for (const evt of ["dragover", "dragleave", "drop"])
  document.addEventListener(evt, (e) => {
    if (!e.dataTransfer?.types.includes("Files")) return;
    e.preventDefault();
    const zone = e.target.closest("#drop");
    if (!zone || state.busy) return;
    zone.classList.toggle("dragging", evt === "dragover");
    if (evt === "drop") {
      const f = e.dataTransfer.files[0];
      if (f) confirmReplace(() => loadFile(f));
    }
  });
window.addEventListener("beforeunload", (e) => {
  if (state.dirty) {
    e.preventDefault();
    e.returnValue = "";
  }
});
render();
