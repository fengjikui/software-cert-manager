import { displayName, levels } from "./engine.js";
import { esc } from "./markup.js";

const sourceLabels = { added: "新增", modified: "修改", missing: "本次缺失", ambiguous: "待核对" };
const awardLabels = { added: "新增", removed: "拟移出", dateChanged: "日期调整" };
const sections = {
  history: "01 历史证书",
  exams: "02 考试成绩",
  rules: "03 认证规则",
  awards: "00 证书影响",
  conflicts: "发放标记差异",
};

const summaryLine = (counts, labels) =>
  Object.entries(labels).map(([key, label]) => label + " " + counts[key]).join(" · ");
const check = (value) => value ? "checked" : "";

export function revisionPanel(view, button, busy) {
  const { report, previous, latest } = view;
  const ready = previous && latest;
  function file(role, label, file, note) {
    return `<section class="revision-file">
      <div><h2>${label}</h2><p>${esc(file?.name ?? "尚未选择")}</p><small>${note}</small></div>
      ${button("revision-pick-" + role, file ? "更换文件" : "选择文件", "")}
    </section>`;
  }
  const pickers = `<div class="revision-files">
    ${file("previous", "上次正式汇总", previous, "提供旧 01、02 和 00 发放记录")}
    ${file("latest", "本次客户工作簿", latest, "提供本次完整的 01、02、03")}
    </div>
    <div class="revision-toolbar">
      <p>本次文件应覆盖与上次相同的部门、人员及时间范围。</p>
      ${button("revision-review", report ? "重新比对" : "比对并预览", "", "secondary", !ready)}
    </div>`;
  if (!report) return pickers;

  const detail = (section) =>
    `<button class="btn secondary" data-act="revision-detail-${section}"
      aria-label="查看${sections[section]}明细" ${busy ? "disabled" : ""}>查看明细</button>`;
  const line = (section, summary, available) =>
    `<div class="revision-report-row">
      <div><h3>${sections[section]}</h3><p>${summary}</p></div>
      ${available ? detail(section) : ""}
    </div>`;
  const source = (section) => {
    const { oldCount, newCount, counts, items } = report[section];
    return line(section,
      `上次 ${oldCount} 条 → 本次 ${newCount} 条 · ${summaryLine(counts, sourceLabels)}`,
      items.length,
    );
  };
  const ruleCount = report.rules.length;
  const { awards, issuedMissing, conflicts } = report;
  const disabled = busy ? "disabled" : "";
  const statusChoice = conflicts.length
    ? `<fieldset class="revision-status">
        <legend>选择发放标记来源</legend>
        <p>新版有 ${conflicts.length} 项非空标记与上次台账不同。此选择只影响有差异的标记。</p>
        <label><input type="radio" id="revision-status-previous" name="revision-status"
          value="previous" ${check(view.statusSource === "previous")} ${disabled}> 沿用上次台账</label>
        <label><input type="radio" id="revision-status-latest" name="revision-status"
          value="latest" ${check(view.statusSource === "latest")} ${disabled}> 采用新版非空标记</label>
        ${detail("conflicts")}
      </fieldset>`
    : "";
  return pickers + `<section class="revision-report">
    <h2>换版预览</h2>
    ${source("history")}${source("exams")}
    ${line("rules", ruleCount ? `${ruleCount} 项认证规则发生变化` : "认证规则无变化", ruleCount)}
    ${line("awards",
      `预计生成 ${report.recordCount} 张证书 · ${summaryLine(awards.counts, awardLabels)}`,
      awards.items.length,
    )}
    ${report.previousAwardCount ? "" : '<p class="hint">上次工作簿的 00 表为空，可继承的发放记录为 0。</p>'}
    ${report.unmatched ? `<p class="warning">新版有 ${report.unmatched} 条通过记录未匹配 03 中的考试名称。</p>` : ""}
    ${issuedMissing ? `<p class="warning">有 ${issuedMissing} 张曾标记为已发放的证书，本次未生成。</p>` : ""}
    ${statusChoice}
    <div class="revision-confirm">
      <label><input id="revision-scope" type="checkbox" ${check(view.scope)} ${disabled}>
        我已确认两份工作簿覆盖相同范围，本次文件含完整记录。</label>
      ${issuedMissing ? `<label><input id="revision-issued" type="checkbox" ${check(view.issued)} ${disabled}>
        我已核对本次未生成的已发放证书。</label>` : ""}
      ${button("revision-export", "下载新版工作簿", "", "primary",
        !view.scope || (issuedMissing && !view.issued) || (conflicts.length && !view.statusSource),
      )}
      <small>结果将写入本次工作簿的 00 表，两个输入文件可留作回查。</small>
    </div></section>`;
}

function sourceItem({ kind, before, after, fields }) {
  const refs = (records, label) => records.map(({ line, name, id, title, time }) =>
    `<li><small>${label}第 ${line} 行</small> ${esc([id, name, title, time].filter(Boolean).join(" · "))}</li>`,
  ).join("");
  const fieldList = fields.length ? `<ul class="revision-fields">${fields.map(({ label, before, after }) =>
    `<li><b>${esc(label)}</b><span>${esc(before || "空白")} → ${esc(after || "空白")}</span></li>`,
  ).join("")}</ul>` : "";
  return `<li><b>${sourceLabels[kind]}</b>
    <ul class="revision-refs">${refs(before, "原表")}${refs(after, "新版")}</ul>${fieldList}</li>`;
}

function awardItem(item) {
  const before = item.beforeDate || "未生成";
  const after = item.afterDate || "未生成";
  return `<li><b>${awardLabels[item.kind]} · ${esc(item.id)} ${esc(item.name)}</b>
    <div>${esc(item.lang)} · ${levels[item.rank]}${item.wasIssued ? " · 原已发放" : ""}</div>
    <small>${esc(before)} → ${esc(after)}${item.sourceRow ? " · 原 00 第 " + item.sourceRow + " 行" : ""}</small></li>`;
}

function conflictItem(item) {
  const sources = [
    item.oldRow ? "原 00 第 " + item.oldRow + " 行" : "原 00 无此证书",
    "新版 00 第 " + item.newRow + " 行",
    item.generated ? "" : "本次未生成",
  ].filter(Boolean).join(" · ");
  return `<li><b>${esc(item.id)} ${esc(item.name)} · ${levels[item.rank]} · ${esc(item.lang)}</b>
    <div>原标记 ${esc(item.oldStatus || "空白")} → 新版标记 ${esc(item.newStatus)}</div>
    <small>${sources}</small></li>`;
}

export function revisionDetails(report, section) {
  const title = sections[section];
  if (section === "rules") return {
    title, items: report.rules, labels: sourceLabels,
    render: (rule) => `<li><b>${sourceLabels[rule.kind]}</b><div>${esc(displayName(rule))}</div></li>`,
  };
  if (section === "conflicts")
    return { title, items: report.conflicts, render: conflictItem };
  if (section === "awards")
    return { title, items: report.awards.items, render: awardItem, labels: awardLabels };
  return { title, items: report[section].items, render: sourceItem, labels: sourceLabels };
}
