import { openWorkbook, exportFull, exportDepartment, workbookTable } from "./workbook.js";
import { summarize, departmentKey } from "./engine.js";
import { previewRevision, revisionRecords } from "./revision.js";

let data, result, revision;
function revisionWorkbook(bytes, label) {
  try {
    return openWorkbook(bytes);
  } catch (error) {
    throw new Error(label + "：" + error.message);
  }
}
function summary() {
  const departments = new Map();
  for (const row of result.departments.values()) {
    const key = departmentKey(row.dept, row.team);
    if (!departments.has(key))
      departments.set(key, {
        key,
        label: [row.dept, row.team].filter(Boolean).join(" / "),
        counts: [0, 0, 0],
      });
  }
  [result.records, data.history, data.exams].forEach((rows, i) => {
    for (const row of rows) {
      const org = result.departments.get(row.id);
      if (org) departments.get(departmentKey(org.dept, org.team)).counts[i]++;
    }
  });
  return {
    records: result.records,
    changes: result.changes,
    unmatched: result.unmatched,
    departments: [...departments.values()].sort((a, b) =>
      a.label.localeCompare(b.label, "zh"),
    ),
  };
}
self.onmessage = ({ data: msg }) => {
  try {
    let value;
    if (msg.action === "open") {
      const next = openWorkbook(msg.bytes);
      data = next;
      result = null;
      value = {
        names: Object.fromEntries(
          Object.entries(data.tabs).map(([c, t]) => [c, t.name]),
        ),
        counts: [
          data.baseline.length,
          data.history.length,
          data.exams.length,
          data.rules.length,
        ],
      };
    } else if (msg.action === "revision-review") {
      revision = null;
      const previous = revisionWorkbook(msg.previous, "上次正式汇总");
      const latest = revisionWorkbook(msg.latest, "本次客户工作簿");
      const preview = previewRevision(previous, latest, msg.time);
      revision = { latest, preview };
      value = preview.report;
    } else if (msg.action === "revision-export") {
      if (!revision || !msg.scope || (revision.preview.report.issuedMissing && !msg.issued))
        throw new Error("请先完成预览并确认导出范围与发放记录。");
      const bytes = exportFull(revision.latest, revisionRecords(revision.preview, msg.statusSource));
      self.postMessage({ value: bytes.buffer }, [bytes.buffer]);
      return;
    } else if (msg.action === "table") {
      value = workbookTable(data, msg.code);
    } else if (msg.action === "run") {
      result = summarize(data, result?.records ?? data.baseline, msg.time);
      value = summary();
    } else if (msg.action === "mark") {
      const ids = new Set(msg.ids);
      for (const record of result.records)
        if (ids.has(record.id)) record.status = msg.status;
      value = summary();
    } else if (msg.action === "export") {
      const bytes = msg.codes
        ? exportDepartment(data, result, msg.department, msg.codes)
        : exportFull(data, result.records);
      const buffer = bytes instanceof ArrayBuffer ? bytes : bytes.buffer;
      self.postMessage({ value: buffer }, [buffer]);
      return;
    }
    self.postMessage({ value });
  } catch (error) {
    self.postMessage({ error: error.message || "处理失败，请检查工作簿。" });
  }
};
