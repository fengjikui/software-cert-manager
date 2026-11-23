import { openWorkbook, exportFull, exportDepartment } from "./workbook.js";
import { summarize, departmentKey } from "./engine.js";

let data, result;
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
