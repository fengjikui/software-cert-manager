import { certKey, compareAwards, normalize, summarize } from "./engine.js";
import { cleanHeader, workbookTable } from "./workbook.js";

function sourceRows(data, code) {
  const history = code === "01";
  const table = workbookTable(data, code);
  const parsed = new Map((history ? data.history : data.exams).map((r) => [r.row, r]));
  const headings = new Map();
  const columns = table.columns.map((label, index) => {
    const heading = cleanHeader(label).replace(/^学员(工号|姓名)$/, "$1");
    const occurrence = (headings.get(heading) ?? 0) + 1;
    headings.set(heading, occurrence);
    return { key: JSON.stringify([heading, occurrence]), heading, label, index };
  }).sort((a, b) => a.key.localeCompare(b.key));
  const startField = history ? "证书获得日期" : "开始答题时间";
  return table.rows.map(({ line, values }) => {
    const person = parsed.get(line);
    const { id, name, title, time, rank, lang } = person;
    const fields = new Map();
    let learnerCode = "";
    for (const { key, heading, label, index } of columns) {
      const value = values[index].trim();
      const comparable = heading === startField ? time : value;
      if (comparable) fields.set(key, { label, value, comparable });
      if (heading === "学员证书编码") learnerCode = value;
    }
    const exam = normalize(title);
    const fallback = history ? certKey({ id, rank, lang })
      : time ? JSON.stringify([id, exam, time]) : "";
    return {
      line, id, name, title, time, exam, learnerCode, fallback, fields,
      signature: JSON.stringify([...fields].map(([key, { comparable }]) => [key, comparable])),
    };
  });
}

function groupBy(rows, keyFor) {
  const groups = new Map();
  for (const row of rows) {
    const key = keyFor(row);
    if (!key) continue;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }
  return groups;
}

function changedFields(before, after) {
  const changes = [];
  for (const key of new Set([...before.fields.keys(), ...after.fields.keys()])) {
    const old = before.fields.get(key), next = after.fields.get(key);
    if (old?.comparable !== next?.comparable)
      changes.push({ label: next?.label ?? old.label, before: old?.value ?? "", after: next?.value ?? "" });
  }
  return changes;
}

function compareSource(previous, latest, code) {
  const before = new Set(sourceRows(previous, code));
  const after = new Set(sourceRows(latest, code));
  const oldCount = before.size, newCount = after.size;
  const oldCodes = groupBy(before, (r) => r.learnerCode);
  const newCodes = groupBy(after, (r) => r.learnerCode);
  const exact = groupBy(after, (r) => r.signature);
  for (const row of before) {
    const match = exact.get(row.signature)?.pop();
    if (match) {
      before.delete(row);
      after.delete(match);
    }
  }

  const counts = { added: 0, modified: 0, missing: 0, ambiguous: 0 };
  const items = [];
  const ref = ({ line, id, name, title, time }) => ({ line, id, name, title, time });
  function add(kind, old, next, fields = []) {
    counts[kind]++;
    items.push({ kind, before: old.map(ref), after: next.map(ref), fields });
  }
  function match(keyFor, canPair) {
    const candidates = groupBy(after, keyFor);
    for (const [key, old] of groupBy(before, keyFor)) {
      const next = candidates.get(key);
      if (!next) continue;
      if (old.length === 1 && next.length === 1 && canPair(old[0], next[0])) {
        const fields = changedFields(old[0], next[0]);
        if (fields.length) add("modified", old, next, fields);
      } else add("ambiguous", old, next);
      old.forEach((row) => before.delete(row));
      next.forEach((row) => after.delete(row));
    }
  }
  if (code === "01") {
    match((r) => r.learnerCode, (r) =>
      oldCodes.get(r.learnerCode).length === 1 && newCodes.get(r.learnerCode).length === 1,
    );
    match((r) => r.fallback, (a, b) => !a.learnerCode || !b.learnerCode);
  } else {
    match((r) => r.fallback, () => true);
    // Changes to the matching fields are possible links, never confirmed modifications.
    match((r) => JSON.stringify([r.id, r.exam]), () => false);
    match((r) => r.time && JSON.stringify([r.id, r.time]), () => false);
  }
  for (const row of before) add("missing", [row], []);
  for (const row of after) add("added", [], [row]);
  return { oldCount, newCount, counts, items };
}

function compareRules(before, after) {
  const signature = (rule) => JSON.stringify(rule.subjects.map((exams) => [...new Set(exams)].sort()));
  const old = new Map(before.map((r) => [JSON.stringify([r.rank, r.lang]), signature(r)]));
  const changes = [];
  for (const rule of after) {
    const key = JSON.stringify([rule.rank, rule.lang]);
    const both = old.has(key);
    if (old.get(key) !== signature(rule))
      changes.push({ rank: rule.rank, lang: rule.lang, kind: both ? "modified" : "added" });
    old.delete(key);
  }
  for (const key of old.keys()) {
    const [rank, lang] = JSON.parse(key);
    changes.push({ rank, lang, kind: "missing" });
  }
  return changes;
}

export function previewRevision(previous, latest, time) {
  const result = summarize(latest, previous.baseline, time);
  const awards = compareAwards(previous.baseline, result.records);
  const old = new Map(previous.baseline.map((row) => [certKey(row), row]));
  const generated = new Set(result.records.map(certKey));
  const conflicts = latest.baseline.flatMap((row) => {
    const key = certKey(row);
    const prior = old.get(key);
    if (!row.status || row.status === prior?.status) return [];
    return [{
      key, id: row.id, name: row.name || prior?.name, rank: row.rank, lang: row.lang,
      oldStatus: prior?.status ?? "", newStatus: row.status,
      oldRow: prior?.row, newRow: row.row, generated: generated.has(key),
    }];
  });
  const issuedMissing = new Set(
    [...previous.baseline, ...latest.baseline]
      .filter((row) => row.status === "已发放")
      .map(certKey)
      .filter((key) => !generated.has(key)),
  ).size;
  return {
    records: result.records,
    report: {
      history: compareSource(previous, latest, "01"),
      exams: compareSource(previous, latest, "02"),
      rules: compareRules(previous.rules, latest.rules),
      awards,
      conflicts,
      previousAwardCount: previous.baseline.length,
      recordCount: result.records.length,
      unmatched: result.unmatched,
      issuedMissing,
    },
  };
}

export function revisionRecords(preview, statusSource) {
  const { records, report } = preview;
  if (report.conflicts.length && !["previous", "latest"].includes(statusSource))
    throw new Error("请先确定采用哪一版的发放标记。");
  if (statusSource !== "latest") return records;
  const statuses = new Map(report.conflicts.map((row) => [row.key, row.newStatus]));
  return records.map((row) => {
    const status = statuses.get(certKey(row));
    return status ? { ...row, status } : row;
  });
}
