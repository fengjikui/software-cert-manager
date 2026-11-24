export const levels = ["入门级", "工作级", "专业级"];
export const headers = [
  "工号",
  "姓名",
  "证书名称",
  "证书级别",
  "证书语言类别",
  "是否最高级别证书",
  "证书获得日期",
  "证书到期日期",
  "0层部门",
  "一层部门",
  "证书发放状态",
  "生成时间",
];
export const normalize = (value) =>
  String(value ?? "")
    .replace(/\s/g, "")
    .replaceAll("（", "(")
    .replaceAll("）", ")")
    .replace(/[，、]/g, ",")
    .replaceAll("JavaScript&TypeScript", "JavaScript")
    .replaceAll("ArkTSt", "ArkTS");

export function parseCertificate(name, location) {
  const match =
    /^(?:华为)?软件开发能力认证证书\((入门级|工作级|专业级),([^,()]+)\)$/.exec(
      normalize(name),
    );
  if (!match)
    throw new Error(location + " 的证书名称无法识别：" + (name || "空白"));
  return { rank: levels.indexOf(match[1]), lang: match[2] };
}

export const certKey = (r) => JSON.stringify([r.id, r.rank, r.lang]);
export const departmentKey = (dept, team) => JSON.stringify([dept, team]);
export const displayName = (r) =>
  "软件开发能力认证证书（" + levels[r.rank] + "，" + r.lang + "）";
export const cellsOf = (r) => [
  r.id,
  r.name,
  displayName(r),
  levels[r.rank],
  r.lang,
  r.highest ? "是" : "否",
  r.date,
  "永久",
  r.dept,
  r.team,
  r.status,
  r.generated,
];

function compareCertificates(a, b) {
  return a.id.localeCompare(b.id) || a.rank - b.rank || a.lang.localeCompare(b.lang);
}

export function compareAwards(baseline, records) {
  const previous = new Map(baseline.map((r) => [certKey(r), r]));
  const items = [];
  const counts = { added: 0, removed: 0, dateChanged: 0 };
  let issuedCount = 0;
  function add(kind, before, after) {
    const { id, rank, lang } = after ?? before;
    const wasIssued = before?.status === "已发放";
    if (wasIssued) issuedCount++;
    counts[kind]++;
    items.push({
      kind, id, rank, lang, wasIssued,
      name: after?.name || before?.name || "",
      beforeDate: before?.date,
      afterDate: after?.date,
      sourceRow: before?.row,
    });
  }
  for (const current of records) {
    const key = certKey(current);
    const prior = previous.get(key);
    if (!prior) add("added", undefined, current);
    else if (prior.date !== current.date) add("dateChanged", prior, current);
    previous.delete(key);
  }
  for (const prior of previous.values()) add("removed", prior, undefined);
  items.sort((a, b) => Number(b.wasIssued) - Number(a.wasIssued) || compareCertificates(a, b));
  return { items, counts, issuedCount };
}

function compileRules(rules) {
  return rules.map((rule) => {
    const subjects = rule.subjects.map((exams) => new Set(exams));
    for (const other of rules) {
      if (other.rank > rule.rank && other.lang === rule.lang) {
        for (const i of [0, 1])
          for (const exam of other.subjects[i]) subjects[i].add(exam);
      }
      if (rule.rank === 1 && other.rank === 2 && other.lang !== "C-系统软件") {
        for (const exam of other.subjects[2]) subjects[2].add(exam);
      }
    }
    return {
      ...rule,
      required: subjects
        .map((exams, i) => ({ exams, i }))
        .filter(({ i }) => rule.subjects[i].length),
    };
  });
}

export function summarize(data, baseline, runTime) {
  const old = new Map(baseline.map((r) => [certKey(r), r]));
  const departments = new Map(),
    passes = new Map(),
    awarded = new Map();
  const rules = compileRules(data.rules);
  const known = new Set(data.rules.flatMap((rule) => rule.subjects.flat()));
  let unmatched = 0;
  function updateDepartment(row) {
    const previous = departments.get(row.id);
    if (row.dept && (!previous || row.time >= previous.time))
      departments.set(row.id, row);
  }
  function add(row, trace) {
    const key = certKey(row),
      previous = awarded.get(key);
    if (previous) {
      previous.trace.push(...trace);
      if (row.date < previous.date) Object.assign(previous, row);
    } else awarded.set(key, { ...row, trace });
  }
  for (const row of data.history) {
    updateDepartment(row);
    add(row, [{ sheet: "01", row: row.row, title: row.title, date: row.date }]);
  }
  for (const row of data.exams) {
    updateDepartment(row);
    if (!row.passed) continue;
    const exam = normalize(row.title);
    if (!known.has(exam)) {
      unmatched++;
      continue;
    }
    if (!passes.has(row.id)) passes.set(row.id, new Map());
    const byExam = passes.get(row.id);
    if (!byExam.has(exam) || row.time < byExam.get(exam).time)
      byExam.set(exam, row);
  }
  for (const [id, exams] of passes) {
    const byDay = new Map();
    for (const { lang, rank, required } of rules) {
      const trace = [];
      let last;
      for (const subject of required) {
        let earliest;
        for (const key of subject.exams) {
          const row = exams.get(key);
          if (row && (!earliest || row.time < earliest.time)) earliest = row;
        }
        if (!earliest) break;
        if (!last || last.time < earliest.time) last = earliest;
        trace.push({
          sheet: "02",
          row: earliest.row,
          title: earliest.title,
          date: earliest.date,
          subject: subject.i + 1,
        });
      }
      if (!last || trace.length !== required.length) continue;
      const key = JSON.stringify([lang, last.date]),
        previous = byDay.get(key);
      if (!previous || previous.rank < rank)
        byDay.set(key, {
          id,
          name: last.name,
          lang,
          rank,
          date: last.date,
          trace,
        });
    }
    for (const { trace, ...r } of byDay.values()) add(r, trace);
  }
  const maxima = new Map();
  for (const r of awarded.values())
    maxima.set(r.id, Math.max(maxima.get(r.id) ?? -1, r.rank));
  const records = Array.from(awarded.values(), (r) => {
    const previous = old.get(certKey(r)),
      org = departments.get(r.id);
    return {
      ...r,
      highest: r.rank === maxima.get(r.id),
      dept: org?.dept ?? "",
      team: org?.team ?? "",
      status: previous?.status ?? "",
      generated:
        previous?.date === r.date && previous.generated
          ? previous.generated
          : runTime,
    };
  });
  records.sort(compareCertificates);
  const changes = data.baseline.length ? compareAwards(data.baseline, records) : null;
  return { records, departments, unmatched, changes };
}
