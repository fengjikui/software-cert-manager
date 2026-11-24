import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { unzipSync } from "fflate";
import * as XLSX from "xlsx";
import { exportFull, openWorkbook } from "../src/workbook.js";
import { summarize } from "../src/engine.js";
import { previewRevision, revisionRecords } from "../src/revision.js";

const file = new URL(
  "../outputs/01a1e9d3-62a5-7b33-a92a-ac33e674e531/证书汇总输入样例.xlsx",
  import.meta.url,
);
const input = await readFile(file);
const oldTime = "2026-11-20 10:00:00";
const newTime = "2026-11-24 10:00:00";
const write = (data) => XLSX.write(data.book, { bookType: "xlsx", type: "array" });
const replaceRows = (data, code, rows) => {
  data.book.Sheets[data.tabs[code].name] = XLSX.utils.aoa_to_sheet(rows);
};

test("换版对照包含 01 新增和缺失、02 修改、03 规则变化和已发放证书影响", () => {
  const source = openWorkbook(input);
  const initial = summarize(source, [], oldTime).records;
  const java = initial.find((row) => row.lang === "Java");
  const rust = initial.find((row) => row.lang === "Rust");
  java.status = rust.status = "已发放";
  const previous = openWorkbook(exportFull(source, initial));

  const next = openWorkbook(input);
  const history = next.tabs["01"].rows;
  const added = [...history[1]];
  added[0] = "000099";
  added[1] = "样例员工";
  added[4] = "ISS-099";
  const rustEarlier = [...history[3]];
  rustEarlier[5]--;
  replaceRows(next, "01", [history[0], rustEarlier, history[4], added, history[1]]);
  next.tabs["02"].sheet.H17.v = "通过";
  next.tabs["03"].sheet.B2.v += "（新版）";
  const latestWithEmpty00 = openWorkbook(write(next));
  // Both files contain a marked copy of the same Rust certificate.
  const latest = openWorkbook(exportFull(latestWithEmpty00, [{ ...rust, status: "未发放" }]));
  const preview = previewRevision(previous, latest, newTime);
  const { report } = preview;
  assert.deepEqual(report.history.counts, { added: 1, modified: 1, missing: 1, ambiguous: 0 });
  const moved = report.history.items.find((item) => item.kind === "modified");
  assert.deepEqual([moved.before[0].line, moved.after[0].line, moved.fields[0].label], [4, 2, "证书获得日期"]);
  const exam = report.exams.items.find((item) => item.kind === "modified");
  assert.deepEqual(exam.fields, [{ label: "状态/Status", before: "判题中", after: "通过" }]);
  assert.equal(report.rules.length, 1);
  assert.equal(report.issuedMissing, 1);
  assert.ok(report.awards.items.some((item) => item.kind === "removed" && item.lang === "Java" && item.wasIssued));
  assert.throws(() => revisionRecords(preview), /发放标记/);
  const oldMark = revisionRecords(preview, "previous").find((row) => row.lang === "Rust");
  const newMark = revisionRecords(preview, "latest").find((row) => row.lang === "Rust");
  assert.deepEqual([oldMark.status, newMark.status, newMark.generated], ["已发放", "未发放", newTime]);
  const kept = revisionRecords(preview, "latest").find((row) => row.id === "000001" && row.rank === 1);
  assert.equal(kept.generated, oldTime);

  const bytes = exportFull(latest, revisionRecords(preview, "latest"));
  const result = openWorkbook(bytes);
  assert.deepEqual(result.book.SheetNames, latest.book.SheetNames);
  assert.equal(result.baseline.find((row) => row.lang === "Rust").status, "未发放");
  assert.equal(result.baseline.some((row) => row.lang === "Java"), false);
  const sourceParts = unzipSync(new Uint8Array(latest.bytes)), outputParts = unzipSync(bytes);
  for (const path of Object.keys(sourceParts).filter((p) => /^xl\/worksheets\/sheet[2-9]/.test(p)))
    assert.deepEqual(outputParts[path], sourceParts[path], path);
});

test("按行重排和匹配日期格式等价，答题改键与重复候选交人工核对", () => {
  const previous = openWorkbook(input);
  const next = openWorkbook(input);
  const [head, first, second, ...remaining] = next.tabs["02"].rows;
  const newTime = [...first];
  newTime[5] += 1 / 24;
  const secondChanged = [...second];
  secondChanged[7] = "判题中";
  const duplicate = [...secondChanged];
  duplicate[7] = "未通过";
  const another = [...remaining[0]];
  another[5] = previous.exams[2].time.replaceAll("-", "/");
  replaceRows(next, "02", [head, ...remaining.slice(1).reverse(), another, newTime, secondChanged, duplicate]);
  const preview = previewRevision(previous, openWorkbook(write(next)), oldTime).report;
  assert.deepEqual(preview.exams.counts, { added: 0, modified: 0, missing: 0, ambiguous: 2 });
  assert.equal(preview.exams.items.reduce((n, item) => n + item.after.length, 0), 3);
  assert.deepEqual(preview.history.items, []);
  assert.equal(preview.awards.counts.added, 8);
});
