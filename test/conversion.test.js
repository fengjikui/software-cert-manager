import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import * as XLSX from "xlsx";
import { unzipSync } from "fflate";
import {
  openWorkbook,
  exportFull,
  departmentRows,
  exportDepartment,
} from "../src/workbook.js";
import { cellsOf, departmentKey, summarize } from "../src/engine.js";

const dir = new URL(
  "../outputs/01a1e9d3-62a5-7b33-a92a-ac33e674e531/",
  import.meta.url,
);
const input = await readFile(new URL("证书汇总输入样例.xlsx", dir));
const expected = openWorkbook(
  await readFile(new URL("证书汇总预期结果.xlsx", dir)),
);
const time = "2026-11-17 10:00:00";
const later = "2026-11-24 09:00:00";
const first = (data) => summarize(data, data.baseline, time);
const ordered = (rows) => rows.map((r) => JSON.stringify(r)).sort();

test("首次汇总的全部 12 列与人工核对用样例逐格一致", () => {
  const data = openWorkbook(input),
    result = first(data);
  assert.equal(data.baseline.length, 0);
  const actual = openWorkbook(exportFull(data, result.records));
  assert.deepEqual(actual.book.SheetNames, expected.book.SheetNames);
  const values = (d) =>
    d.tabs["00"].rows.map((row) =>
      Array.from({ length: 12 }, (_, i) => row[i] ?? ""),
    );
  assert.deepEqual(ordered(values(actual)), ordered(values(expected)));
  assert.equal(actual.baseline.length, 10);
  assert.ok(
    result.records.every(
      (r) => r.trace.length && !cellsOf(r)[2].includes("华为"),
    ),
  );
});

test("重新导入保留标记，只在获得日期改变时刷新生成时间", () => {
  const data = openWorkbook(input),
    original = first(data);
  const record = original.records.find(
    (r) => r.id === "000001" && r.rank === 1,
  );
  record.status = "已发放";
  const bytes = exportFull(data, original.records),
    next = openWorkbook(bytes);
  const unchanged = summarize(next, next.baseline, later);
  const select = (result) =>
    result.records.find((r) => r.id === "000001" && r.rank === 1);
  assert.equal(select(unchanged).generated, time);
  const history = next.history.find((r) => r.id === "000001");
  history.date = "2026-11-01";
  const changed = select(summarize(next, next.baseline, later));
  assert.deepEqual(
    [changed.status, changed.date, changed.generated],
    ["已发放", "2026-11-01", later],
  );
  const prev = unzipSync(input),
    after = unzipSync(bytes);
  assert.deepEqual(Object.keys(prev).sort(), Object.keys(after).sort());
  for (const part of Object.keys(prev).filter(
    (p) => !["xl/worksheets/sheet1.xml", "xl/styles.xml"].includes(p),
  ))
    assert.deepEqual(prev[part], after[part], part);
});

test("按员工最新部门导出选中的工作表，原始部门列和工号保留", () => {
  const data = openWorkbook(input),
    result = first(data),
    key = departmentKey("平台研发部", "应用平台组");
  const grouped = departmentRows(data, result, key);
  assert.deepEqual(
    Object.values(grouped).map((g) => g.length),
    [4, 1, 9],
  );
  const exported = XLSX.read(
    exportDepartment(data, result, key, ["01", "02"]),
    { type: "array" },
  );
  assert.deepEqual(exported.SheetNames, [
    data.tabs["01"].name,
    data.tabs["02"].name,
  ]);
  assert.equal(exported.Sheets[exported.SheetNames[0]].A2.v, "000001");
  assert.equal(exported.Sheets[exported.SheetNames[0]].I2.v, "基础开发组");
});

test("数值工号的显示前导零和 1904 日期系统", () => {
  const data = openWorkbook(input),
    b = data.book;
  b.Workbook.WBProps.date1904 = true;
  for (const t of ["01", "02"]) {
    const tab = data.tabs[t],
      column = t === "01" ? "A" : "D";
    for (let r = 2; r <= tab.rows.length; r++) {
      const id = tab.sheet[column + r];
      id.v = Number(id.v);
      id.t = "n";
      id.z = "000000";
      delete id.w;
      tab.sheet["F" + r].v -= 1462;
    }
  }
  const changed = openWorkbook(
    XLSX.write(b, { bookType: "xlsx", type: "array" }),
  );
  assert.deepEqual(
    ordered(first(changed).records.map(cellsOf)),
    ordered(first(openWorkbook(input)).records.map(cellsOf)),
  );
  const roundtrip = openWorkbook(exportFull(changed, first(changed).records));
  assert.equal(roundtrip.baseline[0].date, "2026-11-02");
});
