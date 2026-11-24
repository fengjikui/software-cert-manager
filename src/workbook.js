import * as XLSX from "xlsx";
import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import {
  cellsOf,
  certKey,
  departmentKey,
  headers,
  normalize,
  parseCertificate,
} from "./engine.js";

const pad = (n) => String(n).padStart(2, "0");
const text = (v) => String(v ?? "").trim();
const cleanHeader = (v) => normalize(v).split("/")[0].replaceAll("零层", "0层");
const populated = (row) =>
  row?.some((v) => v !== null && v !== undefined && v !== "");
const sheetCodes = ["00", "01", "02", "03"];

function stamp(value, epoch1904, location) {
  let year,
    month,
    day,
    hour = 0,
    minute = 0,
    second = 0;
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(Math.round(value * 86400) / 86400, {
      date1904: epoch1904,
    });
    if (parsed)
      ({ y: year, m: month, d: day, H: hour, M: minute, S: second } = parsed);
  } else {
    const match =
      /^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2})(?:\.\d+)?)?)?$/.exec(
        text(value),
      );
    if (match)
      [year, month, day, hour = 0, minute = 0, second = 0] = match
        .slice(1)
        .map((n) => (n === undefined ? undefined : Number(n)));
  }
  const date = new Date(Date.UTC(year, month - 1, day));
  if (
    !year ||
    date.getUTCDate() !== day ||
    date.getUTCMonth() !== month - 1 ||
    hour > 23 ||
    minute > 59 ||
    second > 59
  )
    throw new Error(location + " 的日期为空或无法识别。");
  return (
    year +
    "-" +
    pad(month) +
    "-" +
    pad(day) +
    " " +
    [hour, minute, second].map(pad).join(":")
  );
}

function tab(book, code) {
  const names = book.SheetNames.filter((n) => n.trim().startsWith(code));
  if (names.length !== 1)
    throw new Error(
      "请保留一张名称以 " +
        code +
        " 开头的工作表，当前找到 " +
        names.length +
        " 张。",
    );
  const name = names[0],
    sheet = book.Sheets[name];
  const rows = [];
  for (const [address, cell] of Object.entries(sheet)) {
    if (address.startsWith("!") || cell.v == null) continue;
    const { r, c } = XLSX.utils.decode_cell(address);
    (rows[r] ??= [])[c] = cell.v;
  }
  return { name, sheet, rows };
}

function readTable(tab, fields, handler, empty = false, optional = {}) {
  const start = tab.rows.findIndex(populated);
  if (start < 0 && empty) return [];
  const names = (tab.rows[start] || []).map(cleanHeader);
  const cols = {};
  for (const [field, aliases] of Object.entries(fields)) {
    cols[field] = names.findIndex((n) => aliases.includes(n));
    if (cols[field] < 0)
      throw new Error(tab.name + " 缺少「" + aliases[0] + "」列。");
  }
  for (const [field, aliases] of Object.entries(optional)) {
    const col = names.findIndex((n) => aliases.includes(n));
    if (col >= 0) cols[field] = col;
  }
  tab.headerRow = start;
  tab.idCol = cols.id;
  return tab.rows.flatMap((values, i) => {
    if (i <= start || !populated(values)) return [];
    const obj = Object.fromEntries(
      Object.entries(cols).map(([k, j]) => [k, values[j]]),
    );
    if (cols.id != null) {
      const cell = tab.sheet[XLSX.utils.encode_cell({ r: i, c: cols.id })];
      obj.id = text(
        typeof cell?.v === "number" && /^0+$/.test(cell.z || "")
          ? XLSX.utils.format_cell(cell)
          : cell?.v,
      );
      if (!obj.id)
        throw new Error(tab.name + " 第 " + (i + 1) + " 行缺少工号。");
    }
    return [handler(obj, i + 1)];
  });
}

function readRules(tab) {
  const start = tab.rows.findIndex((row) =>
    row?.some((v) => ["证书", "证书名称"].includes(cleanHeader(v))),
  );
  if (start < 0) throw new Error("03 缺少证书规则表头。");
  const titleCol = tab.rows[start].findIndex((v) =>
    ["证书", "证书名称"].includes(cleanHeader(v)),
  );
  const groups = [[], [], [], []],
    labels = ["一", "二", "三", "四"];
  let active = -1;
  const subhead = (tab.rows[start + 1] || []).some((v) =>
    /^方式/.test(normalize(v)),
  );
  const count = Math.max(
    tab.rows[start].length,
    subhead ? tab.rows[start + 1].length : 0,
  );
  for (let col = 0; col < count; col++) {
    const h = normalize(tab.rows[start][col]),
      match = /^科目([一二三四1-4])/.exec(h);
    if (match)
      active = labels.includes(match[1])
        ? labels.indexOf(match[1])
        : Number(match[1]) - 1;
    else if (!(subhead && /^方式/.test(normalize(tab.rows[start + 1][col]))))
      active = -1;
    if (active >= 0) groups[active].push(col);
  }
  if (groups.some((g) => !g.length))
    throw new Error("03 的表头需包含科目一至科目四。");
  const rules = [],
    seen = new Set();
  tab.rows.forEach((row, i) => {
    if (i <= start + Number(subhead) || !populated(row)) return;
    const loc = "03 第 " + (i + 1) + " 行",
      cert = parseCertificate(row[titleCol], loc);
    const subjects = groups.map((g) =>
      g.map((j) => normalize(row[j])).filter((v) => v && v !== "无"),
    );
    if (!subjects.some((g) => g.length))
      throw new Error(loc + " 缺少考试名称。");
    const key = JSON.stringify(cert);
    if (seen.has(key)) throw new Error(loc + " 的证书规则重复。");
    seen.add(key);
    rules.push({ ...cert, subjects });
  });
  if (!rules.length) throw new Error("03 没有可用的证书规则。");
  return rules;
}

export function openWorkbook(bytes) {
  const book = XLSX.read(bytes, { type: "array", cellNF: true });
  const tabs = Object.fromEntries(
    sheetCodes.map((code) => [code, tab(book, code)]),
  );
  const epoch = Boolean(book.Workbook?.WBProps?.date1904);
  const person = {
    id: ["工号", "学员工号"],
    name: ["姓名", "学员姓名"],
    dept: ["0层部门"],
    team: ["一层部门"],
  };
  function personRow(raw, row, time) {
    return {
      id: raw.id,
      name: text(raw.name),
      dept: text(raw.dept),
      team: text(raw.team),
      title: text(raw.title),
      row,
      time,
      date: time.slice(0, 10),
    };
  }
  const history = readTable(
    tabs["01"],
    { ...person, title: ["证书名称"], date: ["证书获得日期"] },
    (raw, row) => {
      const loc = "01 第 " + row + " 行";
      return {
        ...personRow(raw, row, stamp(raw.date, epoch, loc)),
        ...parseCertificate(raw.title, loc),
      };
    },
    true,
  );
  const exams = readTable(
    tabs["02"],
    {
      ...person,
      title: ["考试名称"],
      date: ["开始答题时间"],
      status: ["状态"],
    },
    (raw, row) => {
      const passed = text(raw.status) === "通过";
      const time =
        !passed && !raw.date
          ? ""
          : stamp(raw.date, epoch, "02 第 " + row + " 行");
      return { ...personRow(raw, row, time), passed };
    },
    true,
  );
  const baseline = readTable(
    tabs["00"],
    { id: ["工号"], title: ["证书名称"], date: ["证书获得日期"] },
    (raw, row) => {
      const location = "00 第 " + row + " 行";
      return {
        id: raw.id,
        name: text(raw.name),
        row,
        ...parseCertificate(raw.title, location),
        date: stamp(raw.date, epoch, location).slice(0, 10),
        status: text(raw.status),
        generated: raw.generated ? stamp(raw.generated, epoch, location) : "",
      };
    },
    true,
    { name: person.name, status: ["证书发放状态"], generated: ["生成时间"] },
  );
  if (new Set(baseline.map(certKey)).size !== baseline.length)
    throw new Error("00 中有重复的工号、级别和语言组合，请先合并重复记录。");
  return {
    bytes,
    book,
    tabs,
    epoch,
    baseline,
    history,
    exams,
    rules: readRules(tabs["03"]),
  };
}

const xmlEscape = (value) =>
  String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
const xml = (files, path) =>
  new DOMParser().parseFromString(strFromU8(files[path]), "application/xml");
const tags = (doc, local) => Array.from(doc.getElementsByTagNameNS("*", local));
const resolvePart = (base, target) =>
  decodeURIComponent(
    new URL(target, "https://workbook.local/" + base).pathname.slice(1),
  );
const serial = (value, epoch) =>
  Date.parse(
    value.replace(" ", "T") + (value.length === 10 ? "T00:00:00Z" : "Z"),
  ) /
    86400000 +
  (epoch ? 24107 : 25569);

function dateStyles(files, stylePath) {
  if (!stylePath || !files[stylePath]) return [];
  const doc = xml(files, stylePath),
    ns = doc.documentElement.namespaceURI;
  const xfs = tags(doc, "cellXfs")[0];
  if (!xfs) return [];
  let formats = tags(doc, "numFmts")[0];
  if (!formats) {
    formats = doc.createElementNS(ns, "numFmts");
    doc.documentElement.insertBefore(formats, doc.documentElement.firstChild);
  }
  const indices = ["yyyy-mm-dd", "yyyy-mm-dd hh:mm:ss"].map((code) => {
    const available = tags(formats, "numFmt");
    let numFmt = available.find(
      (node) => node.getAttribute("formatCode") === code,
    );
    if (!numFmt) {
      numFmt = doc.createElementNS(ns, "numFmt");
      numFmt.setAttribute(
        "numFmtId",
        String(
          1 +
            Math.max(163, ...available.map((n) => +n.getAttribute("numFmtId"))),
        ),
      );
      numFmt.setAttribute("formatCode", code);
      formats.appendChild(numFmt);
      formats.setAttribute("count", String(available.length + 1));
    }
    const items = tags(xfs, "xf"),
      id = numFmt.getAttribute("numFmtId");
    const existing = items.findIndex((i) => i.getAttribute("numFmtId") === id);
    if (existing >= 0) return existing;
    const item = items[0].cloneNode(true);
    item.setAttribute("numFmtId", id);
    item.setAttribute("applyNumberFormat", "1");
    xfs.appendChild(item);
    xfs.setAttribute("count", String(items.length + 1));
    return items.length;
  });
  files[stylePath] = strToU8(new XMLSerializer().serializeToString(doc));
  return indices;
}

// Preserve the original OOXML parts, including extra sheets, formulas, drawings and formatting.
export function exportFull(data, records) {
  const files = unzipSync(new Uint8Array(data.bytes));
  const office = tags(xml(files, "_rels/.rels"), "Relationship").find((n) =>
    n.getAttribute("Type").endsWith("/officeDocument"),
  );
  const bookPath = resolvePart("", office.getAttribute("Target"));
  const relPath = bookPath.replace(/([^/]+)$/, "_rels/$1.rels");
  const relations = tags(xml(files, relPath), "Relationship");
  const sheet = tags(xml(files, bookPath), "sheet").find(
    (n) => n.getAttribute("name") === data.tabs["00"].name,
  );
  const rid = Array.from(sheet.attributes).find(
    (a) => a.localName === "id",
  ).value;
  const target = resolvePart(
    bookPath,
    relations.find((r) => r.getAttribute("Id") === rid).getAttribute("Target"),
  );
  const style = relations.find((r) =>
    r.getAttribute("Type").endsWith("/styles"),
  );
  const dateIndices = dateStyles(
    files,
    style && resolvePart(bookPath, style.getAttribute("Target")),
  );
  const sheetDoc = xml(files, target),
    ns = sheetDoc.documentElement.namespaceURI;
  const oldCols = tags(sheetDoc, "col");
  const widths = [13, 14, 48, 12, 17, 21, 17, 16, 22, 24, 18, 24];
  const cols = widths
    .map(
      (w, i) =>
        '<col min="' +
        (i + 1) +
        '" max="' +
        (i + 1) +
        '" width="' +
        (Number(
          oldCols
            .find(
              (c) =>
                +c.getAttribute("min") <= i + 1 &&
                +c.getAttribute("max") >= i + 1,
            )
            ?.getAttribute("width"),
        ) || w) +
        '" customWidth="1"/>',
    )
    .join("");
  const rows = [headers, ...records.map(cellsOf)].map((values, i) => {
    const cells = values.map((v, j) => {
      if (v === "") return "";
      const ref = XLSX.utils.encode_cell({ r: i, c: j });
      if (i && dateIndices.length && (j === 6 || j === 11))
        return (
          '<c r="' +
          ref +
          '" s="' +
          dateIndices[Number(j === 11)] +
          '"><v>' +
          serial(v, data.epoch) +
          "</v></c>"
        );
      return (
        '<c r="' +
        ref +
        '" t="inlineStr"><is><t xml:space="preserve">' +
        xmlEscape(v) +
        "</t></is></c>"
      );
    });
    return '<row r="' + (i + 1) + '">' + cells.join("") + "</row>";
  });
  const end = records.length + 1;
  files[target] = strToU8(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="' +
      ns +
      '"><dimension ref="A1:L' +
      end +
      '"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>' +
      cols +
      "</cols><sheetData>" +
      rows.join("") +
      '</sheetData><autoFilter ref="A1:L' +
      end +
      '"/></worksheet>',
  );
  return zipSync(files, { level: 1 });
}

export function departmentRows(data, result, key) {
  const permitted = new Set(
    [...result.departments]
      .filter(([, r]) => departmentKey(r.dept, r.team) === key)
      .map(([id]) => id),
  );
  return {
    "00": result.records.filter((r) => permitted.has(r.id)),
    "01": data.history.filter((r) => permitted.has(r.id)),
    "02": data.exams.filter((r) => permitted.has(r.id)),
  };
}

export function exportDepartment(data, result, key, codes) {
  const book = XLSX.utils.book_new();
  const groups = departmentRows(data, result, key);
  for (const code of ["00", "01", "02"].filter((c) => codes.includes(c))) {
    const tab = data.tabs[code];
    let sheet;
    if (code === "00") {
      const rows = groups[code].map(cellsOf);
      sheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
      rows.forEach((row, i) => {
        for (const col of [6, 11]) {
          const cell = sheet[XLSX.utils.encode_cell({ r: i + 1, c: col })];
          cell.t = "n";
          cell.v = serial(row[col], false);
          cell.z = col === 6 ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm:ss";
        }
      });
    } else {
      const indexes = [
        ...(tab.headerRow !== undefined ? [tab.headerRow] : []),
        ...groups[code].map((r) => r.row - 1),
      ];
      sheet = XLSX.utils.aoa_to_sheet(indexes.map((r) => tab.rows[r]));
      // Copy displayed values and number formats only, so the extract is self-contained.
      indexes.forEach((r, i) =>
        tab.rows[r].forEach((_, c) => {
          const src = tab.sheet[XLSX.utils.encode_cell({ r, c })],
            dest = sheet[XLSX.utils.encode_cell({ r: i, c })];
          if (src?.z && dest) {
            dest.z = src.z;
            if (data.epoch && dest.t === "n" && XLSX.SSF.is_date(src.z))
              dest.v += 1462;
          }
        }),
      );
    }
    sheet["!cols"] = (
      code === "00" ? headers : tab.rows[tab.headerRow] || []
    ).map((v, i) => ({ wch: i === 2 || /名称/.test(v) ? 48 : 22 }));
    XLSX.utils.book_append_sheet(book, sheet, tab.name);
  }
  return XLSX.write(book, {
    type: "array",
    bookType: "xlsx",
    compression: true,
  });
}
