import { cellsOf, headers } from "./engine.js";

// Lead with the fields used to identify, review and issue a certificate.
const columnOrder = [0, 1, 3, 4, 6, 10, 8, 9, 5, 2, 7, 11];
const collator = new Intl.Collator("zh-CN", { numeric: true, sensitivity: "base" });

export function resultTable(records) {
  return {
    columns: columnOrder.map((index) => headers[index]),
    rows: records.map((record, index) => {
      const cells = cellsOf(record);
      cells[10] ||= "待标记";
      return {
        line: index + 2,
        values: columnOrder.map((index) => cells[index]),
        order: [undefined, undefined, record.rank],
      };
    }),
  };
}

export function selectRows(rows, { query, filters, sort }) {
  const term = query.trim().toLowerCase();
  const conditions = [...filters].filter(([, value]) => value.trim())
    .map(([column, value]) => [column, value.trim().toLowerCase()]);
  const matches = rows.filter(({ line, values }) =>
    (!term || String(line).includes(term) ||
      values.some((value) => value.toLowerCase().includes(term))) &&
    conditions.every(([column, value]) => values[column].toLowerCase().includes(value)),
  );
  if (!sort) return matches;
  const { column, direction } = sort;
  const valueAt = (row) => column === -1 ? String(row.line) : row.values[column];
  const numeric = column === -1 || matches.every(
    (row) => !valueAt(row) || typeof row.order[column] === "number",
  );
  return matches.sort((a, b) => {
    const first = valueAt(a), second = valueAt(b);
    if (!first || !second) return Number(!first) - Number(!second);
    const compare = numeric && column >= 0
      ? a.order[column] - b.order[column]
      : collator.compare(first, second);
    return direction * compare;
  });
}
