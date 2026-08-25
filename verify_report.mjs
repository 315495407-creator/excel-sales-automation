import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const input = await FileBlob.load("output/merged_sales_report.xlsx");
const workbook = await SpreadsheetFile.importXlsx(input);

const check = await workbook.inspect({
  kind: "table",
  range: "合并明细!A1:F10",
  include: "values,formulas",
  tableMaxRows: 10,
  tableMaxCols: 8,
});
console.log(check.ndjson);

const summary = await workbook.inspect({
  kind: "table",
  range: "汇总!A1:B5",
  include: "values,formulas",
  tableMaxRows: 5,
  tableMaxCols: 4,
});
console.log(summary.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

const preview = await workbook.render({ sheetName: "汇总", autoCrop: "all", scale: 1, format: "png" });
await fs.writeFile("output/summary_preview.png", new Uint8Array(await preview.arrayBuffer()));
