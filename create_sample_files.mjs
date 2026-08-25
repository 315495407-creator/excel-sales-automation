import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const projectDir = new URL(".", import.meta.url).pathname.replace(/^\//, "");
const inputDir = `${projectDir}input`;
await fs.mkdir(inputDir, { recursive: true });

const files = [
  {
    name: "sales_january.xlsx",
    rows: [
      [new Date("2026-01-05"), "小王", "键盘", 2, 398],
      [new Date("2026-01-08"), "小李", "鼠标", 3, 297],
      [new Date("2026-01-12"), "小王", "显示器", 1, 1299],
    ],
  },
  {
    name: "sales_february.xlsx",
    rows: [
      [new Date("2026-02-03"), "小陈", "键盘", 1, 199],
      [new Date("2026-02-14"), "小李", "耳机", 2, 598],
      [new Date("2026-02-21"), "小陈", "鼠标", 4, 396],
    ],
  },
];

for (const file of files) {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("销售数据");
  sheet.getRange("A1:E4").values = [
    ["日期", "销售员", "产品", "数量", "金额"],
    ...file.rows,
  ];
  sheet.getRange("A1:E1").format = {
    fill: "#2563EB",
    font: { bold: true, color: "#FFFFFF" },
  };
  sheet.getRange("A2:A4").format.numberFormat = "yyyy-mm-dd";
  sheet.getRange("D2:D4").format.numberFormat = "#,##0";
  sheet.getRange("E2:E4").format.numberFormat = "#,##0.00";
  sheet.getRange("A1:E4").format.borders = { preset: "all", style: "thin", color: "#D9E2F3" };
  sheet.getRange("A1:E4").format.autofitColumns();
  sheet.freezePanes.freezeRows(1);
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(`${inputDir}/${file.name}`);
}

console.log(`已生成 ${files.length} 个练习文件：${inputDir}`);
