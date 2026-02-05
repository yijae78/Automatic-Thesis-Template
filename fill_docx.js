/**
 * CLI: node fill_docx.js [data.json] [output.docx]
 */
const fs = require("fs");
const path = require("path");
const { fillDocx } = require("./fillModule");

const base = __dirname;
const dataPath = path.join(base, process.argv[2] || "data.json");
const outPath = path.join(base, process.argv[3] || "filled.docx");

if (!fs.existsSync(dataPath)) {
  console.error("데이터 파일 없음:", dataPath);
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(dataPath, "utf8"));
const buffer = fillDocx(data, base);
fs.writeFileSync(outPath, buffer);
console.log("채우기 완료 →", outPath);
