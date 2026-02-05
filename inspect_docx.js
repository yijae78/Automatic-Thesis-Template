// Node로 docx 구조 분석 (Python 불필요)
const fs = require("fs");
const path = require("path");

const base = __dirname;
const docxPath = path.join(base, "오이코스논문 신탬플릿---수정본.docx");
const outPath = path.join(base, "template_structure.txt");

if (!fs.existsSync(docxPath)) {
  fs.writeFileSync(path.join(base, "error_log.txt"), "DOCX file not found: " + docxPath, "utf8");
  console.log("Error: DOCX not found. See error_log.txt");
  process.exit(1);
}

// docx = zip. Node에 zip 내장 없음 → adm-zip 사용 (npm i adm-zip)
let AdmZip;
try {
  AdmZip = require("adm-zip");
} catch (e) {
  console.log("먼저 실행: npm install adm-zip");
  process.exit(1);
}

const zip = new AdmZip(docxPath);
const zipEntries = zip.getEntries();

const NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
function getText(el) {
  const list = el.getElementsByTagNameNS(NS, "t");
  let s = "";
  for (let i = 0; i < list.length; i++) {
    if (list[i].textContent) s += list[i].textContent;
  }
  return s.trim();
}
function getStyleId(p) {
  const pPr = p.getElementsByTagNameNS(NS, "pPr")[0];
  if (!pPr) return null;
  const pStyle = pPr.getElementsByTagNameNS(NS, "pStyle")[0];
  if (!pStyle) return null;
  return pStyle.getAttribute("w:val") || pStyle.getAttribute("val");
}

const styleNames = {};
const styleEntry = zipEntries.find((e) => e.entryName === "word/styles.xml");
if (styleEntry) {
  const xml = styleEntry.getData().toString("utf8");
  const styleDoc = new (require("xmldom").DOMParser)().parseFromString(xml);
  const styles = styleDoc.getElementsByTagNameNS(NS, "style");
  for (let i = 0; i < styles.length; i++) {
    const sid = styles[i].getAttribute("w:styleId") || styles[i].getAttribute("styleId");
    const nameEl = styles[i].getElementsByTagNameNS(NS, "name")[0];
    if (sid && nameEl) {
      const name = nameEl.getAttribute("w:val") || nameEl.getAttribute("val");
      if (name) styleNames[sid] = name;
    }
  }
}

const docEntry = zipEntries.find((e) => e.entryName === "word/document.xml");
if (!docEntry) {
  fs.writeFileSync(outPath, "word/document.xml not found in DOCX\n", "utf8");
  console.log("Done. See template_structure.txt");
  process.exit(0);
}

const docXml = docEntry.getData().toString("utf8");
const docDoc = new (require("xmldom").DOMParser)().parseFromString(docXml);
const body = docDoc.getElementsByTagNameNS(NS, "body")[0] || docDoc.documentElement;
const paragraphs = body.getElementsByTagNameNS(NS, "p");

const outLines = [];
for (let i = 0; i < paragraphs.length; i++) {
  const line = getText(paragraphs[i]);
  if (!line) continue;
  const styleId = getStyleId(paragraphs[i]);
  const styleName = styleNames[styleId] || styleId || "(없음)";
  outLines.push([styleName, line]);
}

let out = "=== 스타일 정의 (styles.xml) ===\n";
Object.keys(styleNames)
  .sort()
  .forEach((sid) => {
    out += "  " + sid + " -> " + styleNames[sid] + "\n";
  });
out += "\n=== 문단 목록 (스타일 | 내용) ===\n";
outLines.forEach(([styleName, line]) => {
  if (line.length > 70) line = line.slice(0, 70) + "...";
  out += "  [" + styleName + "] " + line + "\n";
});

fs.writeFileSync(outPath, out, "utf8");
console.log("Done. See template_structure.txt");
