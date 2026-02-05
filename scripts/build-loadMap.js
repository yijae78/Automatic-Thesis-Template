/**
 * 템플릿 DOCX + fields.json으로 loadMap.json 생성.
 * 각 필드에 대해 (style, matchText)에 처음 매칭되는 문단의 스타일 내 순서(index)를 기록.
 */
const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");
const { DOMParser } = require("xmldom");

const NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const baseDir = path.join(__dirname, "..");
const templatePath = path.join(baseDir, "오이코스논문 신탬플릿---수정본.docx");
const fieldsPath = path.join(baseDir, "fields.json");
const outPath = path.join(baseDir, "loadMap.json");

function getParagraphText(p) {
  const list = p.getElementsByTagNameNS(NS, "t");
  let s = "";
  for (let i = 0; i < list.length; i++) {
    if (list[i].textContent) s += list[i].textContent;
  }
  return s.trim();
}

function getParagraphStyleId(p) {
  const pPr = p.getElementsByTagNameNS(NS, "pPr")[0];
  if (!pPr) return null;
  const pStyle = pPr.getElementsByTagNameNS(NS, "pStyle")[0];
  if (!pStyle) return null;
  return pStyle.getAttribute("w:val") || pStyle.getAttribute("val");
}

const fields = JSON.parse(fs.readFileSync(fieldsPath, "utf8"));
const zip = new AdmZip(templatePath);

let styleIdToName = {};
const styleEntry = zip.getEntry("word/styles.xml");
if (styleEntry) {
  const styleDoc = new DOMParser().parseFromString(styleEntry.getData().toString("utf8"));
  const styles = styleDoc.getElementsByTagNameNS(NS, "style");
  for (let i = 0; i < styles.length; i++) {
    const sid = styles[i].getAttribute("w:styleId") || styles[i].getAttribute("styleId");
    const nameEl = styles[i].getElementsByTagNameNS(NS, "name")[0];
    if (sid && nameEl) {
      const name = nameEl.getAttribute("w:val") || nameEl.getAttribute("val");
      if (name) styleIdToName[sid] = name;
    }
  }
}

const docEntry = zip.getEntry("word/document.xml");
const docDoc = new DOMParser().parseFromString(docEntry.getData().toString("utf8"));
const body = docDoc.getElementsByTagNameNS(NS, "body")[0] || docDoc.documentElement;
const paragraphs = body.getElementsByTagNameNS(NS, "p");

const styleCount = {};
const paraList = [];
for (let i = 0; i < paragraphs.length; i++) {
  const p = paragraphs[i];
  const styleId = getParagraphStyleId(p);
  const styleName = styleIdToName[styleId] || styleId || "(없음)";
  const text = getParagraphText(p);
  if (!styleCount[styleName]) styleCount[styleName] = 0;
  const indexInStyle = styleCount[styleName];
  styleCount[styleName]++;
  paraList.push({ styleName, text, indexInStyle });
}

const used = new Set();
const loadMap = {};
for (const [fieldId, config] of Object.entries(fields)) {
  const styleName = config.style;
  const matchText = config.matchText;
  let found = false;
  for (let i = 0; i < paraList.length; i++) {
    if (used.has(i)) continue;
    const para = paraList[i];
    if (para.styleName === styleName && (para.text === matchText || para.text.indexOf(matchText) >= 0)) {
      loadMap[fieldId] = { style: styleName, index: para.indexInStyle };
      used.add(i);
      found = true;
      break;
    }
  }
  if (!found) console.warn("loadMap: no match for", fieldId, styleName, matchText);
}

fs.writeFileSync(outPath, JSON.stringify(loadMap, null, 2), "utf8");
console.log("loadMap.json written to", outPath);
