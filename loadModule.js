/**
 * DOCX에서 필드별 텍스트 추출. loadMap(style + index)으로 위치 식별.
 */
const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");
const { DOMParser } = require("xmldom");

const NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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

function buildLoadMapFromTemplate(baseDir) {
  const templatePath = path.join(baseDir, "오이코스논문 신탬플릿---수정본.docx");
  const fieldsPath = path.join(baseDir, "fields.json");
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
    for (let i = 0; i < paraList.length; i++) {
      if (used.has(i)) continue;
      const para = paraList[i];
      if (para.styleName === styleName && (para.text === matchText || para.text.indexOf(matchText) >= 0)) {
        loadMap[fieldId] = { style: styleName, index: para.indexInStyle };
        used.add(i);
        break;
      }
    }
  }
  return loadMap;
}

function extractDataFromDocx(buffer, baseDir) {
  baseDir = baseDir || __dirname;
  const loadMapPath = path.join(baseDir, "loadMap.json");
  let loadMap;
  if (fs.existsSync(loadMapPath)) {
    loadMap = JSON.parse(fs.readFileSync(loadMapPath, "utf8"));
  } else {
    loadMap = buildLoadMapFromTemplate(baseDir);
    fs.writeFileSync(loadMapPath, JSON.stringify(loadMap, null, 2), "utf8");
  }

  let zip;
  try {
    zip = new AdmZip(buffer);
  } catch (e) {
    throw new Error("지원하지 않는 파일이거나 파일이 손상되었을 수 있습니다. 템플릿과 같은 양식의 DOCX인지 확인해 주세요.");
  }

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
  if (!docEntry) throw new Error("DOCX 구조가 올바르지 않습니다. 템플릿과 같은 양식의 파일인지 확인해 주세요.");
  const docDoc = new DOMParser().parseFromString(docEntry.getData().toString("utf8"));
  const body = docDoc.getElementsByTagNameNS(NS, "body")[0] || docDoc.documentElement;
  const paragraphs = body.getElementsByTagNameNS(NS, "p");

  const styleCount = {};
  const paragraphsByStyle = {};
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const styleId = getParagraphStyleId(p);
    const styleName = styleIdToName[styleId] || styleId || "(없음)";
    const text = getParagraphText(p);
    if (!paragraphsByStyle[styleName]) paragraphsByStyle[styleName] = [];
    paragraphsByStyle[styleName].push(text);
  }

  const data = {};
  for (const [fieldId, config] of Object.entries(loadMap)) {
    const style = config.style;
    const index = config.index;
    if (paragraphsByStyle[style] && paragraphsByStyle[style][index] != null) {
      data[fieldId] = paragraphsByStyle[style][index];
    }
  }
  return data;
}

module.exports = { extractDataFromDocx, buildLoadMapFromTemplate };
