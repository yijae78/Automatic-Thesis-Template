/**
 * 템플릿 DOCX 채우기 핵심 로직. data 객체를 받아 치환된 DOCX Buffer 반환.
 */
const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");
const { DOMParser, XMLSerializer } = require("xmldom");

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

function setParagraphText(p, newText) {
  const runs = p.getElementsByTagNameNS(NS, "r");
  if (runs.length === 0) return;
  const firstRun = runs[0];
  let tNodes = firstRun.getElementsByTagNameNS(NS, "t");
  if (tNodes.length === 0) {
    const t = firstRun.ownerDocument.createElementNS(NS, "t");
    t.setAttribute("xml:space", "preserve");
    firstRun.appendChild(t);
    tNodes = firstRun.getElementsByTagNameNS(NS, "t");
  }
  tNodes[0].textContent = newText;
  for (let i = 1; i < tNodes.length; i++) tNodes[i].textContent = "";
  for (let i = 1; i < runs.length; i++) {
    const r = runs[i];
    const ts = r.getElementsByTagNameNS(NS, "t");
    for (let j = 0; j < ts.length; j++) ts[j].textContent = "";
  }
}

function fillDocx(data, baseDir) {
  baseDir = baseDir || __dirname;
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

  for (const [fieldId, config] of Object.entries(fields)) {
    const value = data[fieldId];
    if (value == null || value === "") continue;
    const styleName = config.style;
    const matchText = config.matchText;

    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const styleId = getParagraphStyleId(p);
      const resolvedStyle = styleIdToName[styleId] || styleId || "(없음)";
      const text = getParagraphText(p);
      if (resolvedStyle === styleName && (text === matchText || text.indexOf(matchText) >= 0)) {
        setParagraphText(p, value);
        break;
      }
    }
  }

  const serializer = new XMLSerializer();
  zip.updateFile("word/document.xml", Buffer.from(serializer.serializeToString(docDoc), "utf8"));
  return zip.toBuffer();
}

module.exports = { fillDocx };
