const path = require("path");
const fs = require("fs");
const { exec } = require("child_process");
const express = require("express");
const multer = require("multer");
const mammoth = require("mammoth");
const { fillDocx } = require("./fillModule");
const { extractDataFromDocx } = require("./loadModule");

const app = express();
const root = __dirname;
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

app.use(express.json({ limit: "10mb" }));
app.use(express.static(root));

app.get("/api/fields", (req, res) => {
  const p = path.join(root, "fields.json");
  if (!fs.existsSync(p)) return res.status(404).json({ error: "fields.json 없음" });
  res.json(JSON.parse(fs.readFileSync(p, "utf8")));
});

app.post("/api/fill", (req, res) => {
  try {
    const data = req.body || {};
    const buffer = fillDocx(data, root);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", 'attachment; filename="filled.docx"');
    res.send(buffer);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: String(e.message) });
  }
});

app.post("/api/load-docx", upload.single("docx"), (req, res) => {
  try {
    if (!req.file || !req.file.buffer) return res.status(400).json({ error: "DOCX 파일이 없습니다." });
    const data = extractDataFromDocx(req.file.buffer, root);
    res.json(data);
  } catch (e) {
    console.error(e);
    const msg = e.message && e.message.indexOf("템플릿") !== -1 ? e.message : "DOCX 내용을 읽는 중 오류가 발생했습니다. 템플릿과 같은 양식의 파일인지 확인해 주세요.";
    res.status(500).json({ error: msg });
  }
});

app.get("/api/template-preview-html", (req, res) => {
  const templatePath = path.join(root, "오이코스논문 신탬플릿---수정본.docx");
  if (!fs.existsSync(templatePath)) return res.status(404).json({ error: "템플릿 DOCX 없음" });
  mammoth
    .convertToHtml({ path: templatePath })
    .then((result) => {
      res.setHeader("Content-Type", "text/html; charset=utf-8");
      res.send(result.value);
    })
    .catch((err) => {
      console.error(err);
      res.status(500).json({ error: String(err.message) });
    });
});

const port = process.env.PORT || 3750;
app.listen(port, () => {
  const url = "http://localhost:" + port;
  console.log("오이코스 논문 템플릿 서버:");
  console.log("  " + url);
  console.log("  http://127.0.0.1:" + port);
  // 서버 시작 시 Edge에서 자동으로 열기 (Windows). 비활성화: OPEN_EDGE=0
  if (process.platform === "win32" && process.env.OPEN_EDGE !== "0") {
    exec('start msedge "' + url + '"', (err) => {
      if (err) console.error("Edge 실행 실패:", err.message);
    });
  }
});
