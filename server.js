const path = require("path");
const fs = require("fs");
const express = require("express");
const multer = require("multer");
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
    res.status(500).json({ error: String(e.message) });
  }
});

const port = process.env.PORT || 3750;
app.listen(port, () => {
  console.log("오이코스 논문 템플릿 서버: http://localhost:" + port);
});
