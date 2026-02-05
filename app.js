(function () {
  const formWrap = document.getElementById("formWrap");
  const previewDoc = document.getElementById("previewDoc");
  const chatInput = document.getElementById("chatInput");
  const chatSend = document.getElementById("chatSend");
  const btnDownload = document.getElementById("btnDownload");
  const btnSave = document.getElementById("btnSave");
  const btnLoadSaved = document.getElementById("btnLoadSaved");
  const btnLoadDocx = document.getElementById("btnLoadDocx");
  const fileDocx = document.getElementById("fileDocx");

  let fieldsConfig = {};
  let data = {};

  const LABEL_TO_KEY = {
    "영문 제목": "title_en", "한글 제목": "title_ko",
    "표지 이름": "author", "한글 이름 (개요)": "author_ko", "영문 인용 이름": "author_bib_en",
    "제출일 (Month Year)": "date",
    "영문 초록": "abstract_en", "한글 초록": "abstract_ko",
    "헌정": "dedication", "감사의 말": "acknowledgements",
    "제목": "title_ko", "한글제목": "title_ko", "영문제목": "title_en",
    "이름": "author", "날짜": "date", "초록": "abstract_ko", "영문초록": "abstract_en",
  };

  function parseChatLine(line) {
    const t = line.trim();
    const colon = t.indexOf(":");
    if (colon <= 0) return null;
    const keyPart = t.slice(0, colon).trim();
    const value = t.slice(colon + 1).trim();
    const key = LABEL_TO_KEY[keyPart] || null;
    if (!key) return null;
    return { key, value };
  }

  function escapeHtml(s) {
    const div = document.createElement("div");
    div.textContent = s;
    return div.innerHTML;
  }

  function isLongField(key) {
    return /^(abstract_|dedication|acknowledgements|ch\d)/.test(key);
  }

  function renderForm() {
    const byGroup = {};
    for (const [key, config] of Object.entries(fieldsConfig)) {
      const g = config.group || "기타";
      if (!byGroup[g]) byGroup[g] = [];
      byGroup[g].push({ key, config });
    }
    const order = ["논문 제목", "저자 정보", "제출일", "초록", "헌정·감사", "제1장 서론", "제2장 선행연구", "제3장 연구 방법론", "제4장 본론", "제5장 본론", "제6장 본론", "제7장 결론", "기타"];
    let html = "";
    order.forEach(function (group) {
      const list = byGroup[group];
      if (!list || list.length === 0) return;
      html += '<div class="form-section"><div class="form-section-title">' + escapeHtml(group) + '</div>';
      list.forEach(function (_) {
        const key = _.key;
        const config = _.config;
        const label = config.label || key;
        const val = data[key] == null ? "" : String(data[key]);
        if (isLongField(key)) {
          html += '<div class="form-group"><label>' + escapeHtml(label) + '</label><textarea data-key="' + escapeHtml(key) + '" placeholder="' + escapeHtml(label) + '">' + escapeHtml(val) + '</textarea></div>';
        } else {
          html += '<div class="form-group"><label>' + escapeHtml(label) + '</label><input type="text" data-key="' + escapeHtml(key) + '" value="' + escapeHtml(val) + '" placeholder="' + escapeHtml(label) + '" /></div>';
        }
      });
      html += "</div>";
    });
    formWrap.innerHTML = html;
    formWrap.querySelectorAll("input, textarea").forEach(function (el) {
      el.addEventListener("input", function () {
        data[el.dataset.key] = el.value;
        renderPreview();
      });
    });
  }

  function renderPreview() {
    const sections = [
      { title: "논문 제목", keys: ["title_en", "title_ko"], singleLine: true },
      { title: "저자 · 제출일", keys: ["author", "author_ko", "author_bib_en", "date"], singleLine: true },
      { title: "헌정", keys: ["dedication"] },
      { title: "개요 (ABSTRACT)", keys: ["abstract_ko"] },
      { title: "ABSTRACT", keys: ["abstract_en"] },
      { title: "감사의 말", keys: ["acknowledgements"] },
      { title: "제1장 서론", keys: ["ch1_background", "ch1_thesis", "ch1_purpose", "ch1_goals", "ch1_significance", "ch1_central", "ch1_questions", "ch1_delimitations", "ch1_definitions", "ch1_assumptions", "ch1_methodology", "ch1_procedures", "ch1_summary"] },
      { title: "제2장 선행연구", keys: ["ch2_intro", "ch2_1", "ch2_1_1", "ch2_summary"] },
      { title: "제3장 연구 방법론", keys: ["ch3_intro", "ch3_summary"] },
      { title: "제4장 본론", keys: ["ch4_summary"] },
      { title: "제5장 본론", keys: ["ch5_summary"] },
      { title: "제6장 본론", keys: ["ch6_summary"] },
      { title: "제7장 결론", keys: ["ch7_intro", "ch7_summary", "ch7_conclusion", "ch7_recommendations"] },
    ];
    let html = "";
    sections.forEach(function (sec) {
      let hasAny = false;
      let block = '<div class="doc-section"><div class="doc-section-title">' + escapeHtml(sec.title) + '</div>';
      sec.keys.forEach(function (k) {
        const v = data[k];
        if (v == null || v === "") return;
        hasAny = true;
        const label = (fieldsConfig[k] && fieldsConfig[k].label) ? fieldsConfig[k].label + ": " : "";
        if (sec.singleLine) {
          block += '<div class="doc-meta">' + escapeHtml(label) + escapeHtml(v) + '</div>';
        } else {
          block += '<p>' + escapeHtml(v) + '</p>';
        }
      });
      block += "</div>";
      if (hasAny) html += block;
    });
    if (!html) html = '<p class="doc-placeholder">왼쪽에서 항목을 입력하면 여기에 미리보기가 표시됩니다.</p>';
    previewDoc.innerHTML = html;
  }

  function applyChatInput() {
    const line = chatInput.value.trim();
    if (!line) return;
    const parsed = parseChatLine(line);
    if (parsed) {
      data[parsed.key] = parsed.value;
      renderForm();
      renderPreview();
      chatInput.value = "";
    }
  }

  chatSend.addEventListener("click", applyChatInput);
  chatInput.addEventListener("keydown", function (e) {
    if (e.key === "Enter") {
      e.preventDefault();
      applyChatInput();
    }
  });

  btnSave.addEventListener("click", function () {
    try {
      localStorage.setItem("oikos-thesis-data", JSON.stringify(data));
      alert("저장되었습니다.");
    } catch (e) {
      alert("저장 실패: " + (e.message || e));
    }
  });

  btnLoadSaved.addEventListener("click", function () {
    const saved = localStorage.getItem("oikos-thesis-data");
    if (!saved) {
      alert("저장된 내용이 없습니다.");
      return;
    }
    try {
      data = JSON.parse(saved);
      renderForm();
      renderPreview();
      alert("저장된 내용을 불러왔습니다.");
    } catch (e) {
      alert("불러오기 실패: " + (e.message || e));
    }
  });

  btnLoadDocx.addEventListener("click", function () {
    fileDocx.click();
  });

  fileDocx.addEventListener("change", function () {
    const file = fileDocx.files[0];
    if (!file) return;
    const fd = new FormData();
    fd.append("docx", file);
    btnLoadDocx.disabled = true;
    btnLoadDocx.textContent = "불러오는 중...";
    fetch("/api/load-docx", { method: "POST", body: fd })
      .then(function (res) {
        if (!res.ok) return res.json().then(function (j) { throw new Error(j.error || res.status); });
        return res.json();
      })
      .then(function (loaded) {
        data = loaded;
        renderForm();
        renderPreview();
        alert("DOCX 내용을 불러왔습니다.");
      })
      .catch(function (err) {
        alert("불러오기 실패: " + err.message);
      })
      .finally(function () {
        btnLoadDocx.disabled = false;
        btnLoadDocx.textContent = "DOCX에서 불러오기";
        fileDocx.value = "";
      });
  });

  btnDownload.addEventListener("click", function () {
    btnDownload.disabled = true;
    btnDownload.textContent = "생성 중...";
    fetch("/api/fill", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data),
    })
      .then(function (res) {
        if (!res.ok) return res.json().then(function (j) { throw new Error(j.error || res.status); });
        return res.blob();
      })
      .then(function (blob) {
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "filled.docx";
        a.click();
        URL.revokeObjectURL(a.href);
      })
      .catch(function (err) {
        alert("다운로드 실패: " + err.message);
      })
      .finally(function () {
        btnDownload.disabled = false;
        btnDownload.textContent = "DOCX 다운로드";
      });
  });

  fetch("/api/fields")
    .then(function (r) { return r.json(); })
    .then(function (fields) {
      fieldsConfig = fields;
      const saved = localStorage.getItem("oikos-thesis-data");
      if (saved) {
        try {
          data = JSON.parse(saved);
        } catch (e) {}
      }
      renderForm();
      renderPreview();
    })
    .catch(function () {
      formWrap.innerHTML = "<p>필드 목록을 불러올 수 없습니다. 서버가 실행 중인지 확인하세요.</p>";
    });
})();
