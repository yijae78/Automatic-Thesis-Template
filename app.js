(function () {
  const formWrap = document.getElementById("formWrap");
  const previewDoc = document.getElementById("previewDoc");
  const previewToc = document.getElementById("previewToc");
  const previewSheets = document.getElementById("previewSheets");
  const chatInput = document.getElementById("chatInput");
  const chatSend = document.getElementById("chatSend");
  const btnDownload = document.getElementById("btnDownload");
  const btnSave = document.getElementById("btnSave");
  const btnLoadSaved = document.getElementById("btnLoadSaved");
  const btnResetPreview = document.getElementById("btnResetPreview");
  const btnLoadDocx = document.getElementById("btnLoadDocx");
  const fileDocx = document.getElementById("fileDocx");
  const btnExportJson = document.getElementById("btnExportJson");
  const btnImportJson = document.getElementById("btnImportJson");
  const fileJson = document.getElementById("fileJson");
  const autoSaveStatus = document.getElementById("autoSaveStatus");

  let fieldsConfig = {};
  let sheets = [{ name: "문서 1", data: {} }];
  let currentSheetIndex = 0;
  let data = sheets[0].data;
  let templateHtmlCache = null;
  let tocExpanded = {};
  let formSectionExpanded = {};
  const REQUIRED_KEYS = ["title_en", "title_ko", "author", "date"];

  function getTemplateHtml() {
    if (templateHtmlCache) return Promise.resolve(templateHtmlCache);
    return fetch("/api/template-preview-html")
      .then(function (r) {
        if (!r.ok) throw new Error(r.statusText);
        return r.text();
      })
      .then(function (html) {
        templateHtmlCache = html;
        return html;
      });
  }

  /** fields.json 순서대로 placeholder → fieldId. 동일 문구는 등장 순서대로 1:1 치환. */
  function replacePlaceholdersInHtml(html) {
    const orderedKeys = Object.keys(fieldsConfig);
    let out = html;
    for (let i = 0; i < orderedKeys.length; i++) {
      const key = orderedKeys[i];
      const config = fieldsConfig[key];
      const matchText = config && config.matchText;
      if (!matchText) continue;
      const value = data[key] != null ? String(data[key]) : "";
      const escaped = highlightMarkdownForDisplay(value);
      const span = '<span id="sec-' + escapeHtml(key) + '" class="preview-sec" data-section="true">' + escaped + "</span>";
      const idx = out.indexOf(matchText);
      if (idx !== -1) out = out.slice(0, idx) + span + out.slice(idx + matchText.length);
    }
    return out;
  }

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

  function highlightMarkdownForDisplay(text) {
    if (text == null) text = "";
    var s = escapeHtml(String(text));
    var patterns = [
      { re: /\*\*/g, name: "**" },
      { re: /__/g, name: "__" },
      { re: /\*/g, name: "*" },
      { re: /_/g, name: "_" },
      { re: /`/g, name: "`" },
      { re: /^#{1,6}\s/gm, name: "#" },
      { re: /^-\s/gm, name: "-" },
      { re: /^\d+\.\s/gm, name: "1." }
    ];
    patterns.forEach(function (p) {
      s = s.replace(p.re, "<span class=\"markdown-hint\">$&</span>");
    });
    return s;
  }

  let markdownHintVisible = false;

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
      const isExpanded = formSectionExpanded[group] !== false;
      const toggleChar = isExpanded ? "▼" : "▶";
      html += '<div class="form-section' + (isExpanded ? "" : " collapsed") + '" data-group="' + escapeHtml(group) + '">';
      html += '<button type="button" class="form-section-title" data-toggle-group="' + escapeHtml(group) + '"><span class="form-section-toggle">' + toggleChar + "</span> " + escapeHtml(group) + "</button>";
      html += '<div class="form-section-body">';
      list.forEach(function (_) {
        const key = _.key;
        const config = _.config;
        const label = config.label || key;
        const required = config.required === true || REQUIRED_KEYS.indexOf(key) !== -1;
        const star = required ? ' <span class="required-star">*</span>' : "";
        const val = data[key] == null ? "" : String(data[key]);
        if (isLongField(key)) {
          html += '<div class="form-group form-group--long"><label>' + escapeHtml(label) + star + '</label><div class="form-group-field-row">';
          html += '<textarea data-key="' + escapeHtml(key) + '" placeholder="' + escapeHtml(label) + '">' + escapeHtml(val) + '</textarea>';
          html += '<button type="button" class="btn-clear-field" data-clear-key="' + escapeHtml(key) + '">지우기</button></div>';
          html += '<div class="markdown-hint-wrap" data-key="' + escapeHtml(key) + '" style="display:' + (markdownHintVisible ? "block" : "none") + ';"><div class="markdown-hint-content">' + highlightMarkdownForDisplay(val) + '</div></div></div>';
        } else {
          html += '<div class="form-group"><label>' + escapeHtml(label) + star + '</label><div class="form-group-field-row"><input type="text" data-key="' + escapeHtml(key) + '" value="' + escapeHtml(val) + '" placeholder="' + escapeHtml(label) + '" /><button type="button" class="btn-clear-field" data-clear-key="' + escapeHtml(key) + '">지우기</button></div></div>';
        }
      });
      html += "</div></div>";
    });
    formWrap.innerHTML = html;
    formWrap.querySelectorAll(".form-section-title").forEach(function (btn) {
      btn.addEventListener("click", function () {
        const g = btn.dataset.toggleGroup;
        formSectionExpanded[g] = !formSectionExpanded[g];
        var sec = formWrap.querySelector('.form-section[data-group="' + g + '"]');
        if (sec) sec.classList.toggle("collapsed", !formSectionExpanded[g]);
        btn.querySelector(".form-section-toggle").textContent = formSectionExpanded[g] ? "▼" : "▶";
      });
    });
    formWrap.querySelectorAll("input, textarea").forEach(function (el) {
      el.addEventListener("input", function () {
        data[el.dataset.key] = el.value;
        renderPreview();
        if (isLongField(el.dataset.key)) {
          var wrap = formWrap.querySelector('.markdown-hint-wrap[data-key="' + el.dataset.key + '"]');
          if (wrap) wrap.querySelector(".markdown-hint-content").innerHTML = highlightMarkdownForDisplay(el.value);
        }
      });
    });
    formWrap.querySelectorAll(".btn-clear-field").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var key = btn.dataset.clearKey;
        data[key] = "";
        var input = formWrap.querySelector('input[data-key="' + key + '"], textarea[data-key="' + key + '"]');
        if (input) input.value = "";
        var wrap = formWrap.querySelector('.markdown-hint-wrap[data-key="' + key + '"]');
        if (wrap) wrap.querySelector(".markdown-hint-content").innerHTML = "";
        renderPreview();
      });
    });
  }

  function setMarkdownHintVisible(visible) {
    markdownHintVisible = !!visible;
    formWrap.querySelectorAll(".markdown-hint-wrap").forEach(function (w) {
      w.style.display = markdownHintVisible ? "block" : "none";
    });
    if (document.getElementById("markdownHintToggle")) document.getElementById("markdownHintToggle").checked = markdownHintVisible;
  }

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

  function renderSimplePreview() {
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
          block += '<div id="sec-' + escapeHtml(k) + '" class="doc-meta">' + escapeHtml(label) + highlightMarkdownForDisplay(v) + '</div>';
        } else {
          block += '<p id="sec-' + escapeHtml(k) + '">' + highlightMarkdownForDisplay(v) + '</p>';
        }
      });
      block += "</div>";
      if (hasAny) html += block;
    });
    if (!html) html = '<p class="doc-placeholder">왼쪽에서 항목을 입력하면 여기에 미리보기가 표시됩니다.</p>';
    previewDoc.innerHTML = html;
    if (typeof renderToc === "function") renderToc();
  }

  function scrollToSection(secId) {
    const el = document.getElementById(secId);
    if (el) el.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function renderToc() {
    if (!previewToc || !fieldsConfig) return;
    let html = "";
    sections.forEach(function (sec, idx) {
      const isExpanded = tocExpanded[idx] !== false;
      const toggleChar = isExpanded ? "▼" : "▶";
      html += '<div class="toc-chapter" data-chapter="' + idx + '">';
      html += '<button type="button" class="toc-chapter-head" aria-expanded="' + isExpanded + '" data-toggle="' + idx + '">';
      html += '<span class="toc-toggle" aria-hidden="true">' + toggleChar + "</span>";
      html += "<span>" + escapeHtml(sec.title) + "</span>";
      html += "</button>";
      html += '<div class="toc-children' + (isExpanded ? "" : " collapsed") + '">';
      sec.keys.forEach(function (k) {
        const label = (fieldsConfig[k] && fieldsConfig[k].label) ? fieldsConfig[k].label : k;
        html += '<a class="toc-item" href="#sec-' + escapeHtml(k) + '" data-scroll="sec-' + escapeHtml(k) + '">' + escapeHtml(label) + "</a>";
      });
      html += "</div></div>";
    });
    previewToc.innerHTML = html;
    previewToc.querySelectorAll(".toc-chapter-head").forEach(function (btn) {
      btn.addEventListener("click", function () {
        const idx = Number(btn.dataset.toggle);
        tocExpanded[idx] = !tocExpanded[idx];
        renderToc();
      });
    });
    previewToc.querySelectorAll(".toc-item[data-scroll]").forEach(function (a) {
      a.addEventListener("click", function (e) {
        e.preventDefault();
        scrollToSection(a.dataset.scroll);
      });
    });
  }

  function getUniqueSheetName(baseName) {
    const names = new Set(sheets.map(function (s) { return s.name; }));
    if (!names.has(baseName)) return baseName;
    let n = 2;
    while (names.has(baseName + " (" + n + ")")) n++;
    return baseName + " (" + n + ")";
  }

  function renderSheetsTabs() {
    if (!previewSheets || !sheets.length) return;
    let html = "";
    sheets.forEach(function (sheet, i) {
      const active = i === currentSheetIndex ? " active" : "";
      html += '<div class="sheet-tab' + active + '" data-sheet-index="' + i + '" role="tab" aria-selected="' + (i === currentSheetIndex) + '">';
      html += '<span class="sheet-tab-name" title="' + escapeHtml(sheet.name) + '">' + escapeHtml(sheet.name) + "</span>";
      html += '<button type="button" class="sheet-tab-close" data-close="' + i + '" aria-label="시트 닫기">×</button>';
      html += "</div>";
    });
    previewSheets.innerHTML = html;
    previewSheets.querySelectorAll(".sheet-tab").forEach(function (tab) {
      const idx = Number(tab.dataset.sheetIndex);
      tab.addEventListener("click", function (e) {
        if (e.target.classList.contains("sheet-tab-close")) return;
        currentSheetIndex = idx;
        data = sheets[currentSheetIndex].data;
        renderForm();
        renderPreview();
        renderSheetsTabs();
      });
    });
    previewSheets.querySelectorAll(".sheet-tab-close").forEach(function (btn) {
      btn.addEventListener("click", function (e) {
        e.stopPropagation();
        const idx = Number(btn.dataset.close);
        if (sheets.length <= 1) return;
        sheets.splice(idx, 1);
        if (currentSheetIndex >= sheets.length) currentSheetIndex = sheets.length - 1;
        else if (idx < currentSheetIndex) currentSheetIndex--;
        data = sheets[currentSheetIndex].data;
        renderForm();
        renderPreview();
        renderSheetsTabs();
      });
    });
  }

  function renderPreview() {
    if (!fieldsConfig || Object.keys(fieldsConfig).length === 0) {
      renderSimplePreview();
      return;
    }
    const hasData = Object.keys(data).length > 0 && Object.keys(data).some(function (k) {
      const v = data[k];
      return v != null && String(v).trim() !== "";
    });
    if (!hasData) {
      renderSimplePreview();
      if (typeof renderToc === "function") renderToc();
      return;
    }
    if (templateHtmlCache) {
      previewDoc.innerHTML = replacePlaceholdersInHtml(templateHtmlCache);
      if (typeof renderToc === "function") renderToc();
      return;
    }
    renderSimplePreview();
    getTemplateHtml()
      .then(function (html) {
        previewDoc.innerHTML = replacePlaceholdersInHtml(html);
        if (typeof renderToc === "function") renderToc();
      })
      .catch(function () {
        renderSimplePreview();
      });
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
      localStorage.setItem("oikos-thesis-sheets", JSON.stringify({ sheets: sheets, currentSheetIndex: currentSheetIndex }));
      alert("저장되었습니다.");
    } catch (e) {
      alert("저장 실패: " + (e.message || e));
    }
  });

  btnLoadSaved.addEventListener("click", function () {
    const sheetsSaved = localStorage.getItem("oikos-thesis-sheets");
    if (sheetsSaved) {
      try {
        const parsed = JSON.parse(sheetsSaved);
        if (parsed.sheets && Array.isArray(parsed.sheets) && parsed.sheets.length > 0) {
          sheets = parsed.sheets;
          currentSheetIndex = Math.min(parsed.currentSheetIndex || 0, sheets.length - 1);
          data = sheets[currentSheetIndex].data;
          renderForm();
          renderPreview();
          renderSheetsTabs();
          alert("저장된 내용을 불러왔습니다.");
          return;
        }
      } catch (e) {}
    }
    const saved = localStorage.getItem("oikos-thesis-data");
    if (!saved) {
      alert("저장된 내용이 없습니다.");
      return;
    }
    try {
      const loaded = JSON.parse(saved);
      sheets[currentSheetIndex].data = loaded;
      data = sheets[currentSheetIndex].data;
      renderForm();
      renderPreview();
      alert("저장된 내용을 불러왔습니다.");
    } catch (e) {
      alert("불러오기 실패: " + (e.message || e));
    }
  });

  if (btnExportJson) {
    btnExportJson.addEventListener("click", function () {
      try {
        const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "oikos-thesis-data.json";
        a.click();
        URL.revokeObjectURL(a.href);
      } catch (e) {
        alert("내보내기 실패: " + (e.message || e));
      }
    });
  }
  if (btnImportJson) {
    btnImportJson.addEventListener("click", function () { fileJson.click(); });
  }
  if (fileJson) {
    fileJson.addEventListener("change", function () {
      const file = fileJson.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = function () {
        try {
          const parsed = JSON.parse(reader.result);
          sheets[currentSheetIndex].data = parsed;
          data = sheets[currentSheetIndex].data;
          renderForm();
          renderPreview();
          renderSheetsTabs();
          alert("JSON을 불러왔습니다.");
        } catch (e) {
          alert("JSON 파싱 실패: " + (e.message || e));
        }
        fileJson.value = "";
      };
      reader.readAsText(file, "utf8");
    });
  }

  btnResetPreview.addEventListener("click", function () {
    Object.keys(data).forEach(function (k) { delete data[k]; });
    try {
      localStorage.removeItem("oikos-thesis-data");
    } catch (e) {}
    renderForm();
    renderPreview();
    if (typeof renderToc === "function") renderToc();
    alert("미리보기가 초기화되었습니다.");
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
        if (!res.ok) {
          return res.json().catch(function () { return { error: res.statusText || String(res.status) }; })
            .then(function (j) { throw new Error(j.error || res.statusText || res.status); });
        }
        return res.json();
      })
      .then(function (loaded) {
        const name = getUniqueSheetName(file.name || "문서");
        sheets.push({ name: name, data: loaded });
        currentSheetIndex = sheets.length - 1;
        data = sheets[currentSheetIndex].data;
        renderSheetsTabs();
        renderForm();
        renderPreview();
        alert("DOCX 내용을 불러왔습니다.");
      })
      .catch(function (err) {
        var msg = err.message;
        if (msg === "Failed to fetch" || (msg && msg.indexOf("NetworkError") !== -1)) msg = "연결할 수 없습니다. 서버 주소와 네트워크를 확인하세요.";
        alert("불러오기 실패: " + msg);
      })
      .finally(function () {
        btnLoadDocx.disabled = false;
        btnLoadDocx.textContent = "DOCX에서 불러오기";
        fileDocx.value = "";
      });
  });

  function getMissingRequiredFields() {
    const missing = [];
    REQUIRED_KEYS.forEach(function (k) {
      const v = data[k];
      if (v == null || String(v).trim() === "") missing.push((fieldsConfig[k] && fieldsConfig[k].label) || k);
    });
    return missing;
  }

  btnDownload.addEventListener("click", function () {
    const missing = getMissingRequiredFields();
    if (missing.length > 0) {
      if (!confirm("다음 필수 항목이 비어 있습니다: " + missing.join(", ") + "\n그래도 다운로드하시겠습니까?")) return;
    }
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

  function doAutoSave() {
    try {
      localStorage.setItem("oikos-thesis-data", JSON.stringify(data));
      localStorage.setItem("oikos-thesis-sheets", JSON.stringify({ sheets: sheets, currentSheetIndex: currentSheetIndex }));
      if (autoSaveStatus) {
        autoSaveStatus.textContent = "저장됨";
        autoSaveStatus.className = "auto-save-status saved";
        clearTimeout(doAutoSave._hide);
        doAutoSave._hide = setTimeout(function () {
          autoSaveStatus.textContent = "";
          autoSaveStatus.className = "auto-save-status";
        }, 2000);
      }
    } catch (e) {}
  }

  setInterval(doAutoSave, 30 * 1000);

  document.addEventListener("keydown", function (e) {
    if (e.ctrlKey && e.key === "s") {
      e.preventDefault();
      if (btnSave) btnSave.click();
    } else if (e.ctrlKey && e.shiftKey && e.key === "S") {
      e.preventDefault();
      if (btnExportJson) btnExportJson.click();
    }
  });

  var markdownHintToggleEl = document.getElementById("markdownHintToggle");
  if (markdownHintToggleEl) {
    markdownHintToggleEl.addEventListener("change", function () {
      setMarkdownHintVisible(markdownHintToggleEl.checked);
    });
  }

  fetch("/api/fields")
    .then(function (r) { return r.json(); })
    .then(function (fields) {
      fieldsConfig = fields;
      const sheetsSaved = localStorage.getItem("oikos-thesis-sheets");
      if (sheetsSaved) {
        try {
          const parsed = JSON.parse(sheetsSaved);
          if (parsed.sheets && Array.isArray(parsed.sheets) && parsed.sheets.length > 0) {
            sheets = parsed.sheets;
            currentSheetIndex = Math.min(parsed.currentSheetIndex || 0, sheets.length - 1);
            data = sheets[currentSheetIndex].data;
            renderForm();
            renderPreview();
            renderSheetsTabs();
            return;
          }
        } catch (e) {}
      }
      const saved = localStorage.getItem("oikos-thesis-data");
      if (saved) {
        try {
          sheets[0].data = JSON.parse(saved);
          data = sheets[0].data;
        } catch (e) {}
      }
      renderForm();
      renderPreview();
      renderSheetsTabs();
    })
    .catch(function () {
      formWrap.innerHTML = "<p>필드 목록을 불러올 수 없습니다. 서버가 실행 중인지 확인하세요.</p>";
    });
})();
