# -*- coding: utf-8 -*-
import zipfile
import xml.etree.ElementTree as ET
import sys
import os
import traceback

_base = os.path.dirname(os.path.abspath(__file__))
os.chdir(_base)  # 결과 파일이 항상 스크립트 폴더에 생성되도록
_cwd = os.getcwd()
_path_in_project = os.path.join(_base, "오이코스논문 신탬플릿---수정본.docx")
path = _path_in_project if os.path.exists(_path_in_project) else r"e:\##OIKOS(AI융합)-박사과정\##논문\##오이코스 공식 논문형식\오이코스논문 신탬플릿---수정본.docx"
if not os.path.exists(path):
    with open(os.path.join(_base, "error_log.txt"), "w", encoding="utf-8") as f:
        f.write(f"DOCX not found.\n_base={_base}\ncwd={_cwd}\npath={path}\n")
    print("Error: DOCX file not found. See error_log.txt")
    sys.exit(1)
ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

def get_text(elem):
    parts = []
    for t in elem.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"):
        if t.text:
            parts.append(t.text)
        if t.tail:
            parts.append(t.tail)
    return "".join(parts).strip()

def get_style_id(p):
    pPr = p.find("w:pPr", ns)
    if pPr is not None:
        pStyle = pPr.find("w:pStyle", ns)
        if pStyle is not None:
            return pStyle.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
    return None

def get_run_props(r):
    rPr = r.find("w:rPr", ns)
    if rPr is None:
        return {}
    out = {}
    b = rPr.find("w:b", ns)
    if b is not None: out["bold"] = True
    sz = rPr.find("w:sz", ns)
    if sz is not None: out["fontSize"] = sz.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
    return out

try:
    with zipfile.ZipFile(path, "r") as z:
        # 1) 스타일 정의 (styles.xml)
        style_names = {}
        if "word/styles.xml" in z.namelist():
            with z.open("word/styles.xml") as f:
                stree = ET.parse(f)
                sroot = stree.getroot()
                for style in sroot.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style"):
                    sid = style.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId")
                    name = style.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name")
                    if sid and name is not None and name.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"):
                        style_names[sid] = name.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")

        # 2) 본문 문단 + 스타일
        with z.open("word/document.xml") as f:
            tree = ET.parse(f)
            root = tree.getroot()
            body = root.find("w:body", ns)
            if body is None:
                body = root

            out_lines = []
            for p in body.findall(".//w:p", ns):
                line = get_text(p)
                if not line:
                    continue
                style_id = get_style_id(p)
                style_name = style_names.get(style_id, style_id or "(없음)")
                out_lines.append((style_name, line))

    # 결과를 파일로 저장 (스크립트 폴더 + 현재 디렉터리)
    out_path = os.path.join(_base, "template_structure.txt")
    with open(out_path, "w", encoding="utf-8") as out:
        out.write("=== 스타일 정의 (styles.xml) ===\n")
        for sid, name in sorted(style_names.items()):
            out.write(f"  {sid} -> {name}\n")
        out.write("\n=== 문단 목록 (스타일 | 내용) ===\n")
        for style_name, line in out_lines:
            if len(line) > 70:
                line = line[:70] + "..."
            out.write(f"  [{style_name}] {line}\n")
    # Cursor 프로젝트에서 찾을 수 있도록 현재 디렉터리에도 복사
    if _cwd != _base:
        try:
            import shutil
            shutil.copy(out_path, os.path.join(_cwd, "template_structure.txt"))
        except Exception:
            pass
    print("Done. See template_structure.txt")
except Exception as e:
    err_path = os.path.join(_base, "error_log.txt")
    with open(err_path, "w", encoding="utf-8") as f:
        f.write(traceback.format_exc())
    print("Error. See error_log.txt")
    raise
