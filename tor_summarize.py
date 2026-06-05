import streamlit as st
from PyPDF2 import PdfReader
import fitz
from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os as _os
import io, zipfile, re, json as _json, base64
import requests as _requests
try:
    from google import genai as _genai
    from google.genai import types as _gtypes
    _GEMINI_AVAILABLE = True
except ImportError:
    _GEMINI_AVAILABLE = False
from datetime import datetime

st.set_page_config(page_title="TOR Summarizer", page_icon="bb.png", layout="wide")

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
        padding: 2rem; border-radius: 12px; margin-bottom: 2rem;
        text-align: center; color: white;
    }
    .main-header h1 { margin: 0; }
    .main-header p { opacity: 0.8; margin: 0.5rem 0 0; }
    .stButton > button { border-radius: 8px; }
    .export-box { background: #f0fdf4; border: 1px solid #86efac; border-radius: 10px; padding: 1.2rem; margin-top: 1rem; }

    /* โหลด Material Icons สำหรับ Streamlit UI */
    @import url('https://fonts.googleapis.com/icon?family=Material+Icons&display=swap');
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>📄 TOR Summarizer</h1>
    <p>โยน ZIP/PDF → สรุปอัตโนมัติ → Export รวมไฟล์เดียว</p>
</div>
""", unsafe_allow_html=True)

_SUPPORTED_MODELS = {
    "llama3.2":                                          "LLaMA 3.2",
    "hf.co/nectec/Pathumma-llm-text-1.0.0:Q4_K_M":     "Pathumma 7B (NECTEC)",
    #"hf.co/nectec/thai-research-gemma-3-27b-it:Q4_K_M": "Pathumma 27B Gemma (NECTEC)",
}

with st.sidebar:
    try:
        _r = _requests.get("http://localhost:11434/api/tags", timeout=2)
        _installed = [m["name"] for m in _r.json().get("models", [])]
        _ollama_ok = True
    except Exception:
        _installed = []
        _ollama_ok = False

    if not _ollama_ok:
        st.error("Ollama ไม่ได้รัน\nเปิดโปรแกรม Ollama ก่อน")
    elif not _installed:
        st.warning("Ollama รันอยู่ แต่ยังไม่มีโมเดล")
    else:
        st.success(f"Ollama พร้อมใช้งาน ({len(_installed)} โมเดล)")

    st.divider()
    st.header("เลือกโมเดล")

    _available = [m for m in _installed if m] or ["llama3.2"]
    def _make_label(m):
        if m in _SUPPORTED_MODELS:
            return _SUPPORTED_MODELS[m]
        return m.split("/")[-1].replace(":latest", "").replace(":Latest", "")
    _labels = [_make_label(m) for m in _available]
    _label_to_model = dict(zip(_labels, _available))

    _default_label = next(
        (lb for lb, md in _label_to_model.items() if "pathumma" in md.lower() or "nectec" in md.lower()),
        next((lb for lb, md in _label_to_model.items() if "llama3.2" in md), _labels[0])
    )
    _chosen_label = st.selectbox(
        "โมเดลสำหรับสรุป", options=_labels,
        index=_labels.index(_default_label),
        help="เลือกโมเดล LLM ที่ติดตั้งใน Ollama"
    )
    SELECTED_MODEL = _label_to_model[_chosen_label]

    if SELECTED_MODEL not in _installed:
        st.info(f"ติดตั้งโมเดลนี้ด้วย:\n`ollama pull {SELECTED_MODEL}`")

    st.divider()
    st.header("OCR with Google Vision")
    GEMINI_KEY = st.text_input("API Key", type="password", placeholder="AIza...",
        help="ใช้สำหรับอ่าน PDF สแกน แม่นกว่า Tesseract \nรับฟรีที่ aistudio.google.com")
    if not GEMINI_KEY:
        st.caption("⚠️ ไม่มี Gemini Key → ใช้ Tesseract แทน")
    else:
        st.caption("✅ Google Vision OCR พร้อมใช้งาน")

    st.divider()
    st.header("📂 เกี่ยวกับเว็บไซต์")
    st.caption("เว็บไซต์นี้เป็นเครื่องมือเพื่อช่วยสรุปเอกสาร Terms of Reference (TOR) อัตโนมัติเพื่อลดเวลาและภาระในการสรุปเอกสารทั้งหมด พร้อมนำออกในรูปแบบ Word/PDF รวมทุกโครงการ\n\n**หมายเหตุ:** ผลลัพธ์ที่ได้อาจมีความคลาดเคลื่อน ควรตรวจสอบกับเอกสารต้นฉบับอีกครั้ง")
    st.divider()
    st.header("🛠️ เครื่องมือที่ใช้")
    st.caption(f"**{_chosen_label}** — สรุป")
    if GEMINI_KEY:
        st.caption("**Google Vision** — OCR")
    else:
        st.caption("**Tesseract** — OCR")

# Helper
# ชื่อไฟล์ที่รู้ชัดว่าไม่ใช่ TOR
_NOT_TOR = [
    "contract", "bond", "quotation", "bidding", "notice", "noltice",
    "performance", "advance", "payment", "action_plan", "action",
    "definition", "verification", "integrity", "pact", "bid", "spec",
    "สัญญา", "ประกาศ", "ใบเสนอ", "แบบ",
]

def _is_not_tor(filename: str) -> bool:
    lower = filename.lower()
    return any(kw in lower for kw in _NOT_TOR)

def find_tor_in_zip(zip_bytes):
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        names = z.namelist()
        pdfs = [n for n in names if n.lower().endswith(".pdf")]

        # หาจากชื่อที่มีคำว่า tor ก่อน
        tor_files = []
        for name in pdfs:
            lower = name.lower()
            if "attach_tor" in lower:
                tor_files.append((0, name))
            elif lower.startswith("tor") or "_tor" in lower or "tor_" in lower:
                tor_files.append((1, name))
            elif "tor" in lower:
                tor_files.append((2, name))

        if tor_files:
            tor_files.sort()
            return [(z.read(n), n.split("/")[-1]) for _, n in tor_files]

        # กรองที่รู้ว่าไม่ใช่ TOR ออก แล้วเอาที่เหลือ
        candidates = [
            (z.getinfo(n).file_size, n)
            for n in pdfs
            if not _is_not_tor(n.split("/")[-1])
        ]
        if candidates:
            candidates.sort(reverse=True)  # ใหญ่สุดก่อน
            return [(z.read(n), n.split("/")[-1]) for _, n in candidates]

        # ไฟล์ที่ใหญ่ที่สุด
        if pdfs:
            biggest = max(pdfs, key=lambda x: z.getinfo(x).file_size)
            return [(z.read(biggest), biggest.split("/")[-1])]
        return []

def get_project_id(zip_name):
    match = re.match(r"(\d+)", zip_name)
    return match.group(1) if match else zip_name.replace(".zip", "")

# หมุนเอกสารถ้าเจอว่าหมุนอยู่
def rotate_pdf_if_needed(file_bytes):
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
    except Exception:
        return file_bytes

    rotated = False
    for page in doc:
        if page.rotation in (90, 180, 270): # องศาที่หมุนอยู่
            page.set_rotation(0) # ปรับเป็น 0 องศา
            rotated = True

    if rotated:
        out_bytes = io.BytesIO()
        doc.save(out_bytes)
        doc.close()
        return out_bytes.getvalue()

    doc.close()
    return file_bytes

def extract_text_from_pdf(file_bytes):
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
        text = ""
        for page in reader.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
        return text.strip()
    except Exception:
        return ""

def ocr_with_gemini(file_bytes, gemini_key, label=""):
    """ส่ง PDF ทั้งไฟล์ให้ Gemini อ่าน — 1 request ต่อไฟล์ ไม่กิน quota เยอะ"""
    import time, re as _re
    client = _genai.Client(api_key=gemini_key)
    with st.spinner(f"Google Vision อ่าน {label} (1 request)..."):
        for attempt in range(5):
            try:
                resp = client.models.generate_content(
                    model="gemini-2.5-flash-lite",
                    contents=[
                        _gtypes.Part.from_bytes(data=file_bytes, mime_type="application/pdf"),
                        "อ่านและถอดข้อความทั้งหมดจาก PDF นี้ให้ครบถ้วนทุกหน้า รักษาโครงสร้างเดิม ห้ามสรุปหรือตัดทอน"
                    ]
                )
                raw = resp.text
                # Gemini บางครั้งคืนเป็น JSON array — ดึงเฉพาะ text field
                if raw.strip().startswith('[') or raw.strip().startswith('{'):
                    try:
                        import json as _j
                        data = _j.loads(raw)
                        if isinstance(data, list):
                            raw = "\n\n".join(
                                item.get("text", str(item)) if isinstance(item, dict) else str(item)
                                for item in data
                            )
                        elif isinstance(data, dict):
                            raw = data.get("text", raw)
                    except Exception:
                        pass
                return raw
            except Exception as e:
                msg = str(e)
                if "429" in msg or "RESOURCE_EXHAUSTED" in msg:
                    wait = 65
                    m = _re.search(r"retry.{0,10}after[^\d]*(\d+)", msg, _re.IGNORECASE)
                    if m: wait = int(m.group(1)) + 2
                    st.warning(f"รอ {min(wait,120)} วินาที...")
                    time.sleep(min(wait, 120))
                else:
                    st.error(f"❌ Google Vision error: {e}"); return ""
    return ""

def ocr_with_tesseract(file_bytes, label=""):
    """Fallback OCR ด้วย Tesseract"""
    try:
        import pytesseract, numpy as np
        _tess_win = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        if _os.path.exists(_tess_win):
            pytesseract.pytesseract.tesseract_cmd = _tess_win
    except ImportError:
        st.error("❌ ไม่พบ pytesseract"); return ""
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    total = len(doc)
    all_text = []
    progress = st.progress(0, text=f"OCR {label} ด้วย Tesseract ({total} หน้า)")
    for i in range(total):
        page = doc[i]
        pix = page.get_pixmap(matrix=fitz.Matrix(200/72, 200/72))
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
        text = pytesseract.image_to_string(img, lang="tha+eng",
            config="--oem 1 --psm 6 -c preserve_interword_spaces=1")
        all_text.append(text)
        progress.progress((i + 1) / total, text=f"OCR หน้า {i+1}/{total}")
    progress.empty(); doc.close()
    return "\n\n".join(all_text)

def ocr_pdf(file_bytes, label=""):
    """เลือก OCR อัตโนมัติ — Google Vision ถ้ามี Key, Tesseract ถ้าไม่มี"""
    gkey = GEMINI_KEY if "GEMINI_KEY" in globals() else ""
    if gkey and _GEMINI_AVAILABLE:
        return ocr_with_gemini(file_bytes, gkey, label=label)
    return ocr_with_tesseract(file_bytes, label=label)

def extract_spec_raw(text: str) -> str:
    """ดึงเฉพาะส่วนคุณลักษณะเฉพาะของสินค้า (ข้อ ๔) ไม่เอาหลักการ/คุณสมบัติ/เงื่อนไข"""
    import re as _re

    # แยกไฟล์ใน ZIP ออกจากกัน เอาเฉพาะ part ที่มีสเปคสินค้า
    parts = _re.split(r'={10,}', text)
    spec_parts = [p for p in parts if any(kw in p for kw in
        ["คุณลักษณะทั่วไป", "คุณลักษณะทางเทคนิค",
         "รายละเอียดคุณลักษณะเฉพาะ", "ข้อกำหนดทางเทคนิค",
         "คุณลักษณะเฉพาะหรือขอบเขต"])]
    work_text = "\n".join(spec_parts) if spec_parts else text

    lines = work_text.split("\n")
    start_idx = None

    START_RE = [
        r"[๔4][\.,\s]+รายละเอียดคุณลักษณะ",
        r"[๔4][\.,\s]+คุณลักษณะเฉพาะ",
        r"[๔4][\.,\s]+คุณลักษณะ",
        r"คุณลักษณะทั่วไป\s+มีอย่างน้อย",
        r"คุณลักษณะทางเทคนิค\s+มีอย่างน้อย",
        r"[๔4][\.,][๑1][\.,]\s*คุณลักษณะ",
    ]
    END_KW = [
        "กำหนดเวลาส่งมอบ", "ระยะเวลาส่งมอบ",
        "หลักเกณฑ์ในการพิจารณา", "หลักเกณฑ์การพิจารณา",
        "หลักฐานการยื่นข้อเสนอ", "การเสนอราคา",
        "การทำสัญญา", "ค่าจ้างและการจ่าย", "อัตราค่าปรับ",
        "บัญชีเอกสาร", "คุณสมบัติของผู้ยื่น", "คุณสมบัติผู้เสนอ",
    ]
    NOISE_RE = [
        r"^ประธานกรรมการ", r"^กรรมการ\s*ลงชื่อ", r"^ลงชื่อ[.\s]",
        r"^\(นาย", r"^\(นาง", r"^\(น\.ส\.",
        r"^\[Page\s+\d+\]", r"^-[๐-๙0-9]+-$",
    ]

    for i, line in enumerate(lines):
        s = line.strip()
        if any(_re.search(p, s) for p in START_RE):
            start_idx = i; break
        if _re.match(r"[๔4][.][๑1]", s) and "คุณลักษณะ" in s:
            start_idx = i; break

    if start_idx is None:
        return ""

    spec_lines = []
    for line in lines[start_idx:]:
        s = line.strip()
        if s and any(kw in s for kw in END_KW):
            break
        if s and any(_re.match(p, s) for p in NOISE_RE):
            continue
        if len(s) > 8:
            ok = sum(1 for c in s if "\u0e00" <= c <= "\u0e7f" or c.isalnum() or c in " .,()%:-x/°")
            if ok / len(s) < 0.20:
                continue
        spec_lines.append(line)

    cleaned, prev_blank = [], False
    for ln in spec_lines:
        blank = ln.strip() == ""
        if blank and prev_blank:
            continue
        cleaned.append(ln)
        prev_blank = blank

    result = "\n".join(cleaned).strip()
    return result if len(result) > 100 else ""
_THAI_DIGITS = str.maketrans("๐๑๒๓๔๕๖๗๘๙", "0123456789")

def _to_arabic(text: str) -> str:
    return text.translate(_THAI_DIGITS)

def extract_budget_by_regex(text: str) -> dict:
    """
    ดึงงบประมาณและระยะเวลา
    - budget_annual = วงเงินที่ได้รับจัดสรร/ปีนี้ (หลัก)

    - รองรับเลขไทย/อารบิก
    """
    result = {"budget": "", "duration": ""}
    t = _to_arabic(text)
    N = r"[\d,]+(?:\.[\d]+)?"

    EXCLUDE_KEYWORDS = ["ผลงาน", "ไม่น้อยกว่า", "หลักประกัน", "ค้ำประกัน",
                        "ประกันสัญญา", "ค่าปรับ", "สำรองจ่าย", "มูลค่าผลงาน"]

    def find_in_budget_section(full_text: str) -> str:
        section_triggers = ["วงเงิน", "งบประมาณ", "ราคากลาง", "ผูกพัน"]
        lines_all = full_text.split("\n")
        budget_lines = []
        in_section = False
        for line in lines_all:
            if any(kw in line for kw in section_triggers):
                in_section = True
            elif in_section and len(line.strip()) < 40 and any(
                kw in line for kw in ["คุณสมบัติ", "เงื่อนไข", "การจ่าย", "หลักเกณฑ์", "การพิจารณา"]
            ):
                in_section = False
            if in_section and not any(kw in line for kw in EXCLUDE_KEYWORDS):
                budget_lines.append(line)
        return "\n".join(budget_lines) if budget_lines else full_text

    budget_text = find_in_budget_section(t)

    # วงเงินที่ได้รับจัดสรร/ปีนี้ (เฉพาะเจาะจง)
    annual_pats = [
        (rf"วงเงิน[ที่]?ได้รับจัดสรร\s*({N})\s*บาท", False),
        (rf"ผูกพันงบประมาณ[^0-9\n]{{0,50}}จำนวน\s*({N})\s*บาท", False),
        (rf"ราคากลาง[^0-9\n]{{0,20}}({N})\s*ล้านบาท", True),
        (rf"ราคากลาง[^0-9\n]{{0,20}}({N})\s*บาท", False),
        (rf"วงเงินงบประมาณ[^0-9\n]{{0,30}}({N})\s*บาท", False),
        (rf"เป็นเงิน\s*({N})\s*บาท", False),
        (rf"จำนวน\s*({N})\s*บาท", False),
        (rf"({N})\s*บาท", False),
    ]

    def find_num(pats, search_text):
        for pat, is_million in pats:
            m = re.search(pat, search_text)
            if m:
                num_str = m.group(1)
                try:
                    val = float(num_str.replace(",", ""))
                    if val < 1000 and not is_million:
                        continue
                    return f"{num_str} ล้านบาท" if is_million else f"{num_str} บาท"
                except Exception:
                    pass
        return ""

    result["budget"] = find_num(annual_pats, budget_text)

    dur_pats = [
        rf"ผูกพัน[^0-9\n]{{0,30}}([\d]+)\s*เดือน",
        rf"เป็นระยะเวลา\s*([\d]+)\s*(เดือน|ปี|วัน)",
        rf"ระยะเวลา[^0-9\n]{{0,30}}([\d]+)\s*(เดือน|ปี|วัน)",
        rf"([\d]+)\s*เดือน",
        rf"([\d]+)\s*ปี",
    ]
    for pat in dur_pats:
        m = re.search(pat, t)
        if m:
            groups = m.groups()
            if len(groups) == 1:
                unit = "เดือน" if "เดือน" in pat else "ปี"
                result["duration"] = f"{groups[0]} {unit}"
            else:
                result["duration"] = f"{groups[0]} {groups[1]}"
            break
    return result

def summarize_tor(text, api_key="", project_id=""):
    extracted = extract_budget_by_regex(text)
    budget_fact = extracted["budget"] or "ไม่พบในเอกสาร"
    duration_fact = extracted["duration"] or "ไม่พบในเอกสาร"
    _model = SELECTED_MODEL if "SELECTED_MODEL" in globals() else "llama3.2"
    is_pathumma = "pathumma" in _model.lower()

    prompt_general = f"""คุณเป็นผู้เชี่ยวชาญด้านการวิเคราะห์เอกสาร TOR (Terms of Reference) ของราชการไทย

กฎสำคัญ:
- ดึงข้อมูลจากเอกสารเท่านั้น ห้ามแต่งหรือเดา
- ถ้าไม่พบในหัวข้อใด ให้เขียนว่า "ไม่ระบุในเอกสาร"
- ตอบเป็นภาษาไทย ยกเว้นคำเทคนิคภาษาอังกฤษให้คงไว้

สรุปเอกสาร TOR นี้:

**รหัสโครงการ:** {project_id}
**ชื่อโครงการ:**
**หน่วยงานเจ้าของโครงการ:**
**วัตถุประสงค์:**
**ขอบเขตงาน:**
**คุณสมบัติผู้รับจ้าง:**
**วงเงินงบประมาณ:** {budget_fact}
**ระยะเวลาดำเนินงาน:** {duration_fact}
**เงื่อนไขสำคัญ:**

เนื้อหาเอกสาร:
{text[:3000]}
"""
    try:
        status = st.empty()
        status.info(f"⏳ [{_model.split('/')[-1].split(':')[0]}] กำลังสรุป...")
        if is_pathumma:
            resp = _requests.post("http://localhost:11434/api/chat",
                json={"model": _model, "messages": [
                    {"role": "system", "content": "คุณเป็นผู้เชี่ยวชาญด้านการวิเคราะห์เอกสาร TOR ของราชการไทย ตอบเป็นภาษาไทยเสมอ ห้ามแต่งข้อมูลที่ไม่มีในเอกสาร"},
                    {"role": "user", "content": prompt_general}
                ], "stream": True, "options": {"num_predict": 2000, "num_ctx": 4096, "temperature": 0.1}},
                stream=True, timeout=(10, 1200))
        else:
            resp = _requests.post("http://localhost:11434/api/generate",
                json={"model": _model, "prompt": prompt_general, "stream": True,
                      "options": {"num_predict": 2000, "num_ctx": 4096, "temperature": 0.1}},
                stream=True, timeout=(10, 1200))
        resp.raise_for_status()
        summary_general = ""
        for line in resp.iter_lines():
            if line:
                try:
                    chunk = _json.loads(line)
                    if is_pathumma:
                        summary_general += chunk.get("message", {}).get("content", "")
                    else:
                        summary_general += chunk.get("response", "")
                    if chunk.get("done"): break
                except _json.JSONDecodeError:
                    continue
        status.empty()
    except _requests.exceptions.ConnectionError:
        st.error("❌ ไม่พบ Ollama — กรุณาเปิด Ollama ก่อน"); return ""
    except Exception as e:
        st.error(f"❌ Ollama error: {e}"); return ""

    if not summary_general:
        return ""

    # ส่วนคุณลักษณะเฉพาะ ไม่ตัดทอน
    spec_raw = extract_spec_raw(text)
    spec_out = ("\n\n---\n**📋 คุณลักษณะเฉพาะ (ข้อความจากเอกสาร)**\n\n" + spec_raw
                if spec_raw else "\n\n---\n**📋 คุณลักษณะเฉพาะ**\nไม่พบในเอกสาร")

    return summary_general.strip() + spec_out

# Export Word
def add_page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._r.append(br)

def build_combined_word(results):
    doc = DocxDocument()
    doc.add_paragraph(); doc.add_paragraph()
    cover_title = doc.add_heading("รายงานสรุปเอกสาร TOR", 0)
    cover_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover_title.runs[0].font.color.rgb = RGBColor(0x0f, 0x34, 0x60)
    sub = doc.add_paragraph(f"จำนวน {len(results)} โครงการ")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].font.size = Pt(14)
    sub.runs[0].font.color.rgb = RGBColor(0x44, 0x44, 0x44)
    date_p = doc.add_paragraph(f"วันที่จัดทำ: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_p.runs[0].font.size = Pt(11)
    date_p.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    doc.add_paragraph()
    add_page_break(doc)
    toc = doc.add_heading("สารบัญ", level=1)
    toc.runs[0].font.color.rgb = RGBColor(0x0f, 0x34, 0x60)
    for idx, r in enumerate(results, 1):
        doc.add_paragraph(style="List Number").add_run(f"รหัสโครงการ {r['project_id']}").bold = True
    for idx, r in enumerate(results, 1):
        add_page_break(doc)
        h = doc.add_heading(f"โครงการที่ {idx}", level=1)
        h.runs[0].font.color.rgb = RGBColor(0x0f, 0x34, 0x60)
        m = doc.add_paragraph(); m.add_run("รหัสโครงการ: ").bold = True; m.add_run(r["project_id"])
        m2 = doc.add_paragraph(); m2.add_run("ไฟล์ต้นฉบับ: ").bold = True; m2.add_run(r["source"])
        doc.add_paragraph()
        for line in r["summary"].split("\n"):
            line = line.strip()
            if not line:
                doc.add_paragraph(); continue
            if line.startswith("**") and line.endswith("**"):
                hh = doc.add_heading(line.replace("**", ""), level=2)
                if hh.runs: hh.runs[0].font.color.rgb = RGBColor(0x0f, 0x34, 0x60)
            elif line.startswith("# ") or line.startswith("📋"):
                # spec header
                hh = doc.add_heading(line.lstrip("# "), level=2)
                if hh.runs: hh.runs[0].font.color.rgb = RGBColor(0x0f, 0x34, 0x60)
            elif line.startswith("- ") or line.startswith("• "):
                doc.add_paragraph(style="List Bullet").add_run(line[2:])
            elif line.startswith("===") or line.startswith("---"):
                # separator
                p = doc.add_paragraph()
                p.add_run("─" * 40)
            else:
                p = doc.add_paragraph()
                for i, part in enumerate(line.split("**")):
                    run = p.add_run(part)
                    if i % 2 == 1: run.bold = True
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()

# Export Excel
def build_combined_excel(results):
    import pandas as pd
    data = []
    for r in results:
        extracted = extract_budget_by_regex(r["summary"])
        data.append({
            "project_id": r["project_id"],
            "source_file": r["source"],
            "budget": extracted["budget"],
            "duration": extracted["duration"],
            "summary": r["summary"]
        })
    df = pd.DataFrame(data)
    buf = io.BytesIO()
    try:
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="TOR Summary")
            ws = writer.sheets["TOR Summary"]
            ws.set_column("A:A", 20)
            ws.set_column("B:B", 30)
            ws.set_column("C:D", 20)
            ws.set_column("E:E", 15)
            ws.set_column("F:F", 100)
    except ImportError:
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="TOR Summary")
            ws = writer.sheets["TOR Summary"]
            widths = {"A": 20, "B": 30, "C": 20, "D": 20, "E": 15, "F": 100}
            for col, width in widths.items():
                ws.column_dimensions[col].width = width
    buf.seek(0)
    return buf.getvalue()

# Process
def process_one(uploaded_file):
    name = uploaded_file.name
    file_bytes = uploaded_file.read()
    if name.lower().endswith(".zip"):
        project_id = get_project_id(name)
        tor_files = find_tor_in_zip(file_bytes)
        if not tor_files:
            st.error(f"ไม่พบไฟล์ TOR ใน {name}")
            return None
        filenames = [fn for _, fn in tor_files]
        st.success(f"**{project_id}** — พบ TOR {len(tor_files)} ไฟล์: {', '.join(filenames)}")
        # รวม text จากทุกไฟล์ TOR
        all_texts = []
        for pdf_b, fn in tor_files:
            pdf_b = rotate_pdf_if_needed(pdf_b)
            t = extract_text_from_pdf(pdf_b)
            if t and len(t) > 100:
                st.info(f"{fn} — อ่านข้อความโดยตรง")
                all_texts.append(f"=== {fn} ===\n{t}")
            else:
                st.info(f"{fn} — PDF สแกน, OCR ด้วย Tesseract")
                t = ocr_pdf(pdf_b, label=fn)
                all_texts.append(f"=== {fn} ===\n{t}")
        text = "\n\n".join(all_texts)
        source_label = ", ".join(filenames)
    else:
        project_id = name.replace(".pdf", "")
        pdf_bytes = rotate_pdf_if_needed(file_bytes)
        source_label = name
        st.success(f"**{project_id}**")
        text = extract_text_from_pdf(pdf_bytes)
        if not (text and len(text) > 100):
            st.info("PDF สแกน — OCR ด้วย Tesseract")
            text = ocr_pdf(pdf_bytes, label=project_id)
        else:
            st.info("PDF พิมพ์ — อ่านข้อความโดยตรง")

    summary = summarize_tor(text, project_id=project_id)
    return {"project_id": project_id, "source": source_label, "summary": summary}

# Main UI
uploaded_files = st.file_uploader(
    "อัปโหลดไฟล์ (ZIP หรือ PDF)",
    type=["zip", "pdf"],
    accept_multiple_files=True,
    help="โยนได้หลายไฟล์พร้อมกัน"
)

if uploaded_files:
    n = len(uploaded_files)
    if st.button(f"สรุปทั้งหมด ({n} โครงการ)", type="primary"):
        results = []
        st.markdown("### กำลังประมวลผล...")
        for i, f in enumerate(uploaded_files):
            st.markdown(f"**{i+1}/{n}** — `{f.name}`")
            try:
                result = process_one(f)
                if result:
                    results.append(result)
            except Exception as e:
                st.error(f"{f.name}: {str(e)}")

        if results:
            st.markdown("---")
            st.markdown(f"### สรุปครบ {len(results)}/{n} โครงการ")
            for r in results:
                with st.expander(f"📋 {r['project_id']}", expanded=False):
                    st.markdown(r["summary"])
            st.markdown("**⬇️ Export รวมทุกโครงการ:**")
            now_str = datetime.now().strftime("%Y%m%d_%H%M")
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("Word (.docx)",
                    data=build_combined_word(results),
                    file_name=f"สรุป_TOR_รวม_{now_str}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True)
            with c2:
                st.download_button("excel (.xlsx)",
                    data=build_combined_excel(results),
                    file_name=f"สรุป_TOR_รวม_{now_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
else:
    st.markdown("""
    <div style="text-align:center;padding:3rem;color:#888;border:2px dashed #ddd;border-radius:12px;">
        <div style="font-size:3rem;">📦</div>
        <div style="margin-top:1rem;font-size:1.1rem;">โยนไฟล์ ZIP หรือ PDF</div>
        <div style="font-size:0.85rem;margin-top:0.5rem;">รองรับหลายโครงการพร้อมกัน พร้อม Export รวมเป็นไฟล์เดียว</div>
    </div>
    """, unsafe_allow_html=True)