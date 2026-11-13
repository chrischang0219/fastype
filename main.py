import os
import gc
import json
import uuid
import threading
from urllib.parse import quote
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, Response
from fastapi.middleware.cors import CORSMiddleware

import base64
import re
import time
from datetime import datetime, timedelta, timezone
import io
import os
import requests
import urllib3
import fitz  # PyMuPDF
from PIL import Image
try:
    import numpy as np
except ImportError:
    np = None
try:
    import cv2
except ImportError:
    cv2 = None
from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_TAB_ALIGNMENT
try:
    import latex2mathml.converter
    import mathml2omml
    HAS_OMML = True
except ImportError:
    HAS_OMML = False

# 忽略 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================
# ⚙️ 全域字體與模型設定
# ==========================================
DOC_TITLE_FONT = "王漢宗顏楷體繁"
DOC_CHINESE_FONT = "標楷體"
DOC_ENGLISH_FONT = "Times New Roman"
VALID_MODELS = [
    "gemini-3.8-flash",      # 推薦：最新主力，辨識精準度高 (5 RPM / 250K TPM)
    "gemini-3.6-flash",      # 備用主力：同等級高精準度 (5 RPM / 250K TPM)
    "gemini-3.5-flash-lite", # 省額度首選：速度極快、每分鐘上限高 (15 RPM / 250K TPM)
    "gemini-3.1-flash-lite", # 輕量備援：同樣享有高每分鐘次數 (15 RPM / 250K TPM)
    "gemini-2.5-flash-lite", # 穩定輕量版 (10 RPM / 250K TPM)
    "gemini-2.5-flash"       # 舊版備援 (5 RPM / 250K TPM)
]


# ==========================================
# 🛠️ 核心排版與處理函式
# ==========================================
def parse_page_selection(s, max_p):
    if not s or not s.strip():
        return list(range(max_p))
    selected = set()
    try:
        for part in s.split(','):
            if '-' in part:
                start, end = map(int, part.split('-'))
                for p in range(max(1, start), min(max_p, end) + 1):
                    selected.add(p - 1)
            else:
                p = int(part)
                if 1 <= p <= max_p:
                    selected.add(p - 1)
        return sorted(list(selected)) if selected else list(range(max_p))
    except Exception:
        return list(range(max_p))

def extract_math_placeholders(text):
    """將 $$...$$ 與 $...$ 暫存為 __MATH_i__，避免被全形標點轉換破壞 LaTeX 語法"""
    math_list = []
    def _repl(m):
        idx = len(math_list)
        math_list.append(m.group(1).strip())
        return f"__MATH_{idx}__"
    text = re.sub(r'\$\$(.+?)\$\$', _repl, text, flags=re.DOTALL)
    text = re.sub(r'\$([^$]+?)\$', _repl, text)
    return text, math_list

def add_text_with_equations(p, text, size, is_english_mode, bold=False, math_list=None):
    """將含有 __MATH_i__ 的文字寫入段落，並把 LaTeX 轉為 Word 原生新式方程式 (OMML)"""
    if not math_list:
        run = p.add_run(text)
        set_run_font(run, size=size, is_english_mode=is_english_mode, bold=bold)
        return

    parts = re.split(r'(__MATH_\d+__)', text)
    for part in parts:
        if not part:
            continue
        m_idx = re.match(r'^__MATH_(\d+)__$', part)
        if m_idx:
            latex_inner = math_list[int(m_idx.group(1))]
            # 判斷是否真的需要方程式編輯器（只有分數、根號、線段、上下標、向量等才轉 OMML）
            needs_omml = any(k in latex_inner for k in ['\\frac', '\\sqrt', '\\overline', '^', '_', '\\vec', '\\sum', '\\int'])
            # 若為單純線段 \overline{...} 或 \bar{...}（即使內含 \text{} 或同一組 $...$ 內有多個線段），直接套用 Word 原生上方橫線 (<m:bar>) + 正體字模板
            latex_norm = re.sub(r'\\(?:text|mathrm|mathbf|mathit)\{([^{}]*)\}', r'\1', latex_inner)
            if re.search(r'\\(?:overline|bar)\{', latex_norm) and not any(k in latex_norm for k in ['\\frac', '\\sqrt', '^', '_', '\\vec', '\\sum', '\\int']):
                sub_parts = re.split(r'(\\(?:overline|bar)\{\s*[^{}]+?\s*\})', latex_norm)
                for sp in sub_parts:
                    if not sp:
                        continue
                    m_b = re.match(r'^\\(?:overline|bar)\{\s*([^{}]+?)\s*\}$', sp)
                    if m_b:
                        seg_txt = m_b.group(1).strip()
                        omml_str = (
                            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                            '<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>'
                            f'<m:e><m:r><m:rPr><m:sty m:val="p"/></m:rPr><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><m:t>{seg_txt}</m:t></m:r></m:e>'
                            '</m:bar></m:oMath>'
                        )
                        p._element.append(parse_xml(omml_str))
                    else:
                        sp_txt = custom_format_text(sp)
                        if sp_txt:
                            run = p.add_run(sp_txt)
                            set_run_font(run, size=size, is_english_mode=is_english_mode, bold=bold)
                continue

            if HAS_OMML and needs_omml:
                try:
                    # 比例冒號轉全形、平行符號轉 //
                    latex_clean = latex_inner.replace(':', '：')
                    latex_clean = latex_clean.replace('\\parallel', '//').replace('||', '//')
                    mathml = latex2mathml.converter.convert(latex_clean)
                    omml_str = mathml2omml.convert(mathml)

                    # 1. 將短上標線 (<m:acc>) 轉為全寬線段 (<m:bar>)
                    omml_str = re.sub(
                        r'<m:acc>(?:\s*<m:accPr>.*?</m:accPr>)?\s*<m:e>(.*?)</m:e>\s*</m:acc>',
                        r'<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>\1</m:e></m:bar>',
                        omml_str,
                        flags=re.DOTALL
                    )

                    # 2. 隱藏根號左上角多餘的次數虛線方框 (<m:degHide m:val="1"/>)
                    omml_str = re.sub(
                        r'<m:rad>\s*(?!<m:radPr>)',
                        '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr>',
                        omml_str
                    )

                    # 3. 強制將方程式內的英數字設為正體 (<m:sty m:val="p"/>)，不要斜體
                    omml_str = re.sub(r'<m:rPr>.*?</m:rPr>', '', omml_str, flags=re.DOTALL)
                    omml_str = re.sub(r'<m:r>', '<m:r><m:rPr><m:sty m:val="p"/></m:rPr><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>', omml_str)

                    if 'xmlns:m=' not in omml_str:
                        omml_str = omml_str.replace(
                            '<m:oMath>',
                            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
                            1
                        )
                    omml_el = parse_xml(omml_str)
                    p._element.append(omml_el)
                    continue
                except Exception:
                    pass
            # 若為單純英文字母、比例或轉換失敗，直接以純文字輸出
            fallback_txt = custom_format_text(latex_inner)
            run = p.add_run(fallback_txt)
            set_run_font(run, size=size, is_english_mode=is_english_mode, bold=bold)
        else:
            run = p.add_run(part)
            set_run_font(run, size=size, is_english_mode=is_english_mode, bold=bold)

def custom_format_text(text):
    text = re.sub(r'</?[ub*]+>', '', text)
    
    # 清除 LaTeX 數學與幾何符號 ($...$, \text, \overline, \frac, 次方等)
    text = text.replace('$$', '').replace('$', '')
    text = re.sub(r'\\(?:text|mathrm|mathbf|mathit)\{([^}]*)\}', r'\1', text)
    text = re.sub(r'\\overline\{([^}]+)\}', r'\1', text)
    text = re.sub(r'\\frac\{([^}]+)\}\{([^}]+)\}', r'\1__UNITSLASH__\2', text)
    sup_map = {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹'}
    text = re.sub(r'\^\{?([0-9])\}?', lambda m: sup_map.get(m.group(1), m.group(0)), text)
    text = text.replace('\\times', '×').replace('\\div', '÷').replace('\\pm', '±')
    text = text.replace('\\leq', '≤').replace('\\geq', '≥').replace('\\neq', '≠').replace('\\approx', '≈')
    text = text.replace('\\bot', '⊥').replace('\\perp', '⊥').replace('\\parallel', '//')
    text = text.replace('\\triangle', '△').replace('\\angle', '∠').replace('^\\circ', '°').replace('\\circ', '°')
    
    # 統一將大題標題的標點（如 二. 或 二．）轉為頓號（二、）
    text = re.sub(r'^\s*([一二三四五六七八九十]+)\s*[、.．]\s*', r'\1、', text)

    # 統一轉換選項格式為 (A) (B) (C) (D)
    # 僅針對 ①②③④ 或 (1)(2)(3)(4)，避免誤殺題號 1. 2. 10.
    text = re.sub(r'^\s*(?:①|[（\(]1[）\)])\s*[.．、\s]*', '(A)', text)
    text = re.sub(r'^\s*(?:②|[（\(]2[）\)])\s*[.．、\s]*', '(B)', text)
    text = re.sub(r'^\s*(?:③|[（\(]3[）\)])\s*[.．、\s]*', '(C)', text)
    text = re.sub(r'^\s*(?:④|[（\(]4[）\)])\s*[.．、\s]*', '(D)', text)
    
    # 針對 A~D 選項，必須有括號或後方接標點/空白，避免誤殺英文單字開頭 (如 CONTENT_START)
    text = re.sub(r'^\s*[（\(]A[）\)]\s*[.．、\s]*|^\s*A[.．、]\s*', '(A)', text, flags=re.IGNORECASE)
    text = re.sub(r'^\s*[（\(]B[）\)]\s*[.．、\s]*|^\s*B[.．、]\s*', '(B)', text, flags=re.IGNORECASE)
    text = re.sub(r'^\s*[（\(]C[）\)]\s*[.．、\s]*|^\s*C[.．、]\s*', '(C)', text, flags=re.IGNORECASE)
    text = re.sub(r'^\s*[（\(]D[）\)]\s*[.．、\s]*|^\s*D[.．、]\s*', '(D)', text, flags=re.IGNORECASE)

    # 修正英文縮寫撇號後方多餘空白 (如 It’ s -> It's)
    text = re.sub(r"([a-zA-Z])['’]\s+([a-zA-Z])", r"\1'\2", text)

    # 標點符號全形轉換（保護小數點、題號點、單位斜線）
    # 1. 先將受保護的符號暫存
    text = re.sub(r'^(\s*\d+)\.', r'\1__QDOT__', text)          # 題號的點 (如 7.)
    text = re.sub(r'(?<=\d)\.(?=\d)', '__DECDOT__', text)        # 小數點 (如 100.0)
    text = text.replace('//', '__PARALLEL__') # 保護平行符號 //，避免被轉成 ÷÷
    text = re.sub(r'(?<=[a-zA-Z0-9³²])/(?=[a-zA-Z0-9³²])', '__UNITSLASH__', text) # 單位斜線 (如 g/cm³)

    fullwidth_map = {
        ';': '；', '~': '～', ',': '，', '?': '？', '!': '！', ':': '：', '.': '．',
        '*': '×', '/': '÷', '+': '＋', '-': '－', '<': '＜', '>': '＞', '=': '＝'
    }
    for k, v in fullwidth_map.items():
        text = text.replace(k, v)

    # 移除數學運算符號前後無謂的空白
    text = re.sub(r'\s*([＋－×÷±＜＞＝])\s*', r'\1', text)

    # 2. 還原受保護的符號
    text = text.replace('__QDOT__', '.').replace('__DECDOT__', '.').replace('__UNITSLASH__', '/').replace('__PARALLEL__', '//')
    text = text.replace('（', '(').replace('）', ')')
    # 移除 ∠ 和 △ 後方的多餘空白（含全形空白與各種三角形符號）
    text = re.sub(r'([∠△▵▲▽])[\s\u3000]+', r'\1', text)
    return text

def set_run_font(run, size, is_english_mode, bold=False, use_title_font=False):
    run.font.size = Pt(size)
    run.bold = bold
    
    ch_font = DOC_TITLE_FONT if use_title_font else DOC_CHINESE_FONT
    if use_title_font:
        # 標題的英數字使用標楷體，避免王漢宗顏楷體繁缺少英數字元時退回 Times New Roman
        ascii_font = DOC_CHINESE_FONT
    elif is_english_mode:
        ascii_font = DOC_ENGLISH_FONT
    else:
        ascii_font = ch_font
    
    run.font.name = ascii_font
    rPr = run._element.get_or_add_rPr()
    f = rPr.get_or_add_rFonts()
    f.set(qn('w:eastAsia'), ch_font)
    f.set(qn('w:ascii'), ascii_font)
    f.set(qn('w:hAnsi'), ascii_font)
    f.set(qn('w:cs'), ch_font)

def remove_red_blue_marks(pil_img):
    """使用色彩過濾去除紅筆與藍筆手寫筆跡，保留黑色印刷線條"""
    if np is None:
        return pil_img
    img_rgb = pil_img.convert("RGB")
    arr = np.array(img_rgb, dtype=np.int16)
    r, g, b = arr[:, :, 0], arr[:, :, 1], arr[:, :, 2]
    
    # 偵測偏紅或偏藍的筆跡
    is_red = (r - np.maximum(g, b)) > 35
    is_blue = (b - np.maximum(r, g)) > 35
    mask = is_red | is_blue
    
    arr[mask] = [255, 255, 255]
    return Image.fromarray(arr.astype(np.uint8))

def refine_diagram_crop_cv2(pil_img):
    """使用 OpenCV 連通區塊分析，過濾周圍干擾文字行並精準貼齊圖形邊緣"""
    if cv2 is None or np is None:
        return pil_img
    img_np = np.array(pil_img.convert("RGB"))
    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    # 找出深色前景像素 (線條與文字)
    _, thresh = cv2.threshold(gray, 225, 255, cv2.THRESH_BINARY_INV)
    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(thresh, connectivity=8)
    if num_labels <= 1:
        return pil_img

    # 1. 找出構成幾何圖形/裝置圖的主體大區塊 (濾掉細碎的題幹文字與選項)
    large_boxes = []
    all_boxes = []
    for i in range(1, num_labels):
        x, y, w, h, area = stats[i]
        if area < 8:
            continue
        all_boxes.append((x, y, w, h, area))
        if w >= 35 or h >= 35 or area >= 250:
            large_boxes.append((x, y, x + w, y + h))

    if not large_boxes:
        return pil_img

    # 主體圖形的聯集邊界
    mx0 = min(b[0] for b in large_boxes)
    my0 = min(b[1] for b in large_boxes)
    mx1 = max(b[2] for b in large_boxes)
    my1 = max(b[3] for b in large_boxes)

    # 2. 納入緊鄰圖形邊緣的頂點字母/數字標示 (如 A, B, C, 4, 12)，排除較遠的題幹與選項文字
    margin = 28
    fx0, fy0, fx1, fy1 = mx0, my0, mx1, my1
    for x, y, w, h, area in all_boxes:
        bx0, by0, bx1, by1 = x, y, x + w, y + h
        if bx0 >= mx0 - margin and by0 >= my0 - margin and bx1 <= mx1 + margin and by1 <= my1 + margin:
            fx0 = min(fx0, bx0)
            fy0 = min(fy0, by0)
            fx1 = max(fx1, bx1)
            fy1 = max(fy1, by1)

    pad = 4
    h_img, w_img = img_np.shape[:2]
    fx0 = max(0, fx0 - pad)
    fy0 = max(0, fy0 - pad)
    fx1 = min(w_img, fx1 + pad)
    fy1 = min(h_img, fy1 + pad)
    if fx1 > fx0 and fy1 > fy0:
        return Image.fromarray(img_np[fy0:fy1, fx0:fx1])
    return pil_img

def extract_diagrams_from_page_cv2(pil_img, start_id=1):
    """使用 OpenCV 先行框出整頁中的幾何圖形/裝置圖並標上 [PIC: id] 編號"""
    if cv2 is None or np is None:
        return pil_img, {}
    img_np = np.array(pil_img.convert("RGB"))
    h_img, w_img = img_np.shape[:2]
    # 先濾除紅藍手寫筆跡，再做灰階與膨脹，避免手寫大括號或筆記被誤框為圖形
    cleaned_full = remove_red_blue_marks(pil_img)
    gray = cv2.cvtColor(np.array(cleaned_full), cv2.COLOR_RGB2GRAY)
    _, thresh = cv2.threshold(gray, 220, 255, cv2.THRESH_BINARY_INV)

    # 膨脹以將鄰近的圖形線條與頂點字母合併為候選區塊
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (14, 14))
    dilated = cv2.dilate(thresh, kernel, iterations=1)
    num_labels, _, stats, _ = cv2.connectedComponentsWithStats(dilated, connectivity=8)

    candidates = []
    for i in range(1, num_labels):
        x, y, w, h, area = stats[i]
        # 過濾掉單行文字、細長括號與整頁外框
        if w >= 90 and h >= 75 and area >= 1500 and not (w > w_img * 0.92 and h > h_img * 0.85):
            candidates.append((y, x, x + w, y + h))

    # 由上到下、由左到右排序
    candidates.sort()

    annotated = img_np.copy()
    extracted_pics = {}
    cur_id = start_id

    for y0, x0, x1, y1 in candidates:
        # 檢查該區域內是否真的有構成圖形的「大元件」（排除兩行文字黏在一起或純中文字）
        roi_thresh = thresh[y0:y1, x0:x1]
        n_sub, _, sub_stats, _ = cv2.connectedComponentsWithStats(roi_thresh, connectivity=8)
        has_diagram_stroke = False
        for j in range(1, n_sub):
            sw, sh, sarea = sub_stats[j, cv2.CC_STAT_WIDTH], sub_stats[j, cv2.CC_STAT_HEIGHT], sub_stats[j, cv2.CC_STAT_AREA]
            if sw >= 55 or sh >= 55 or sarea >= 450:
                has_diagram_stroke = True
                break
        if not has_diagram_stroke:
            continue

        pad = 6
        cx0, cy0 = max(0, x0 - pad), max(0, y0 - pad)
        cx1, cy1 = min(w_img, x1 + pad), min(h_img, y1 + pad)
        cropped = pil_img.crop((cx0, cy0, cx1, cy1))
        cleaned = remove_red_blue_marks(cropped)
        cleaned = refine_diagram_crop_cv2(cleaned)

        bio = io.BytesIO()
        cleaned.save(bio, format="PNG")
        extracted_pics[cur_id] = bio.getvalue()

        # 在給 AI 看的頁面複本上畫出紅框與醒目編號
        cv2.rectangle(annotated, (cx0, cy0), (cx1, cy1), (220, 30, 30), 3)
        label = f"PIC:{cur_id}"
        cv2.rectangle(annotated, (cx0, max(0, cy0 - 28)), (cx0 + 115, cy0), (220, 30, 30), -1)
        cv2.putText(annotated, label, (cx0 + 5, max(20, cy0 - 6)), cv2.FONT_HERSHEY_SIMPLEX, 0.75, (255, 255, 255), 2)
        cur_id += 1

    return Image.fromarray(annotated), extracted_pics

def add_floating_picture(paragraph, img_bytes, width_cm=4.5):
    """將圖片插入段落並設定為『矩形文繞圖』(靠右對齊，文字自動避開圖片)"""
    run = paragraph.add_run()
    inline_shape = run.add_picture(io.BytesIO(img_bytes), width=Cm(width_cm))
    inline = inline_shape._inline
    
    # 取得尺寸與圖片 XML
    extent = inline.find(qn('wp:extent'))
    cx = extent.get('cx')
    cy = extent.get('cy')
    docPr = inline.find(qn('wp:docPr'))
    graphic = inline.find(qn('a:graphic'))
    
    anchor_xml = f"""
    <wp:anchor {nsdecls('wp')} distT="0" distB="0" distL="114300" distR="114300"
               simplePos="0" relativeHeight="251658240" behindDoc="0"
               locked="0" layoutInCell="1" allowOverlap="1">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="margin">
            <wp:align>right</wp:align>
        </wp:positionH>
        <wp:positionV relativeFrom="paragraph">
            <wp:posOffset>0</wp:posOffset>
        </wp:positionV>
        <wp:extent cx="{cx}" cy="{cy}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:wrapNone/>
        <wp:docPr id="{docPr.get('id')}" name="{docPr.get('name')}"/>
        <wp:cNvGraphicFramePr>
            <a:graphicFrameLocks {nsdecls('a')} noChangeAspect="1"/>
        </wp:cNvGraphicFramePr>
    </wp:anchor>
    """
    anchor = parse_xml(anchor_xml)
    anchor.append(graphic)
    
    drawing = inline.getparent()
    drawing.replace(inline, anchor)

def add_formatted_table(doc, table_lines, indent_pt, is_english_mode):
    """將 Markdown 表格轉為 Word 表格，並對齊題目選項的縮排"""
    rows_data = []
    for tline in table_lines:
        tline = tline.strip()
        # 跳過 Markdown 表格分隔線 (如 |---|---|)
        if re.match(r'^\|[\s\-:|]+\|$', tline):
            continue
        cells = [c.strip() for c in tline.strip('|').split('|')]
        rows_data.append(cells)

    if not rows_data:
        return

    num_cols = max(len(r) for r in rows_data)
    table = doc.add_table(rows=len(rows_data), cols=num_cols)
    table.style = 'Table Grid'

    # 設定表格左縮排，對齊該題選項 (indent_pt 單位為 Pt，1 Pt = 20 dxa)
    if indent_pt > 0:
        tblPr = table._tbl.tblPr
        tblInd = OxmlElement('w:tblInd')
        tblInd.set(qn('w:w'), str(int(indent_pt * 20)))
        tblInd.set(qn('w:type'), 'dxa')
        tblPr.append(tblInd)

    for r_idx, row_vals in enumerate(rows_data):
        row = table.rows[r_idx]
        for c_idx in range(num_cols):
            val = row_vals[c_idx] if c_idx < len(row_vals) else ""
            val = custom_format_text(val)
            cell = row.cells[c_idx]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(val)
            set_run_font(run, size=12, is_english_mode=is_english_mode, bold=(r_idx == 0))

def create_word_document(text, is_english_mode, extracted_pics=None, answers_text=None, ans_title=None):
    doc = Document()
    settings = doc.settings.element
    compat = settings.find(qn('w:compat')) or OxmlElement('w:compat')
    settings.append(compat)
    ult = OxmlElement('w:ulTrailSpace')
    compat.append(ult)

    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(1)

    header = "考卷排版結果"
    h_match = re.search(r"HEADER:\s*(.*)", text)
    if h_match:
        header = h_match.group(1).strip()
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    # A4 寬度 21cm，左右邊界各 1cm，內文寬度為 19cm，正中央為 9.5cm
    tab_stops = p.paragraph_format.tab_stops
    if "|" in header:
        tab_stops.add_tab_stop(Cm(9.5), WD_TAB_ALIGNMENT.CENTER)
        tab_stops.add_tab_stop(Cm(19), WD_TAB_ALIGNMENT.RIGHT)
        parts = header.split("|")
        display_header = f"{parts[0].strip()}\t{parts[1].strip()}\t\xa0"
    else:
        clean_header = header.replace(';', '；').replace('~', '～').rstrip()
        # 自動偵測「學校」與「第X次段考試題」，將考試名稱置中於 9.5cm
        m_hdr = re.match(r'^(.*?)\s+((?:第[一二三]次|期中|期末).*)$', clean_header)
        tab_stops.add_tab_stop(Cm(9.5), WD_TAB_ALIGNMENT.CENTER)
        tab_stops.add_tab_stop(Cm(19), WD_TAB_ALIGNMENT.RIGHT)
        if m_hdr:
            display_header = f"{m_hdr.group(1).strip()}\t{m_hdr.group(2).strip()}\t\xa0"
        else:
            display_header = f"{clean_header}\t\xa0"

    run_h = p.add_run(display_header)
    run_h.underline = WD_UNDERLINE.SINGLE
    set_run_font(run_h, size=16, is_english_mode=is_english_mode, bold=False, use_title_font=True)

    # 萃取內文並防呆清除殘留標籤
    content_match = re.search(r"CONTENT_START(.*?)CONTENT_END", text, re.DOTALL)
    if content_match:
        main_content = content_match.group(1).strip()
    else:
        main_content = re.sub(r"HEADER:.*\n?", "", text)
        main_content = main_content.replace("CONTENT_START", "").replace("CONTENT_END", "").strip()

    current_q_indent = 0  # 追蹤目前題目的縮排寬度，讓延續行(如對話)能對齊
    current_sec_is_open = False  # 追蹤目前大題是否為填充題/計算題/應用題
    current_sec_is_calc = False  # 追蹤目前大題是否為計算/應用/證明題（空 6 行）
    pending_blank_lines = 0  # 延後到小題結束後才插入的空白行數
    # 保險機制：強制將同一行內的選項 (A)(B)(C)(D) 或 ①②③④ 拆成獨立的新行
    main_content = re.sub(r'[ \t]+(?=(?:\([A-Da-d]\)|[①②③④]))', '\n', main_content)
    lines = main_content.split('\n')
    idx_line = 0

    while idx_line < len(lines):
        line = lines[idx_line]
        # 偵測連續的 Markdown 表格行 (以 | 開頭與結尾)
        if line.strip().startswith('|') and '|' in line.strip()[1:]:
            table_buf = []
            while idx_line < len(lines) and lines[idx_line].strip().startswith('|') and '|' in lines[idx_line].strip()[1:]:
                table_buf.append(lines[idx_line])
                idx_line += 1
            add_formatted_table(doc, table_buf, current_q_indent if current_q_indent > 0 else 48, is_english_mode)
            continue
        idx_line += 1
        # 擷取 OpenCV 預先框選的圖片編號標記 [PIC: id]
        pic_ids = []
        if extracted_pics:
            for m_pic in re.finditer(r'\[PIC:\s*(\d+)\s*\]', line):
                pic_ids.append(int(m_pic.group(1)))
        line = re.sub(r'\[PIC:[^\]]*\]', '', line)

        line, math_list = extract_math_placeholders(line)
        line = custom_format_text(line)
        # 移除數學式前後與中文/全形標點之間的多餘空白
        line = re.sub(r'\s+(__MATH_\d+__)', r'\1', line)
        line = re.sub(r'(__MATH_\d+__)\s+', r'\1', line)
        if not line.strip() and not pic_ids:
            continue

        # 安全過濾網：自動剔除殘留的頁碼、試題結束、答案卡/答案卷作答規則提示與班級座號欄
        clean_chk = line.strip()
        if (
            re.search(r'第\s*\d+\s*頁\s*[，,]\s*共\s*\d+\s*頁', clean_chk)
            or '試題結束' in clean_chk
            or ('答案卡' in clean_chk and ('2B' in clean_chk or '鉛筆' in clean_chk or '畫記' in clean_chk))
            or ('答案卷' in clean_chk and ('作答' in clean_chk or '黑筆' in clean_chk or '試題卷共' in clean_chk))
            or ('班' in clean_chk and '座號' in clean_chk and '姓名' in clean_chk)
        ):
            continue
        
        match_sec = re.match(r'^\s*(?:[一二三四五六七八九十]+[、.．]|【題組】|題組)', line)
        match_q = re.match(r'^\s*(\d+)\s*[.、．]\s*(.*)', line)
        match_opt = re.match(r'^\s*\([A-Da-d]\)\s*(.*)', line)

        if (match_sec or match_q) and pending_blank_lines > 0:
            for _ in range(pending_blank_lines):
                blank_p = doc.add_paragraph()
                blank_p.paragraph_format.line_spacing = 1.15
                blank_p.paragraph_format.space_after = Pt(2)
            pending_blank_lines = 0

        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(2)

        is_open_question = False

        if match_sec:
            # 大題標題（如 二、填充題）靠最左對齊，並重置題目縮排
            current_q_indent = 0
            line = line.strip()
            p.paragraph_format.left_indent = Pt(0)
            p.paragraph_format.first_line_indent = Pt(0)
            current_sec_is_open = any(k in line for k in ["填充", "計算", "應用", "非選", "作圖", "證明"])
            current_sec_is_calc = any(k in line for k in ["計算", "應用", "證明"])

        elif match_q:
            q_num = int(match_q.group(1))
            q_content = match_q.group(2).strip()
            
            is_drawing = any(k in q_content for k in ["作圖", "畫出", "計算", "證明"])
            is_fill = "＿＿" in q_content or "{BLANK}" in q_content
            need_bracket = (not current_sec_is_open) and ("(   )" in line or (not is_drawing and not is_fill))

            if need_bracket:
                if re.match(r'^\(\s*\)', q_content):
                    q_content = re.sub(r'^\(\s*\)\s*', '(   )', q_content)
                else:
                    q_content = f"(   ){q_content}"

            if q_num <= 9:
                line = f" {q_num}.{q_content}"
            else:
                line = f"{q_num}.{q_content}"
            
            current_q_indent = 48 if need_bracket else 18
            p.paragraph_format.left_indent = Pt(current_q_indent)
            p.paragraph_format.first_line_indent = -Pt(current_q_indent)
            if current_sec_is_open or (not need_bracket):
                is_open_question = True

        elif match_opt:
            line = line.strip()
            p.paragraph_format.left_indent = Pt(current_q_indent if current_q_indent > 0 else 48)

        else:
            # 延續行 / 對話行（例如英文題目的 Ryan : ...），對齊上一題的題幹起點
            if current_q_indent > 0:
                p.paragraph_format.left_indent = Pt(current_q_indent)

        if line.strip():
            add_text_with_equations(p, line, size=12, is_english_mode=is_english_mode, math_list=math_list)

        # 若為填充題、計算題或應用題，記錄空白行數，待該題（含小題）結束後再插入
        target_img_p = p
        if is_open_question:
            pending_blank_lines = 6 if (current_sec_is_calc or any(k in line for k in ["計算", "應用", "證明"])) else 4

        # 若該題有配圖，直接從 OpenCV 預先裁切好的圖庫取出並貼於右側
        if pic_ids and extracted_pics:
            for pid in pic_ids:
                if pid in extracted_pics:
                    add_floating_picture(target_img_p, extracted_pics[pid], width_cm=4.5)

    if pending_blank_lines > 0:
        for _ in range(pending_blank_lines):
            blank_p = doc.add_paragraph()
            blank_p.paragraph_format.line_spacing = 1.15
            blank_p.paragraph_format.space_after = Pt(2)

    # 若有辨識出解答，於考卷最下方附上【檔名 ANS】緊湊解答
    if answers_text and answers_text.strip():
        clean_ans_hdr = ans_title.strip() if ans_title else re.sub(r'\s+', ' ', header).strip()
        ans_title_p = doc.add_paragraph()
        ans_title_p.paragraph_format.space_before = Pt(12)
        ans_title_p.paragraph_format.space_after = Pt(2)
        run_at = ans_title_p.add_run(f"【{clean_ans_hdr} ANS】")
        set_run_font(run_at, size=12, is_english_mode=is_english_mode, bold=False)

        for ans_line in answers_text.strip().split('\n'):
            ans_line = ans_line.strip()
            if not ans_line:
                continue
            ans_line, ans_math = extract_math_placeholders(ans_line)
            ans_line = custom_format_text(ans_line)
            ans_p = doc.add_paragraph()
            ans_p.paragraph_format.line_spacing = 1.15
            ans_p.paragraph_format.space_after = Pt(2)
            add_text_with_equations(ans_p, ans_line, size=12, is_english_mode=is_english_mode, math_list=ans_math)

    # 將 Word 檔案存進記憶體中以便網頁下載
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()


def get_global_usage(webhook_url):
    if not webhook_url:
        return None
    try:
        res = requests.get(webhook_url, timeout=5)
        if res.status_code == 200:
            return res.json()
    except Exception:
        pass
    return None

def log_usage_to_gsheet(webhook_url, filename, pages_count, model_id, tokens):
    if not webhook_url:
        return
    try:
        payload = {
            "filename": filename,
            "pages": pages_count,
            "model": model_id,
            "tokens": tokens
        }
        requests.post(webhook_url, json=payload, timeout=5)
    except Exception:
        pass

def call_api_with_retry(parts, model_id, api_key):
    models_to_try = [model_id] + [m for m in ["gemini-3.8-flash", "gemini-3.6-flash", "gemini-3.5-flash-lite", "gemini-3.1-flash-lite", "gemini-2.5-flash-lite", "gemini-2.5-flash"] if m != model_id]
    
    for mid in models_to_try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{mid}:generateContent?key={api_key}"
        for attempt in range(3):
            try:
                res = requests.post(
                    url, 
                    json={"contents": [{"parts": parts}], "generationConfig": {"temperature": 0.1, "maxOutputTokens": 65536}}, 
                    timeout=600, 
                    verify=False
                )
                if res.status_code == 200:
                    data = res.json()
                    text_out = data['candidates'][0]['content']['parts'][0]['text']
                    usage = data.get('usageMetadata', {})
                    return text_out, usage, mid
                elif res.status_code == 429:
                    time.sleep(15)
                elif res.status_code == 503:
                    time.sleep(8)
                else:
                    break
            except Exception:
                break
    return None, {}, None


app = FastAPI(title="考卷排版神器 Pro")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.middleware("http")
async def allow_google_sites_iframe(request: Request, call_next):
    response = await call_next(request)
    response.headers["Content-Security-Policy"] = "frame-ancestors 'self' https://sites.google.com https://*.googleusercontent.com;"
    if "X-Frame-Options" in response.headers:
        del response.headers["X-Frame-Options"]
    return response

# 環境變數設定（與原本 Streamlit Secrets 相同名稱）
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
APP_PASSWORD = os.environ.get("APP_PASSWORD", "").strip()
GSHEET_WEBHOOK_URL = os.environ.get("GSHEET_WEBHOOK_URL", "").strip()

TASKS = {}
GENERATED_FILES = {}
SESSION_STATS = {"request_count": 0, "total_tokens": 0}
HISTORY_LIST = []


def _build_app_icon(size: int) -> bytes:
    """讀取 icon.png，徹底去除外圍白邊並自動裁切留白放大主體比例"""
    if os.path.exists("icon.png"):
        try:
            img = Image.open("icon.png").convert("RGBA")
            if np is not None and cv2 is not None:
                arr = np.array(img, dtype=np.int16)
                r, g, b = arr[:, :, 0], arr[:, :, 1], arr[:, :, 2]
                white_mask = (r > 215) & (g > 215) & (b > 215)
                h, w = white_mask.shape
                ff_mask = np.zeros((h + 2, w + 2), np.uint8)
                flood = white_mask.astype(np.uint8) * 255
                for seed in [(0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1)]:
                    if flood[seed[1], seed[0]] == 255:
                        cv2.floodFill(flood, ff_mask, seed, 128)
                bg_mask = (flood == 128).astype(np.uint8) * 255
                kernel = np.ones((3, 3), np.uint8)
                bg_mask = cv2.dilate(bg_mask, kernel, iterations=1)
                alpha = 255 - bg_mask
                alpha = cv2.GaussianBlur(alpha, (3, 3), 0)
                arr[:, :, 3] = alpha
                img = Image.fromarray(arr.astype(np.uint8))

            bbox = img.getbbox()
            if bbox:
                img = img.crop(bbox)
            pad_ratio = 0.02
            inner_size = max(1, int(size * (1 - pad_ratio * 2)))
            img = img.resize((inner_size, inner_size), Image.Resampling.LANCZOS)
            canvas = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            offset = (size - inner_size) // 2
            canvas.paste(img, (offset, offset), img)

            bio = io.BytesIO()
            canvas.save(bio, format="PNG")
            return bio.getvalue()
        except Exception:
            pass

    from PIL import ImageDraw
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    pad = int(size * 0.04)
    draw.rounded_rectangle([pad, pad, size - pad, size - pad], radius=int(size * 0.22), fill=(140, 94, 60, 255))
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    return bio.getvalue()

@app.get("/manifest.json")
async def get_manifest():
    return {
        "name": "考卷排版神器 Pro",
        "short_name": "考卷排版神器",
        "start_url": "/",
        "display": "standalone",
        "background_color": "#F4F1EA",
        "theme_color": "#8C5E3C",
        "icons": [
            {"src": "/icon-192.png", "sizes": "192x192", "type": "image/png", "purpose": "any maskable"},
            {"src": "/icon-512.png", "sizes": "512x512", "type": "image/png", "purpose": "any maskable"}
        ]
    }

@app.get("/icon-192.png")
async def get_icon_192():
    return Response(content=_build_app_icon(192), media_type="image/png")

@app.get("/icon-512.png")
async def get_icon_512():
    return Response(content=_build_app_icon(512), media_type="image/png")

@app.get("/apple-touch-icon.png")
async def get_apple_icon():
    return Response(content=_build_app_icon(180), media_type="image/png")

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    with open("index.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

@app.get("/api/config")
async def get_config():
    return {
        "requires_password": bool(APP_PASSWORD),
        "has_api_key": bool(GEMINI_API_KEY),
        "models": VALID_MODELS,
    }

@app.post("/api/login")
async def login(request: Request):
    if not APP_PASSWORD:
        return {"ok": True}
    body = await request.json()
    pwd = body.get("password", "").strip()
    if pwd == APP_PASSWORD:
        return {"ok": True}
    raise HTTPException(status_code=401, detail="密碼錯誤")

@app.get("/api/usage")
async def get_usage():
    global_stats = get_global_usage(GSHEET_WEBHOOK_URL) if GSHEET_WEBHOOK_URL else None
    if global_stats is not None:
        disp_count = global_stats.get("today_requests", global_stats.get("total_requests", 0))
        disp_tokens = global_stats.get("today_tokens", global_stats.get("total_tokens", 0))
    else:
        disp_count = SESSION_STATS["request_count"]
        disp_tokens = SESSION_STATS["total_tokens"]
    return {
        "count": disp_count,
        "tokens": disp_tokens,
        "limit": 500000,
        "history": HISTORY_LIST[:20],
    }

@app.post("/api/pdf-info")
async def pdf_info(file: UploadFile = File(...)):
    pdf_bytes = await file.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pages = len(doc)
    doc.close()
    return {"filename": file.filename, "total_pages": total_pages}

def _run_parse_job(
    task_id: str,
    pdf_bytes: bytes,
    filename: str,
    page_input: str,
    selected_model: str,
    is_english_mode: bool,
    enable_image_crop: bool,
    enable_omml_math: bool,
    extract_answers: bool,
):
    try:
        if not GEMINI_API_KEY:
            TASKS[task_id].update({"status": "error", "pct": 0, "message": "❌ 系統尚未設定 GEMINI_API_KEY 環境變數"})
            return

        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        indices = parse_page_selection(page_input, len(doc))

        math_rule_txt = (
            "**僅在分數、根號、線段本體時使用標準 LaTeX 語法並以單一 `$` 包夾**（例如 `$\\frac{1}{2}$`、`$\\sqrt{3}$`、`$\\overline{AB}$`）。**等號、數字與其他文字絕對不要包進 `$` 裡面**（例如寫成 `$\\overline{BC}$＝1`，不可寫成 `$\\overline{BC}＝1$`）。單純的英文字母（如 `ABCD`、`x`）、數字比例（如 `5：6：3`）或一般中文敘述請維持純文字，**不要**放進 `$` 裡面。"
            if enable_omml_math else
            "**絕對不要使用 LaTeX 語法或 `$` 符號**（例如不要寫 `$x^2$`、`\\frac{1}{2}`、`\\overline{AB}`、`\\bot`），線段直接寫 `AB`、分數直接寫 `1/2`、垂直寫 `⊥`、平行寫 `//`，請直接輸出純文字與 Unicode 符號（如 `x²`、`1/2`、`×`、`÷`）。數學算式中的 `+ - × ÷ ± < > = :` 請使用**全形符號**（`＋ － × ÷ ± ＜ ＞ ＝ ：`）。"
        )
        prompt = """
            請分析考卷圖片並嚴格遵守以下規則進行逐字轉錄與排版：
            0. **最高原則（過濾無關雜訊、手寫與精確轉錄）**：
               - **完全忽略所有手寫筆跡與解答詳解**（包含題號旁的紅筆/藍筆/鉛筆答案、選項上的圈選或刪除線、空白處的計算過程，以及答案卷上附帶的「評分標準／參考算式／詳解」），只辨識原本的**印刷體題目**。
               - **完全忽略所有考試規則說明與頁首頁尾雜訊**（例如：`※試題卷共 X 頁...`、`請使用 2B 鉛筆畫在答案卡上`、`請使用黑筆在答案卷上作答`、`〔第 X 頁，共 X 頁〕`、`「試題結束」`、答案卷頂端的學校/範圍/班級座號姓名欄），這些提示字句一律不要輸出。
               - **答案卷／作答卷處理原則**：不需要整張轉錄答案卷！答案卷上的「選擇題／填充題填答格子表格」（用來給學生填寫答案的空白表格）**絕對不要輸出**；只有當答案卷上印有**計算題、證明題、作圖題或應用題的完整題目敘述**時，才需要把該大題標題與題目（不含解答算式）轉錄出來。
               - 必須**逐字精確轉錄**印刷體題目與選項，絕對不可自行改寫、腦補或省略。
               - 保留小數點（如 `100.0`）與單位斜線（如 `g/cm³`）的原貌。
               - {math_rule_txt}
            1. HEADER: 強制重組標題為 `[年份]([學期])[學校]                  [考試名稱]                            `。
               - `[年份]` 必須是民國年數字（如 `115`、`114`、`113`），**絕對不可寫成縣市名稱（如新北、台北）**。
               - [考試名稱]後面的空格一定要加上
               - 絕對不要加上 <u> 或 ** 等任何 HTML 或 Markdown 標籤
               - 學校只需要名稱前兩個字（如 `中和`），考試只需要標示第一次、第二次、或第三次段考試題並且要在整行標題包含後面空格加上底線，然後要靠左對齊。
               - 範例：`115(一)中和                  第一次段考試題                            `。
            2. **選擇題 / 是非題**：
               - 必須在題號後加上括號 `(   )`，且括號後面**不要加空格**，直接接題目文字。
               - 範例：`1.(   )這是題目`。
            3. **填充題 / 作圖題 / 計算題**：
               - 題號後 **不要** 加括號，直接空格接題目。
               - 範例：`2. 請畫出三角形` 或 `3. ＿＿＿`。
            4. **選項處理**：
               - 若選項是 1,2,3,4 或 (1)(2)(3)(4)，請轉為 ①②③④。
               - 若選項是 (A)(B)(C)(D)，保留原樣。
               - **務必將每個選項 (A)(B)(C)(D) 或 ①②③④ 放在新的一行**。
            5. 重要：忽略圖片、插圖，不要輸出圖片描述；但若考卷或題目附有**文字/數據表格、分數對照表或配分表**，請務必使用標準 Markdown 表格語法（`| 欄位1 | 欄位2 |`）完整輸出表格內容。
            6. 如果找不到標題，請把第一行文字當作標題。
            格式回傳（務必完整包含頭尾標記）：
            HEADER: [標題]
            CONTENT_START
            ...
            CONTENT_END
            """

        if enable_image_crop:
            prompt += """
            7. **配圖編號配對**：
               - 圖片中已經用紅框標記並寫上 `PIC:編號`（例如 `PIC:1`、`PIC:2`）。
               - 若某題目附有該幾何圖形、實驗裝置圖或圖表，請在該題文字的最後面加上 `[PIC: 編號]`（例如 `[PIC: 1]`）。
               - **若紅框內是「手寫筆跡／手寫計算過程」、「純文字」、「空白作答欄框線」或是「已經輸出為 Markdown 表格的文字/數據表格」，請絕對不要輸出 `[PIC: 編號]`**。只針對「印刷體的幾何圖形、實驗裝置圖或圖表」輸出 `[PIC: 編號]`。
               - 若該題沒有配圖，則不要輸出 `[PIC: ...]`。
        """

        prompt = prompt.replace("{math_rule_txt}", math_rule_txt)

        batches = [indices[i:i + 2] for i in range(0, len(indices), 2)]
        total_steps = len(batches) + (1 if extract_answers else 0) + 1
        all_extracted_pics = {}
        next_pic_id = 1
        merged_contents = []
        merged_answers = []
        final_header = "考卷排版結果"
        total_used_tokens = 0
        used_models = []
        all_ok = True

        for b_idx, batch in enumerate(batches):
            if TASKS[task_id].get("stop"):
                doc.close()
                TASKS[task_id].update({"status": "stopped", "message": "⏹ 已手動停止生成"})
                return

            pct = max(8, int(((b_idx + 0.5) / total_steps) * 100))
            TASKS[task_id].update({
                "pct": pct,
                "message": f"⏳ 正在解析... {b_idx + 1}/{total_steps} 辨識題目與數學公式..."
            })

            parts = [{"text": prompt}]
            for idx in batch:
                pix = doc[idx].get_pixmap(dpi=180)
                jpg_bytes = pix.tobytes("jpeg")
                if enable_image_crop:
                    pil_page = Image.open(io.BytesIO(jpg_bytes)).convert("RGB")
                    annotated_page, page_pics = extract_diagrams_from_page_cv2(pil_page, start_id=next_pic_id)
                    next_pic_id += len(page_pics)
                    all_extracted_pics.update(page_pics)
                    buf = io.BytesIO()
                    annotated_page.save(buf, format="JPEG", quality=92)
                    jpg_bytes = buf.getvalue()
                b64 = base64.b64encode(jpg_bytes).decode("utf-8")
                parts.append({"inline_data": {"mime_type": "image/jpeg", "data": b64}})

            txt_b, usage_b, actual_mid_b = call_api_with_retry(parts, selected_model, GEMINI_API_KEY)
            del parts
            gc.collect()
            if not txt_b:
                all_ok = False
                break
            if actual_mid_b and actual_mid_b not in used_models:
                used_models.append(actual_mid_b)
            total_used_tokens += usage_b.get("totalTokenCount", 0)
            if b_idx == 0:
                hm = re.search(r"HEADER:\s*(.*)", txt_b)
                if hm:
                    final_header = hm.group(1).strip()
            am = re.search(r"ANSWERS_START(.*?)ANSWERS_END", txt_b, re.DOTALL)
            if am and am.group(1).strip():
                merged_answers.append(am.group(1).strip())
            txt_b_clean = re.sub(r"ANSWERS_START.*?ANSWERS_END", "", txt_b, flags=re.DOTALL)
            cm = re.search(r"CONTENT_START(.*?)CONTENT_END", txt_b_clean, re.DOTALL)
            if cm:
                merged_contents.append(cm.group(1).strip())
            else:
                clean_b = re.sub(r"HEADER:.*\n?", "", txt_b_clean).replace("CONTENT_START", "").replace("CONTENT_END", "").strip()
                merged_contents.append(clean_b)

        if all_ok and extract_answers:
            if TASKS[task_id].get("stop"):
                doc.close()
                TASKS[task_id].update({"status": "stopped", "message": "⏹ 已手動停止生成"})
                return
            pct = int(((len(batches) + 0.5) / total_steps) * 100)
            TASKS[task_id].update({
                "pct": pct,
                "message": f"⏳ 正在解析... {len(batches) + 1}/{total_steps} 辨識考卷參考解答..."
            })
            ans_prompt = """請只專心辨識這幾頁考卷上的「正確解答」（包含題號旁手寫紅筆／藍筆／鉛筆寫的答案、括號內的選項字母、圈選的選項，或考卷附帶的解答區）。
            **不要輸出題目文字**，請嚴格依照以下緊湊格式逐大題輸出：
            - 選擇題：只列選項字母，每 5 題空一格，例如 `【一、ADCBC ACBDA】`。
            - 填充題／計算題：只列最後答案（不寫算式過程），題號用 `(1)` 或 `(2-1)` 緊接答案、各題之間空一格，例如 `【二、(1)0 (2-1)$\\frac{\\sqrt{2}}{5}$ (2-2)$2\\sqrt{2}-\\sqrt{3}$】`。
            請一題一題看清楚括號內手寫字母（特別區分 D 與 B、A 與 C），絕對不可猜測或跳號。若完全沒有解答則輸出空字串。"""
            ans_parts = [{"text": ans_prompt}]
            for idx in indices:
                pix = doc[idx].get_pixmap(dpi=180)
                b64 = base64.b64encode(pix.tobytes("jpeg")).decode("utf-8")
                ans_parts.append({"inline_data": {"mime_type": "image/jpeg", "data": b64}})
            ans_out, ans_usage, actual_mid_ans = call_api_with_retry(ans_parts, selected_model, GEMINI_API_KEY)
            del ans_parts
            gc.collect()
            if ans_out and ans_out.strip():
                if actual_mid_ans and actual_mid_ans not in used_models:
                    used_models.append(actual_mid_ans)
                total_used_tokens += ans_usage.get("totalTokenCount", 0)
                merged_answers = [ans_out.strip()]

        doc.close()
        del pdf_bytes
        gc.collect()
        txt = f"HEADER: {final_header}\nCONTENT_START\n" + "\n".join(merged_contents) + "\nCONTENT_END" if (all_ok and merged_contents) else None

        if not txt:
            TASKS[task_id].update({"status": "error", "pct": 0, "message": "❌ 轉換失敗，請更換模型或稍後重試"})
            return

        TASKS[task_id].update({
            "pct": 92,
            "message": f"⏳ 正在解析... {total_steps}/{total_steps} 產生 Word (.docx) 精緻排版..."
        })

        SESSION_STATS["request_count"] += 1
        SESSION_STATS["total_tokens"] += total_used_tokens
        actual_model_str = ", ".join(used_models) if used_models else selected_model
        if GSHEET_WEBHOOK_URL:
            log_usage_to_gsheet(GSHEET_WEBHOOK_URL, filename, len(indices), actual_model_str, total_used_tokens)

        base_name = os.path.splitext(filename)[0]
        answers_txt = "\n".join(merged_answers) if (extract_answers and merged_answers) else None
        docx_bytes = create_word_document(
            txt,
            is_english_mode,
            extracted_pics=all_extracted_pics if enable_image_crop else None,
            answers_text=answers_txt,
            ans_title=base_name,
        )
        out_name = f"{base_name}.docx"
        now_str = datetime.now(timezone(timedelta(hours=8))).strftime("%H:%M:%S")
        file_id = str(uuid.uuid4())
        GENERATED_FILES[file_id] = {"name": out_name, "bytes": docx_bytes}

        item_meta = {
            "id": file_id,
            "name": out_name,
            "time": now_str,
            "model": actual_model_str,
            "pages": len(indices),
        }
        HISTORY_LIST.insert(0, item_meta)
        if len(HISTORY_LIST) > 10:
            old = HISTORY_LIST.pop()
            GENERATED_FILES.pop(old["id"], None)

        TASKS[task_id].update({
            "status": "done",
            "pct": 100,
            "message": "✅ 排版完成！已自動更新至右側最新產出",
            "result": item_meta,
        })
    except Exception as e:
        TASKS[task_id].update({"status": "error", "pct": 0, "message": f"❌ 處理發生錯誤：{str(e)}"})
    finally:
        gc.collect()

@app.post("/api/parse")
async def start_parse(
    file: UploadFile = File(...),
    page_range: str = Form(""),
    model: str = Form("gemini-3.8-flash"),
    is_english_mode: bool = Form(False),
    enable_image_crop: bool = Form(False),
    enable_omml_math: bool = Form(False),
    extract_answers: bool = Form(False),
):
    pdf_bytes = await file.read()
    task_id = str(uuid.uuid4())
    TASKS[task_id] = {"status": "running", "pct": 5, "message": "⏳ 正在準備 PDF 頁面影像...", "stop": False}
    t = threading.Thread(
        target=_run_parse_job,
        args=(
            task_id,
            pdf_bytes,
            file.filename or "考卷排版結果.pdf",
            page_range,
            model,
            is_english_mode,
            enable_image_crop,
            enable_omml_math,
            extract_answers,
        ),
        daemon=True,
    )
    t.start()
    return {"task_id": task_id}

@app.get("/api/status/{task_id}")
async def get_status(task_id: str):
    task = TASKS.get(task_id)
    if not task:
        raise HTTPException(status_code=404, detail="找不到任務")
    return task

@app.post("/api/stop/{task_id}")
async def stop_task(task_id: str):
    task = TASKS.get(task_id)
    if task:
        task["stop"] = True
    return {"ok": True}

@app.get("/api/download/{file_id}")
async def download_file(file_id: str):
    item = GENERATED_FILES.get(file_id)
    if not item:
        raise HTTPException(status_code=404, detail="檔案已過期或不存在")
    encoded_name = quote(item["name"])
    return Response(
        content=item["bytes"],
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_name}"},
    )
