"""
Trộn Đề Word Online - AIOMT Premium (Dashboard Layout)
Author: Phan Trường Duy - THPT Minh Đức
"""

import os
import sys
import subprocess
import time

# ==================== TỰ ĐỘNG CÀI ĐẶT THƯ VIỆN ====================
def install_libs():
    required_packages = {
        "pandas": "pandas",
        "xlsxwriter": "XlsxWriter",
        "openpyxl": "openpyxl",
        "streamlit": "streamlit"
    }
    for lib, package_name in required_packages.items():
        try:
            __import__(lib)
        except ImportError:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])

install_libs()

# ==================== IMPORTS ====================
import streamlit as st
import re
import random
import zipfile
import io
import pandas as pd
from xml.dom import minidom

# ==================== CẤU HÌNH TRANG & CSS ====================

st.set_page_config(
    page_title="AIOMT Premium",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS tối ưu không gian cho màn hình Laptop
st.markdown("""
<style>
    /* Giảm lề trên cùng */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 1rem !important;
    }
    
    /* Header gọn hơn */
    .main-header {
        background: linear-gradient(90deg, #006266, #009432);
        padding: 1rem;
        border-radius: 8px;
        color: white;
        text-align: center;
        margin-bottom: 1rem;
    }
    .main-header h1 { font-size: 1.8rem; margin: 0; }
    .main-header p { font-size: 0.9rem; margin: 0; opacity: 0.9; }

    /* Card style */
    .stCard {
        background-color: white;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        border: 1px solid #eee;
    }

    /* Error log box */
    .error-log {
        max-height: 250px;
        overflow-y: auto;
        background: #fff5f5;
        border: 1px solid #feb2b2;
        padding: 10px;
        border-radius: 5px;
        font-size: 0.9rem;
        color: #c53030;
    }
    .success-log {
        background: #f0fff4;
        border: 1px solid #9ae6b4;
        padding: 10px;
        border-radius: 5px;
        color: #2f855a;
    }
    
    /* Nút bấm full width */
    .stButton > button { width: 100%; border-radius: 6px; }
</style>
""", unsafe_allow_html=True)

# ==================== CORE LOGIC ====================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

def get_pure_text(block):
    texts = []
    t_nodes = block.getElementsByTagNameNS(W_NS, "t")
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue:
            texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def is_answer_marked(node):
    rPr_list = node.getElementsByTagNameNS(W_NS, "rPr")
    if not rPr_list: return False
    rPr = rPr_list[0]
    
    color_list = rPr.getElementsByTagNameNS(W_NS, "color")
    if color_list:
        val = color_list[0].getAttributeNS(W_NS, "val")
        if val in ["red", "FF0000", "C00000"]: return True
            
    u_list = rPr.getElementsByTagNameNS(W_NS, "u")
    if u_list:
        val = u_list[0].getAttributeNS(W_NS, "val")
        if val and val != "none": return True
    return False

def check_mcq_options(q_blocks):
    """Kiểm tra câu trắc nghiệm: Đủ A,B,C,D không? Có đáp án đúng không?"""
    text_content = " ".join([get_pure_text(b) for b in q_blocks])
    options_found = re.findall(r'\b([A-D])[\.\)]', text_content)
    unique_opts = set(opt.upper() for opt in options_found)
    
    # Check 1: Đủ 4 đáp án (cơ bản)
    missing = []
    for char in ['A', 'B', 'C', 'D']:
        if char not in unique_opts:
            missing.append(char)
            
    # Check 2: Có đáp án đúng (Màu đỏ/Gạch chân)
    has_correct = False
    for block in q_blocks:
        runs = block.getElementsByTagNameNS(W_NS, "r")
        for r in runs:
            if is_answer_marked(r):
                # Kiểm tra xem run này có chứa text A., B. hay nội dung đáp án không
                # Lấy text của run
                t_nodes = r.getElementsByTagNameNS(W_NS, "t")
                t_val = "".join([t.firstChild.nodeValue for t in t_nodes if t.firstChild])
                if t_val.strip():
                    has_correct = True
                    break
        if has_correct: break
        
    return missing, has_correct

def fix_floating_images_in_xml(doc_xml_str):
    """
    Chuyển đổi <wp:anchor> (Floating) thành <wp:inline> (Inline)
    bằng cách thay thế tag và loại bỏ các thẻ định vị tuyệt đối.
    """
    # Parse lại DOM để xử lý an toàn
    dom = minidom.parseString(doc_xml_str)
    
    # Tìm tất cả thẻ wp:anchor
    anchors = dom.getElementsByTagName("wp:anchor")
    count = 0
    
    # Duyệt ngược để không ảnh hưởng index khi thay thế
    for anchor in reversed(anchors):
        # Tạo thẻ wp:inline mới
        inline = dom.createElement("wp:inline")
        
        # Copy các attributes (trừ những cái anchor-specific nếu cần, nhưng thường giữ lại dist cũng được)
        # wp:inline không chịu thuộc tính 'behindDoc', 'locked', 'layoutInCell', 'allowOverlap', 'simplePos'
        # Chỉ giữ lại distT, distB, distL, distR nếu muốn
        
        # Move children
        # wp:inline CHỈ CHỨA: extent, effectExtent, docPr, cNvGraphicFramePr, a:graphic
        # Cần loại bỏ: simplePos, positionH, positionV
        
        valid_children = ["wp:extent", "wp:effectExtent", "wp:docPr", "wp:cNvGraphicFramePr", "a:graphic"]
        
        for child in list(anchor.childNodes):
            if child.nodeName in valid_children:
                inline.appendChild(child.cloneNode(True))
            # Nếu là a:graphic (phần hình ảnh chính), chắc chắn phải giữ
            elif child.localName == "graphic":
                inline.appendChild(child.cloneNode(True))
                
        # Thay thế anchor bằng inline
        anchor.parentNode.replaceChild(inline, anchor)
        count += 1
        
    return dom.toxml(), count

def validate_document(blocks):
    """Hàm kiểm tra lỗi logic của đề"""
    errors = []
    warnings = []
    
    # Tách câu hỏi để check
    questions = []
    current_q = []
    q_num_map = {} # Map index câu hỏi trong list -> Số câu thực tế (Câu 1, Câu 2)
    
    real_q_num = 0
    for block in blocks:
        text = get_pure_text(block)
        m = re.match(r'^Câu\s*(\d+)', text, re.IGNORECASE)
        if m:
            if current_q: questions.append(current_q)
            current_q = [block]
            real_q_num = m.group(1)
            q_num_map[len(questions)] = real_q_num
        else:
            if current_q: current_q.append(block)
    if current_q: questions.append(current_q)
    
    # Bắt đầu check từng câu
    for idx, q_blocks in enumerate(questions):
        q_label = f"Câu {q_num_map.get(idx, 'Unknown')}"
        
        # 1. Check hình ảnh floating (Cảnh báo)
        for b in q_blocks:
            if b.getElementsByTagName("wp:anchor"):
                warnings.append(f"{q_label}: Chứa hình ảnh trôi nổi (Floating). Nên chuyển về Inline.")

        # 2. Logic kiểm tra đáp án
        # Cần biết đang ở phần nào (P1 hay P3). Tạm thời detect dựa trên format
        q_text = " ".join([get_pure_text(b) for b in q_blocks])
        
        # Nếu có A. B. C. D. -> Coi là trắc nghiệm
        if re.search(r'\bA[\.\)]', q_text) and re.search(r'\bD[\.\)]', q_text):
            missing, has_correct = check_mcq_options(q_blocks)
            if missing:
                errors.append(f"❌ {q_label}: Thiếu phương án {', '.join(missing)}")
            if not has_correct:
                errors.append(f"❌ {q_label}: Chưa chọn đáp án đúng (Tô đỏ hoặc gạch chân)")
        
        # Nếu có ĐS: -> Coi là tự luận P3
        elif "ĐS" in q_text or "đs" in q_text:
            # Check xem có tô đỏ ĐS ko
            has_red_ds = False
            for b in q_blocks:
                runs = b.getElementsByTagNameNS(W_NS, "r")
                for r in runs:
                    if is_answer_marked(r):
                        t_nodes = r.getElementsByTagNameNS(W_NS, "t")
                        t_val = "".join([t.firstChild.nodeValue for t in t_nodes if t.firstChild])
                        if "ĐS" in t_val or "đs" in t_val or ":" in t_val: # Loose check
                            has_red_ds = True
            if not has_red_ds:
                 errors.append(f"❌ {q_label}: Đáp án 'ĐS:...' chưa được tô đỏ.")

    return errors, warnings

# ... [Giữ nguyên các hàm xử lý XML cũ: find_part_index, process_mcq_question, etc.] ...
# Để tiết kiệm không gian, tôi sẽ gọi lại logic cũ nhưng tích hợp vào flow mới.
# (Code xử lý trộn đề y hệt phiên bản trước, tôi sẽ paste phần update quan trọng bên dưới)

# --- RE-USE OLD FUNCTIONS (Simulated for brevity, ensure you keep full logic) ---
def update_question_label(paragraph, new_number):
    # (Giữ nguyên logic cũ)
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    found = False
    for t in t_nodes:
        if not t.firstChild: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)(Câu\s*)(\d+)([\.:])?', txt, re.IGNORECASE)
        if m:
            prefix_space = m.group(1) or ""
            punct = m.group(4) or "."
            remain = txt[m.end():]
            t.firstChild.nodeValue = f"{prefix_space}Câu {new_number}{punct}{remain}"
            found = True
            break

def get_text_with_formatting(block):
    # (Giữ nguyên logic cũ)
    texts = []
    is_correct = False
    runs = block.getElementsByTagNameNS(W_NS, "r")
    for r in runs:
        t_nodes = r.getElementsByTagNameNS(W_NS, "t")
        for t in t_nodes:
            if t.firstChild and t.firstChild.nodeValue:
                text_val = t.firstChild.nodeValue
                texts.append(text_val)
                if is_answer_marked(r) and text_val.strip():
                    is_correct = True
    return "".join(texts).strip(), is_correct

def update_mcq_label(paragraph, new_label):
    # (Giữ nguyên logic cũ)
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    new_letter = new_label[0].upper()
    full_text = get_pure_text(paragraph)
    if not re.match(r'^\s*[A-D][\.\)]', full_text, re.IGNORECASE): return
    for t in t_nodes:
        if not t.firstChild: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)([A-D])([\.\)])?', txt, re.IGNORECASE)
        if m:
            prefix = m.group(1) or ""
            punct = m.group(3) or "."
            remain = txt[m.end():]
            t.firstChild.nodeValue = f"{prefix}{new_letter}{punct}{remain}"
            break

def extract_part3_answer(block):
    # (Giữ nguyên logic cũ)
    runs = block.getElementsByTagNameNS(W_NS, "r")
    full_text = ""
    for r in runs:
        t_nodes = r.getElementsByTagNameNS(W_NS, "t")
        for t in t_nodes:
            if t.firstChild and t.firstChild.nodeValue:
                full_text += t.firstChild.nodeValue
    match = re.search(r'ĐS\s*[:\.]\s*(.+)', full_text, re.IGNORECASE)
    has_red = False
    for r in runs:
        if is_answer_marked(r): has_red = True; break
    if match and has_red: return match.group(1).strip()
    return None

def process_mcq_question(q_blocks):
    # (Giữ nguyên logic cũ)
    header = []
    options = []
    for i, block in enumerate(q_blocks):
        text = get_pure_text(block)
        if re.match(r'^\s*[A-D][\.\)]', text, re.IGNORECASE):
            options.append(block)
        else:
            header.append(block)
    if len(options) < 2: return q_blocks, ""
    original_correct_idx = -1
    for idx, opt in enumerate(options):
        _, is_correct = get_text_with_formatting(opt)
        if is_correct: original_correct_idx = idx; break
    perm = list(range(len(options)))
    random.shuffle(perm)
    shuffled_options = [options[i] for i in perm]
    new_correct_char = ""
    if original_correct_idx != -1:
        new_pos = perm.index(original_correct_idx)
        letters = ["A", "B", "C", "D", "E", "F"]
        if new_pos < len(letters): new_correct_char = letters[new_pos]
    final_letters = ["A", "B", "C", "D", "E", "F"]
    for idx, opt in enumerate(shuffled_options):
        letter = final_letters[idx] if idx < len(final_letters) else "Z"
        update_mcq_label(opt, f"{letter}.")
    return header + shuffled_options, new_correct_char

def process_part3_question(q_blocks):
    # (Giữ nguyên logic cũ)
    answer_val = ""
    for block in q_blocks:
        val = extract_part3_answer(block)
        if val: answer_val = val; break
    return q_blocks, answer_val

def parse_questions(blocks):
    # (Giữ nguyên logic cũ)
    questions = []
    current_q = []
    for block in blocks:
        text = get_pure_text(block)
        if re.match(r'^Câu\s*\d+\b', text, re.IGNORECASE):
            if current_q: questions.append(current_q)
            current_q = [block]
        else:
            if current_q: current_q.append(block)
    if current_q: questions.append(current_q)
    return questions

def find_part_index(blocks, part_number):
    pattern = re.compile(rf'PHẦN\s*{part_number}\b', re.IGNORECASE)
    for i, block in enumerate(blocks):
        text = get_pure_text(block)
        if pattern.search(text): return i
    return -1

# ==================== MAIN PROCESSING FUNCTION ====================

def process_document_final(file_bytes, num_versions, filename_prefix, auto_fix_img):
    input_buffer = io.BytesIO(file_bytes)
    zip_in = zipfile.ZipFile(input_buffer, 'r')
    doc_xml = zip_in.read("word/document.xml").decode('utf-8')
    
    # BƯỚC 1: AUTO FIX IMAGE (Nếu được chọn)
    if auto_fix_img:
        doc_xml, fixed_count = fix_floating_images_in_xml(doc_xml)
    
    dom = minidom.parseString(doc_xml)
    body = dom.getElementsByTagNameNS(W_NS, "body")[0]
    all_blocks = [node for node in body.childNodes if node.nodeType == node.ELEMENT_NODE and node.localName in ["p", "tbl"]]
    
    # BƯỚC 2: VALIDATE (Chỉ chạy 1 lần để báo lỗi)
    # Tuy nhiên hàm process này chạy khi bấm nút "Trộn".
    # Việc validate UI tách riêng ở hàm main.
    
    # ... (Logic chia phần giống code cũ) ...
    p1_idx = find_part_index(all_blocks, 1)
    p2_idx = find_part_index(all_blocks, 2)
    p3_idx = find_part_index(all_blocks, 3)
    
    parts_data = {"intro": [], "p1": [], "p2": [], "p3": []}
    cursor = 0
    if p1_idx != -1:
        parts_data["intro"] = all_blocks[cursor:p1_idx+1]
        cursor = p1_idx + 1
        end1 = p2_idx if p2_idx != -1 else (p3_idx if p3_idx != -1 else len(all_blocks))
        parts_data["p1"] = all_blocks[cursor:end1]
        cursor = end1
    else:
        parts_data["p1"] = all_blocks
        cursor = len(all_blocks)
    if p2_idx != -1:
        end2 = p3_idx if p3_idx != -1 else len(all_blocks)
        parts_data["p2"] = all_blocks[cursor:end2]
        cursor = end2
    if p3_idx != -1:
        parts_data["p3"] = all_blocks[cursor:]

    all_keys = []
    zip_out_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_out_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_final:
        for i in range(num_versions):
            v_name = f"{101 + i}"
            final_blocks = []
            final_blocks.extend(parts_data["intro"])
            answer_key = {"Mã đề": v_name}
            global_q_idx = 1
            
            # COPY LIST để trộn (Quan trọng: phải clone node nếu cần, ở đây ta shuffle list reference)
            # DOM Node manipulation is destructive (move node), so we must clone for versions > 1?
            # Minidom move node remove it from old parent.
            # ==> FIX: Với minidom, khi appendChild vào body mới, nó remove khỏi chỗ cũ.
            # Do đó với multi-version, ta phải PARSE lại XML gốc cho mỗi vòng lặp hoặc Clone deep.
            # Để đơn giản và an toàn nhất: Ta parse lại string XML cho mỗi version.
            
            # --- RE-PARSE STRATEGY ---
            dom_v = minidom.parseString(doc_xml)
            body_v = dom_v.getElementsByTagNameNS(W_NS, "body")[0]
            blocks_v = [n for n in body_v.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName in ["p", "tbl"]]
            
            # Recalculate indices for this version instance
            p1_i = find_part_index(blocks_v, 1)
            p2_i = find_part_index(blocks_v, 2)
            p3_i = find_part_index(blocks_v, 3)
            
            # Slicing logic again for this instance
            parts_v = {"intro": [], "p1": [], "p2": [], "p3": []}
            cur = 0
            if p1_i != -1:
                parts_v["intro"] = blocks_v[cur:p1_i+1]
                cur = p1_i + 1
                e1 = p2_i if p2_i != -1 else (p3_i if p3_i != -1 else len(blocks_v))
                parts_v["p1"] = blocks_v[cur:e1]
                cur = e1
            else:
                parts_v["p1"] = blocks_v
                cur = len(blocks_v)
            if p2_i != -1:
                e2 = p3_i if p3_i != -1 else len(blocks_v)
                parts_v["p2"] = blocks_v[cur:e2]
                cur = e2
            if p3_i != -1:
                parts_v["p3"] = blocks_v[cur:]
                
            # Build Layout
            layout_blocks = []
            layout_blocks.extend(parts_v["intro"])
            
            # P1 Mix
            if parts_v["p1"]:
                qs = parse_questions(parts_v["p1"])
                random.shuffle(qs)
                for q in qs:
                    update_question_label(q[0], global_q_idx)
                    p_blocks, ans = process_mcq_question(q)
                    layout_blocks.extend(p_blocks)
                    if ans: answer_key[f"Câu {global_q_idx}"] = ans
                    global_q_idx += 1
            
            # P2 Mix (Shuffle Questions only)
            if parts_v["p2"]:
                h_p2 = [parts_v["p2"][0]] if parts_v["p2"] else []
                c_p2 = parts_v["p2"][1:] if len(parts_v["p2"]) > 1 else []
                layout_blocks.extend(h_p2)
                qs = parse_questions(c_p2)
                random.shuffle(qs)
                for q in qs:
                    update_question_label(q[0], global_q_idx)
                    layout_blocks.extend(q)
                    global_q_idx += 1
                    
            # P3 Mix
            if parts_v["p3"]:
                h_p3 = [parts_v["p3"][0]] if parts_v["p3"] else []
                c_p3 = parts_v["p3"][1:] if len(parts_v["p3"]) > 1 else []
                layout_blocks.extend(h_p3)
                qs = parse_questions(c_p3)
                random.shuffle(qs)
                for q in qs:
                    update_question_label(q[0], global_q_idx)
                    p_blocks, ans = process_part3_question(q)
                    layout_blocks.extend(p_blocks)
                    if ans: answer_key[f"Câu {global_q_idx}"] = ans
                    global_q_idx += 1
            
            # Rebuild XML body
            while body_v.hasChildNodes():
                body_v.removeChild(body_v.firstChild)
            for b in layout_blocks:
                body_v.appendChild(b)
                
            new_xml = dom_v.toxml()
            fname = f"{filename_prefix}_{v_name}.docx"
            
            # Write to Zip
            # Need to construct a DOCX bytes for this version
            # Use original zip_in to copy other files, but replace document.xml
            ver_io = io.BytesIO()
            with zipfile.ZipFile(ver_io, 'w', zipfile.ZIP_DEFLATED) as z_ver:
                for item in zip_in.infolist():
                    if item.filename == "word/document.xml":
                        z_ver.writestr(item, new_xml.encode('utf-8'))
                    else:
                        z_ver.writestr(item, zip_in.read(item.filename))
            
            zip_final.writestr(fname, ver_io.getvalue())
            all_keys.append(answer_key)
            
    # Excel
    df = pd.DataFrame(all_keys)
    cols = list(df.columns)
    if "Mã đề" in cols: cols.remove("Mã đề")
    def sort_key(s):
        m = re.search(r'(\d+)', s)
        return int(m.group(1)) if m else 0
    q_cols = sorted(cols, key=sort_key)
    final_cols = ["Mã đề"] + q_cols
    df = df.reindex(columns=final_cols)
    
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='DapAn')
        
    return zip_out_buffer.getvalue(), excel_buffer.getvalue()

# ==================== GUI LOGIC ====================

def main():
    st.markdown("""
    <div class="main-header">
        <h1>⚡ AIOMT Premium</h1>
        <p>Hệ thống trộn đề thi trắc nghiệm & Tự luận thông minh</p>
    </div>
    """, unsafe_allow_html=True)

    # --- SIDEBAR (CONFIGURATION) ---
    with st.sidebar:
        st.header("⚙️ Cấu hình")
        uploaded_file = st.file_uploader("Chọn file Word (.docx)", type=["docx"])
        
        st.markdown("---")
        num_versions = st.number_input("Số lượng mã đề", 1, 50, 4)
        filename_prefix = st.text_input("Mã đề gốc", "KiemTra")
        
        st.markdown("---")
        st.markdown("**Tính năng nâng cao:**")
        auto_fix_img = st.checkbox("Tự động chuyển hình về Inline", value=False, help="Cố gắng chuyển hình ảnh dạng Floating sang Inline with Text để không bị trôi hình.")
        
        st.info("Phiên bản: 2.1 (Laptop Optimized)")

    # --- MAIN SCREEN (DASHBOARD) ---
    if not uploaded_file:
        st.warning("👈 Vui lòng tải file đề gốc lên từ thanh bên trái.")
        st.markdown("""
        ### Quy định soạn thảo:
        * **Câu hỏi:** Bắt đầu bằng `Câu 1`, `Câu 2`...
        * **Đáp án đúng (MCQ):** Tô **màu đỏ** hoặc **gạch chân**.
        * **Đáp án P3:** Dạng `ĐS: kết quả` (Tô đỏ).
        """)
        return

    # LOAD FILE & VALIDATE
    file_bytes = uploaded_file.read()
    
    # Pre-process for validation (Need XML)
    try:
        input_buffer = io.BytesIO(file_bytes)
        zip_in = zipfile.ZipFile(input_buffer, 'r')
        doc_xml = zip_in.read("word/document.xml").decode('utf-8')
        dom = minidom.parseString(doc_xml)
        body = dom.getElementsByTagNameNS(W_NS, "body")[0]
        blocks = [n for n in body.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName in ["p", "tbl"]]
        
        errors, warnings = validate_document(blocks)
        
    except Exception as e:
        st.error(f"Lỗi khi đọc file: {e}")
        return

    # LAYOUT 2 CỘT: [VALIDATION REPORT] | [ACTION & RESULT]
    col1, col2 = st.columns([1, 1], gap="medium")

    with col1:
        st.markdown("### 🔍 Kiểm tra cấu trúc")
        if not errors and not warnings:
            st.markdown('<div class="success-log">✅ File chuẩn! Sẵn sàng trộn.</div>', unsafe_allow_html=True)
            is_valid = True
        else:
            if errors:
                st.markdown(f"**Phát hiện {len(errors)} lỗi nghiêm trọng:**")
                error_html = "".join([f"<div>• {e}</div>" for e in errors])
                st.markdown(f'<div class="error-log">{error_html}</div>', unsafe_allow_html=True)
                is_valid = False # Có lỗi nghiêm trọng -> Không cho trộn
            else:
                is_valid = True # Chỉ có warning -> Cho phép trộn

            if warnings:
                st.markdown(f"**Cảnh báo ({len(warnings)}):**")
                warn_html = "".join([f"<div>⚠️ {w}</div>" for w in warnings])
                st.markdown(f'<div class="error-log" style="background:#fffaf0; color:#b7791f; border-color:#f6e05e">{warn_html}</div>', unsafe_allow_html=True)

    with col2:
        st.markdown("### 🚀 Tác vụ")
        
        if is_valid:
            if st.button(f"Trộn ngay {num_versions} đề", type="primary"):
                with st.spinner("Đang xử lý..."):
                    try:
                        # Reset file pointer
                        uploaded_file.seek(0)
                        
                        zip_data, excel_data = process_document_final(
                            uploaded_file.read(), 
                            num_versions, 
                            filename_prefix,
                            auto_fix_img
                        )
                        
                        st.success("Hoàn tất!")
                        
                        # Khu vực tải xuống gọn gàng
                        d_col1, d_col2 = st.columns(2)
                        with d_col1:
                            st.download_button("📥 File Đề (ZIP)", zip_data, f"{filename_prefix}_Mix.zip", "application/zip")
                        with d_col2:
                            st.download_button("📊 Đáp án (Excel)", excel_data, f"{filename_prefix}_Key.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                            
                    except Exception as e:
                        st.error(f"Lỗi xử lý: {e}")
        else:
            st.error("⛔ Vui lòng sửa các lỗi cú pháp trong file Word trước khi trộn.")
            st.markdown("Download file mẫu chuẩn tại Zalo nhóm.")

    # Footer gọn
    st.markdown("---")
    st.markdown('<div style="text-align:center; color: grey; font-size: 0.8rem;">© 2025 Phan Trường Duy - THPT Minh Đức</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
