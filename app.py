"""
Trộn Đề Word Online - AIOMT Premium (Merged Final Version)
Author: Phan Trường Duy - THPT Minh Đức
"""

import streamlit as st
import re
import random
import zipfile
import io
import pandas as pd
from xml.dom import minidom
import sys

# ==================== CẤU HÌNH TRANG ====================

st.set_page_config(
    page_title="App Trộn Đề 2025",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==================== CSS CUSTOM DESIGN (THEO YÊU CẦU) ====================
st.markdown("""
<style>
    /* 1. HEADER (TO 200%, KHÍT DÒNG, CAO 75%) */
    .main-header {
        text-align: center;
        padding: 1.5rem 0; /* Giảm padding */
        background: linear-gradient(to right, #009688, #00796b);
        color: white;
        border-radius: 0 0 15px 15px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 8px rgba(0,0,0,0.15);
    }
    .main-header h1 {
        font-family: 'Arial', sans-serif;
        font-weight: 800; /* Nét đều, đậm */
        font-size: 3rem !important; /* To 200% */
        text-transform: uppercase;
        margin: 0;
        line-height: 1.1;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    .main-header p {
        font-family: 'Arial', sans-serif;
        font-size: 1.1rem;
        margin-top: 5px; /* Khít lại */
        margin-bottom: 0;
        opacity: 0.9;
        font-weight: 500;
        letter-spacing: 2px;
    }

    /* 2. STYLE CHO THẺ (CARD) */
    .step-label {
        font-size: 1.1rem;
        font-weight: 700;
        color: #004d40;
        margin-bottom: 10px;
        display: flex;
        align-items: center;
    }
    .step-badge {
        background-color: #009688;
        color: white;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        margin-right: 10px;
        font-size: 0.9rem;
        font-weight: bold;
    }

    /* 3. KHUNG HƯỚNG DẪN (HTML CHUẨN) */
    .instruction-card {
        background-color: #e0f2f1;
        border-radius: 10px;
        padding: 15px;
        color: #004d40;
        font-size: 0.9rem;
        border: 1px solid #b2dfdb;
    }
    .part-title {
        font-weight: bold;
        color: #00796b;
        display: inline-block;
        width: 70px;
    }
    .warning-box {
        background-color: #fff8e1;
        border: 1px solid #ffe082;
        border-radius: 8px;
        padding: 10px;
        margin-top: 10px;
        color: #5d4037;
    }
    .code-tag {
        background-color: #fff;
        padding: 2px 5px;
        border-radius: 4px;
        border: 1px solid #e0e0e0;
        font-family: monospace;
        color: #d84315;
        font-weight: bold;
    }

    /* 4. CUSTOM RADIO BUTTONS (DẠNG THẺ DỌC - NHƯ HÌNH) */
    div[role="radiogroup"] {
        display: flex;
        flex-direction: column; 
        gap: 10px;
    }
    div[role="radiogroup"] > label {
        width: 100%;
        background-color: white;
        border: 1px solid #cfd8dc;
        border-radius: 8px;
        padding: 15px;
        display: flex;
        align-items: center;
        transition: all 0.2s;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 0px !important;
    }
    div[role="radiogroup"] > label:hover {
        border-color: #009688;
        background-color: #f0fdfa;
        transform: translateX(5px);
    }
    
    /* 5. UPLOAD BOX */
    .stFileUploader {
        border: 2px dashed #009688;
        border-radius: 10px;
        padding: 15px;
        background-color: white;
        text-align: center;
    }

    /* 6. BUTTON */
    .stButton > button {
        background: #009688;
        color: white;
        border-radius: 8px;
        padding: 12px;
        font-weight: bold;
        border: none;
        width: 100%;
        font-size: 1.1rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        margin-top: 10px;
    }
    .stButton > button:hover {
        background: #00796b;
        transform: translateY(-2px);
    }

    .block-container { padding-top: 1rem !important; }
</style>
""", unsafe_allow_html=True)

# ==================== CORE LOGIC (BỘ NÃO XỬ LÝ) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_pure_text(block):
    """Lấy text thuần từ block"""
    texts = []
    t_nodes = block.getElementsByTagNameNS(W_NS, "t")
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue:
            texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def is_answer_marked(node):
    """Kiểm tra xem run có được tô đỏ hoặc gạch chân không"""
    rPr_list = node.getElementsByTagNameNS(W_NS, "rPr")
    if not rPr_list: return False
    rPr = rPr_list[0]
    
    # Check màu đỏ
    color_list = rPr.getElementsByTagNameNS(W_NS, "color")
    if color_list:
        val = color_list[0].getAttributeNS(W_NS, "val")
        if val in ["red", "FF0000", "C00000"]: return True
            
    # Check gạch chân
    u_list = rPr.getElementsByTagNameNS(W_NS, "u")
    if u_list:
        val = u_list[0].getAttributeNS(W_NS, "val")
        if val and val != "none": return True
    return False

def get_text_with_formatting(block):
    """Lấy text và kiểm tra có phải đáp án đúng không"""
    texts = []
    is_correct = False
    runs = block.getElementsByTagNameNS(W_NS, "r")
    for r in runs:
        t_nodes = r.getElementsByTagNameNS(W_NS, "t")
        for t in t_nodes:
            if t.firstChild and t.firstChild.nodeValue:
                texts.append(t.firstChild.nodeValue)
                if is_answer_marked(r) and t.firstChild.nodeValue.strip():
                    is_correct = True
    return "".join(texts).strip(), is_correct

def style_run_blue_bold(run):
    """Format xanh đậm cho nhãn A. B. C."""
    doc = run.ownerDocument
    rPr_list = run.getElementsByTagNameNS(W_NS, "rPr")
    if rPr_list: rPr = rPr_list[0]
    else:
        rPr = doc.createElementNS(W_NS, "w:rPr")
        run.insertBefore(rPr, run.firstChild)
    
    color_list = rPr.getElementsByTagNameNS(W_NS, "color")
    if color_list: color_el = color_list[0]
    else:
        color_el = doc.createElementNS(W_NS, "w:color")
        rPr.appendChild(color_el)
    color_el.setAttributeNS(W_NS, "w:val", "0000FF") # Màu xanh
    
    b_list = rPr.getElementsByTagNameNS(W_NS, "b")
    if not b_list:
        b_el = doc.createElementNS(W_NS, "w:b")
        rPr.appendChild(b_el)

def update_mcq_label(paragraph, new_label):
    """Cập nhật nhãn A. B. C. D."""
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    new_letter = new_label[0].upper()
    for i, t in enumerate(t_nodes):
        if not t.firstChild: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)([A-D])([\.\)])?', txt, re.IGNORECASE)
        if m:
            t.firstChild.nodeValue = m.group(1) + new_letter + "." + txt[m.end():]
            run = t.parentNode
            if run and run.localName == "r": style_run_blue_bold(run)
            break

def update_tf_label(paragraph, new_label):
    """Cập nhật nhãn a) b) c) d)"""
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    new_letter = new_label[0].lower()
    for i, t in enumerate(t_nodes):
        if not t.firstChild: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)([a-d])(\))?', txt, re.IGNORECASE)
        if m:
            t.firstChild.nodeValue = m.group(1) + new_letter + ")" + txt[m.end():]
            run = t.parentNode
            if run and run.localName == "r": style_run_blue_bold(run)
            break

def update_question_label(paragraph, new_label):
    """Cập nhật Câu 1."""
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    for i, t in enumerate(t_nodes):
        if not t.firstChild: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)(Câu\s*)(\d+)([\.:])?', txt, re.IGNORECASE)
        if m:
            t.firstChild.nodeValue = m.group(1) + new_label + txt[m.end():]
            run = t.parentNode
            if run and run.localName == "r": style_run_blue_bold(run)
            break

def find_part_index(blocks, part_number):
    """Tìm dòng PHẦN 1, 2, 3"""
    pattern = re.compile(rf'PHẦN\s*{part_number}\b', re.IGNORECASE)
    for i, block in enumerate(blocks):
        if pattern.search(get_pure_text(block)): return i
    return -1

def parse_questions_in_range(blocks, start, end):
    """Tách câu hỏi trong range"""
    part_blocks = blocks[start:end]
    intro = []
    questions = []
    i = 0
    # Tách intro
    while i < len(part_blocks):
        if re.match(r'^Câu\s*\d+\b', get_pure_text(part_blocks[i]), re.IGNORECASE): break
        intro.append(part_blocks[i])
        i += 1
    # Tách câu
    while i < len(part_blocks):
        text = get_pure_text(part_blocks[i])
        if re.match(r'^Câu\s*\d+\b', text, re.IGNORECASE):
            group = [part_blocks[i]]
            i += 1
            while i < len(part_blocks):
                t2 = get_pure_text(part_blocks[i])
                # Dừng nếu gặp Câu mới hoặc PHẦN mới
                if re.match(r'^Câu\s*\d+\b', t2, re.IGNORECASE) or re.match(r'^PHẦN\s*\d\b', t2, re.IGNORECASE): break
                group.append(part_blocks[i])
                i += 1
            questions.append(group)
        else:
            # Dòng rác hoặc trôi nổi
            intro.append(part_blocks[i])
            i += 1
    return intro, questions

def shuffle_array(arr):
    out = arr.copy()
    random.shuffle(out)
    return out

def process_mcq_question_with_key(question_blocks):
    """Trộn MCQ và lấy đáp án"""
    indices = []
    for i, block in enumerate(question_blocks):
        if re.match(r'^\s*[A-D][\.\)]', get_pure_text(block), re.IGNORECASE): indices.append(i)
    
    if len(indices) < 2: return question_blocks, ""

    options = [question_blocks[idx] for idx in indices]
    
    # Tìm đáp án gốc
    original_correct_idx = -1
    for idx, opt in enumerate(options):
        _, is_correct = get_text_with_formatting(opt)
        if is_correct:
            original_correct_idx = idx
            break
            
    # Trộn
    perm = list(range(len(options)))
    random.shuffle(perm)
    shuffled_options = [options[i] for i in perm]
    
    # Xác định key mới
    new_correct_char = ""
    letters = ["A", "B", "C", "D", "E", "F"]
    if original_correct_idx != -1 and original_correct_idx in perm:
        new_pos = perm.index(original_correct_idx)
        new_correct_char = letters[new_pos] if new_pos < len(letters) else ""

    # Đánh lại nhãn
    for idx, block in enumerate(shuffled_options):
        letter = letters[idx] if idx < len(letters) else "Z"
        update_mcq_label(block, f"{letter}.")

    # Ghép lại
    min_idx = min(indices)
    max_idx = max(indices)
    before = question_blocks[:min_idx]
    after = question_blocks[max_idx+1:]
    
    return before + shuffled_options + after, new_correct_char

def process_tf_question(question_blocks):
    """Trộn Đúng/Sai"""
    option_indices = {}
    for i, block in enumerate(question_blocks):
        m = re.match(r'^\s*([a-d])\)', get_pure_text(block), re.IGNORECASE)
        if m: option_indices[m.group(1).lower()] = i
    
    abc_idx = [option_indices.get(k) for k in ["a","b","c"] if option_indices.get(k) is not None]
    if len(abc_idx) < 2: return question_blocks
    
    abc_nodes = [question_blocks[idx] for idx in abc_idx]
    shuffled_abc = shuffle_array(abc_nodes)
    
    all_vals = [v for v in option_indices.values() if v is not None]
    min_i, max_i = min(all_vals), max(all_vals)
    
    before = question_blocks[:min_i]
    after = question_blocks[max_i+1:]
    
    middle = shuffled_abc.copy()
    if "d" in option_indices: middle.append(question_blocks[option_indices["d"]])
    
    for idx, block in enumerate(middle):
        if idx < 3: update_tf_label(block, f"{['a','b','c'][idx]})")
        
    return before + middle + after

def extract_part3_answer(blocks):
    """Lấy đáp án P3"""
    full_text = ""
    has_red = False
    for b in blocks:
        full_text += get_pure_text(b)
        runs = b.getElementsByTagNameNS(W_NS, "r")
        for r in runs:
            if is_answer_marked(r): has_red = True
    match = re.search(r'ĐS\s*[:\.]\s*(.+)', full_text, re.IGNORECASE)
    if match and has_red: return match.group(1).strip()
    return None

def fix_floating_images_in_xml(doc_xml_str):
    """Sửa lỗi hình ảnh"""
    dom = minidom.parseString(doc_xml_str)
    anchors = dom.getElementsByTagName("wp:anchor")
    for anchor in reversed(anchors):
        inline = dom.createElement("wp:inline")
        valid_children = ["wp:extent", "wp:effectExtent", "wp:docPr", "wp:cNvGraphicFramePr", "a:graphic"]
        for child in list(anchor.childNodes):
            if child.nodeName in valid_children: inline.appendChild(child.cloneNode(True))
            elif child.localName == "graphic": inline.appendChild(child.cloneNode(True))
        anchor.parentNode.replaceChild(inline, anchor)
    return dom.toxml()

def validate_document(blocks):
    """Kiểm tra lỗi file"""
    errors = []
    warnings = []
    
    full_text = " ".join([get_pure_text(b) for b in blocks])
    
    if not re.search(r'Câu\s*1', full_text, re.IGNORECASE):
        errors.append("❌ Không tìm thấy 'Câu 1'. File phải bắt đầu câu hỏi bằng 'Câu 1.'")
        
    if not re.search(r'A\.', full_text) and not re.search(r'A\)', full_text):
        warnings.append("⚠️ Cảnh báo: Không tìm thấy đáp án A. B. C. D.")
        
    for b in blocks:
        if b.getElementsByTagName("wp:anchor"):
            warnings.append("⚠️ Có hình ảnh trôi nổi (Floating). Hệ thống sẽ tự sửa.")
            break
            
    return errors, warnings

def process_document_final(file_bytes, num_versions, filename_prefix, auto_fix_img, shuffle_mode="auto"):
    # TẠO BYTES MỚI CHO MỖI LẦN GỌI ĐỂ TRÁNH BAD MAGIC NUMBER
    input_buffer = io.BytesIO(file_bytes)
    if not zipfile.is_zipfile(input_buffer): raise Exception("File không hợp lệ.")
    
    zip_in = zipfile.ZipFile(input_buffer, 'r')
    doc_xml = zip_in.read("word/document.xml").decode('utf-8')
    if auto_fix_img: doc_xml = fix_floating_images_in_xml(doc_xml)
    
    dom = minidom.parseString(doc_xml)
    all_keys = []
    zip_out_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_out_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_final:
        for i in range(num_versions):
            v_name = f"{101 + i}"
            dom_v = minidom.parseString(doc_xml)
            body_v = dom_v.getElementsByTagNameNS(W_NS, "body")[0]
            
            # Giữ sectPr
            sectPr = None
            if body_v.lastChild and body_v.lastChild.localName == 'sectPr':
                sectPr = body_v.lastChild.cloneNode(True)
                
            blocks_v = [n for n in body_v.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName in ["p", "tbl"]]
            
            # Chia phần
            p1_i = find_part_index(blocks_v, 1)
            p2_i = find_part_index(blocks_v, 2)
            p3_i = find_part_index(blocks_v, 3)
            
            parts = {"intro": [], "p1": [], "p2": [], "p3": []}
            cur = 0
            
            if shuffle_mode == "auto":
                if p1_i >= 0:
                    parts["intro"] = blocks_v[cur:p1_i+1]
                    cur = p1_i+1
                    end1 = p2_i if p2_i >= 0 else (p3_i if p3_i >= 0 else len(blocks_v))
                    parts["p1"] = blocks_v[cur:end1]
                    cur = end1
                else:
                    parts["p1"] = blocks_v
                    cur = len(blocks_v)
                if p2_i >= 0:
                    end2 = p3_i if p3_i >= 0 else len(blocks_v)
                    parts["p2"] = blocks_v[cur:end2]
                    cur = end2
                if p3_i >= 0:
                    parts["p3"] = blocks_v[cur:]
            elif shuffle_mode == "mcq":
                parts["p1"] = blocks_v
            elif shuffle_mode == "tf":
                parts["p2"] = blocks_v

            final_layout = []
            final_layout.extend(parts["intro"])
            ans_key = {"Mã đề": v_name}
            g_idx = 1
            
            # Xử lý P1
            if parts["p1"]:
                intro1, qs1 = parse_questions_in_range(parts["p1"], 0, len(parts["p1"]))
                final_layout.extend(intro1)
                random.shuffle(qs1)
                for q in qs1:
                    update_question_label(q[0], f"Câu {g_idx}.")
                    new_q, ans = process_mcq_question_with_key(q)
                    final_layout.extend(new_q)
                    if ans: ans_key[f"Câu {g_idx}"] = ans
                    g_idx += 1
            
            # Xử lý P2
            if parts["p2"]:
                intro2, qs2 = parse_questions_in_range(parts["p2"], 0, len(parts["p2"]))
                final_layout.extend(intro2)
                random.shuffle(qs2)
                for q in qs2:
                    update_question_label(q[0], f"Câu {g_idx}.")
                    new_q = process_tf_question(q)
                    final_layout.extend(new_q)
                    g_idx += 1
            
            # Xử lý P3
            if parts["p3"]:
                intro3, qs3 = parse_questions_in_range(parts["p3"], 0, len(parts["p3"]))
                final_layout.extend(intro3)
                random.shuffle(qs3)
                for q in qs3:
                    update_question_label(q[0], f"Câu {g_idx}.")
                    final_layout.extend(q)
                    val = extract_part3_answer(q)
                    if val: ans_key[f"Câu {g_idx}"] = val
                    g_idx += 1
            
            # Rebuild XML
            while body_v.hasChildNodes(): body_v.removeChild(body_v.firstChild)
            for b in final_layout: body_v.appendChild(b)
            if sectPr: body_v.appendChild(sectPr)
            
            # Save version
            ver_io = io.BytesIO()
            with zipfile.ZipFile(ver_io, 'w', zipfile.ZIP_DEFLATED) as z_ver:
                for item in zip_in.infolist():
                    if item.filename == "word/document.xml":
                        z_ver.writestr(item, dom_v.toxml().encode('utf-8'))
                    else:
                        z_ver.writestr(item, zip_in.read(item.filename))
            zip_final.writestr(f"{filename_prefix}_{v_name}.docx", ver_io.getvalue())
            all_keys.append(ans_key)
            
    # Excel
    df = pd.DataFrame(all_keys)
    cols = list(df.columns)
    if "Mã đề" in cols: cols.remove("Mã đề")
    def sort_key(s):
        m = re.search(r'(\d+)', s)
        return int(m.group(1)) if m else 0
    q_cols = sorted(cols, key=sort_key)
    df = df.reindex(columns=["Mã đề"] + q_cols)
    
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='DapAn')
        
    return zip_out_buffer.getvalue(), excel_buf.getvalue()

# ==================== MAIN UI ====================

def main():
    st.markdown("""
    <div class="main-header">
        <h1>TRƯỜNG THPT MINH ĐỨC</h1>
        <p>APP TRỘN ĐỀ 2025</p>
    </div>
    """, unsafe_allow_html=True)

    col_left, col_right = st.columns([1, 1], gap="medium")

    with col_left:
        # HƯỚNG DẪN HTML (KHÔNG THỤT DÒNG ĐỂ TRÁNH LỖI)
        with st.expander("📄 Hướng dẫn & Cấu trúc (Bấm để xem)", expanded=False):
            st.markdown("""
<div style="text-align: right; margin-bottom: 10px;">
<a href="https://docs.google.com/document/d/1pC6rw04BSnNQnWRAn9an-1HyWQEHDDQB/edit?usp=sharing&ouid=112824050529887271694&rtpof=true&sd=true" target="_blank" 
style="background-color:#009688; color:white; padding:5px 10px; border-radius:5px; text-decoration:none; font-weight:bold;">
📥 Tải File Mẫu
</a>
</div>
<div class="instruction-card">
<div>📌 <b>Cấu trúc file Word chuẩn:</b></div>
<div style="margin-top:5px;"><span class="part-title">PHẦN 1:</span> Trắc nghiệm (A. B. C. D.)</div>
<div><span class="part-title">PHẦN 2:</span> Đúng/Sai (a) b) c) d))</div>
<div><span class="part-title">PHẦN 3:</span> Trả lời ngắn (ĐS:...)</div>
<div class="warning-box">
<div style="font-weight:bold; color:#e65100; margin-bottom:5px;">⚠️ Lưu ý quan trọng:</div>
<ul style="margin-bottom: 0; padding-left: 20px;">
<li>Câu hỏi bắt đầu bằng <span class="code-tag">Câu 1.</span></li>
<li>Đáp án đúng: <span style="text-decoration:underline;">gạch chân</span> hoặc <span style="color:blue; font-weight:bold;">tô màu đỏ</span>.</li>
<li style="margin-top:5px; border-top:1px dashed #ccc; padding-top:5px;">
<b>Đáp án Phần 3:</b> Ghi <span style="color:red; font-weight:bold;">ĐS: Kết quả</span> và tô đỏ.
</li>
</ul>
</div>
</div>
""", unsafe_allow_html=True)
        
        st.markdown('<div class="step-label"><div class="step-badge">1</div>Chọn file đề Word (*.docx)</div>', unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Kéo thả file vào đây", type=["docx"], label_visibility="collapsed")
        
        if uploaded_file is not None:
            # FIX: LƯU FILE VÀO SESSION ĐỂ TRÁNH LỖI BAD MAGIC NUMBER KHI RERUN
            uploaded_file.seek(0)
            st.session_state['file_bytes'] = uploaded_file.read()
            st.success(f"✅ Đã tải lên: {uploaded_file.name}")
            
            if st.button("🔍 Kiểm tra cấu trúc & Lỗi"):
                try:
                    # Đọc từ Session
                    input_buffer = io.BytesIO(st.session_state['file_bytes'])
                    zip_in = zipfile.ZipFile(input_buffer, 'r')
                    doc_xml = zip_in.read("word/document.xml").decode('utf-8')
                    dom = minidom.parseString(doc_xml)
                    body = dom.getElementsByTagNameNS(W_NS, "body")[0]
                    blocks = [n for n in body.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName in ["p", "tbl"]]
                    errors, warnings = validate_document(blocks)
                    
                    if not errors and not warnings:
                        st.success("✅ File chuẩn! Sẵn sàng trộn.")
                        st.session_state['is_valid'] = True
                        st.session_state['auto_fix_img'] = False
                    else:
                        if errors:
                            st.error(f"❌ Phát hiện {len(errors)} lỗi:")
                            for e in errors: st.write(e)
                            st.session_state['is_valid'] = False
                        else:
                            st.session_state['is_valid'] = True
                        if warnings:
                            st.warning(f"⚠️ {len(warnings)} hình ảnh bị trôi.")
                            st.info("💡 Hệ thống sẽ tự động sửa khi trộn.")
                            st.session_state['auto_fix_img'] = True
                except Exception as e:
                    st.error(f"Lỗi đọc file: {str(e)}")

    with col_right:
        st.markdown('<div class="step-label"><div class="step-badge">2</div>Chọn kiểu trộn</div>', unsafe_allow_html=True)
        mode = st.radio("Mode", ["auto", "mcq", "tf"], format_func=lambda x: {
            "auto": "🔄 Tự động (Phần 1, 2, 3)", "mcq": "📝 Trắc nghiệm", "tf": "✅ Đúng/Sai"
        }[x], label_visibility="collapsed", horizontal=False)
        
        st.write("")
        st.markdown('<div class="step-label"><div class="step-badge">3</div>Số mã đề cần tạo</div>', unsafe_allow_html=True)
        c_num1, c_num2 = st.columns([1, 2])
        with c_num1:
            num_mix = st.number_input("Số lượng", 1, 50, 4, label_visibility="collapsed")
        with c_num2:
            st.markdown("""<div style="font-size:0.9rem; color:#666; padding-top:10px;">● 1 mã → File Word<br>● Nhiều mã → File ZIP</div>""", unsafe_allow_html=True)

        st.markdown("---")
        
        if st.button("🚀 Trộn đề & Tải xuống"):
            if 'file_bytes' in st.session_state:
                if st.session_state.get('is_valid', True):
                    with st.spinner("Đang xử lý..."):
                        do_fix = st.session_state.get('auto_fix_img', True)
                        try:
                            z_data, e_data = process_document_final(
                                st.session_state['file_bytes'], num_mix, "KiemTra", do_fix, mode
                            )
                            st.success("Thành công!")
                            d1, d2 = st.columns(2)
                            with d1:
                                st.download_button("📥 Tải Đề (ZIP)", z_data, "De_Tron.zip", "application/zip", use_container_width=True)
                            with d2:
                                st.download_button("📊 Đáp án (Excel)", e_data, "Dap_An.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                        except Exception as e:
                            st.error(f"Lỗi xử lý: {e}")
                else:
                    st.error("File lỗi. Vui lòng kiểm tra lại.")
            else:
                st.warning("Vui lòng tải file ở Bước 1.")

    st.markdown('<div style="text-align:center; color: #aaa; margin-top:20px; font-size:0.8rem;">© 2025 Phan Trường Duy - THPT Minh Đức</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
