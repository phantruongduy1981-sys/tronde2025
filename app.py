"""
Trộn Đề Word Online - AIOMT Premium (Compact & Auto-Fix UI)
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
    page_title="Trộn Đề 2025",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==================== CSS TỐI ƯU KHÔNG GIAN ====================
st.markdown("""
<style>
    /* 1. Header nhỏ gọn (70% so với bản cũ) */
    .main-header {
        text-align: center;
        padding: 1rem 0; /* Giảm padding */
        background: linear-gradient(to right, #009688, #004d40);
        color: white;
        border-radius: 0 0 15px 15px;
        margin-bottom: 1rem; /* Giảm khoảng cách dưới */
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .main-header h1 {
        font-family: 'Arial', sans-serif;
        font-weight: 800;
        font-size: 2.2rem !important; /* Giảm size chữ */
        text-transform: uppercase;
        margin: 0;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.2);
    }
    .main-header p {
        font-size: 0.9rem;
        margin-top: 5px;
        opacity: 0.9;
        letter-spacing: 1px;
    }

    /* 2. Thẻ Card Compact (Gần nhau hơn) */
    .step-card {
        background-color: white;
        border-radius: 10px;
        padding: 15px; /* Giảm padding trong thẻ */
        margin-bottom: 10px; /* Giảm khoảng cách giữa các thẻ */
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        border: 1px solid #e0e0e0;
    }
    
    .card-header {
        display: flex;
        align-items: center;
        margin-bottom: 10px;
        border-bottom: 1px solid #eee;
        padding-bottom: 5px;
    }
    .step-number {
        background: #00796b;
        color: white;
        width: 24px; /* Nhỏ lại */
        height: 24px;
        border-radius: 50%;
        display: flex;
        justify-content: center;
        align-items: center;
        font-weight: bold;
        font-size: 0.9rem;
        margin-right: 8px;
    }
    .step-title {
        font-size: 1rem;
        font-weight: 700;
        color: #004d40;
    }

    /* 3. Nút bấm & Input tối ưu */
    .stButton > button {
        background: #009688;
        color: white;
        width: 100%;
        padding: 0.5rem;
        border-radius: 6px;
        font-weight: bold;
        border: none;
    }
    .stButton > button:hover {
        background: #00796b;
        color: white;
    }
    
    /* Upload gọn */
    .stFileUploader {
        padding: 5px; 
    }
    
    /* Thu nhỏ khoảng cách chung của Streamlit */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
    }
    
    /* Box kiểm tra lỗi */
    .validation-box {
        background: #f1f8e9;
        border: 1px solid #c5e1a5;
        padding: 10px;
        border-radius: 5px;
        margin-top: 10px;
        font-size: 0.9rem;
    }
</style>
""", unsafe_allow_html=True)

# ==================== LOGIC XỬ LÝ (CORE) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

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
    text_content = " ".join([get_pure_text(b) for b in q_blocks])
    options_found = re.findall(r'\b([A-D])[\.\)]', text_content)
    unique_opts = set(opt.upper() for opt in options_found)
    missing = []
    for char in ['A', 'B', 'C', 'D']:
        if char not in unique_opts: missing.append(char)
    has_correct = False
    for block in q_blocks:
        runs = block.getElementsByTagNameNS(W_NS, "r")
        for r in runs:
            if is_answer_marked(r):
                t_nodes = r.getElementsByTagNameNS(W_NS, "t")
                t_val = "".join([t.firstChild.nodeValue for t in t_nodes if t.firstChild])
                if t_val.strip(): has_correct = True; break
        if has_correct: break
    return missing, has_correct

def fix_floating_images_in_xml(doc_xml_str):
    dom = minidom.parseString(doc_xml_str)
    anchors = dom.getElementsByTagName("wp:anchor")
    count = 0
    for anchor in reversed(anchors):
        inline = dom.createElement("wp:inline")
        valid_children = ["wp:extent", "wp:effectExtent", "wp:docPr", "wp:cNvGraphicFramePr", "a:graphic"]
        for child in list(anchor.childNodes):
            if child.nodeName in valid_children: inline.appendChild(child.cloneNode(True))
            elif child.localName == "graphic": inline.appendChild(child.cloneNode(True))
        anchor.parentNode.replaceChild(inline, anchor)
        count += 1
    return dom.toxml(), count

def validate_document(blocks):
    errors = []
    warnings = []
    questions = []
    current_q = []
    q_num_map = {}
    
    for block in blocks:
        text = get_pure_text(block)
        m = re.match(r'^Câu\s*(\d+)', text, re.IGNORECASE)
        if m:
            if current_q: questions.append(current_q)
            current_q = [block]
            q_num_map[len(questions)] = m.group(1)
        else:
            if current_q: current_q.append(block)
    if current_q: questions.append(current_q)
    
    for idx, q_blocks in enumerate(questions):
        q_label = f"Câu {q_num_map.get(idx, 'Unknown')}"
        # Check hình ảnh
        for b in q_blocks:
            if b.getElementsByTagName("wp:anchor"):
                warnings.append(f"{q_label}")

        q_text = " ".join([get_pure_text(b) for b in q_blocks])
        # Check trắc nghiệm
        if re.search(r'\bA[\.\)]', q_text) and re.search(r'\bD[\.\)]', q_text):
            missing, has_correct = check_mcq_options(q_blocks)
            if missing: errors.append(f"❌ {q_label}: Thiếu {', '.join(missing)}")
            if not has_correct: errors.append(f"❌ {q_label}: Chưa tô đáp án")
        # Check tự luận
        elif "ĐS" in q_text or "đs" in q_text:
            has_red_ds = False
            for b in q_blocks:
                runs = b.getElementsByTagNameNS(W_NS, "r")
                for r in runs:
                    if is_answer_marked(r): has_red_ds = True; break
            if not has_red_ds: errors.append(f"❌ {q_label}: 'ĐS' chưa tô đỏ")
            
    return errors, warnings

# ... [Giữ nguyên các hàm update_question_label, process_document_final...] ...
# (Để tiết kiệm không gian hiển thị code, tôi dùng lại logic cũ cho phần xử lý lõi)
# ... Bạn hãy giữ nguyên các hàm xử lý XML từ phiên bản trước ...
def update_question_label(paragraph, new_number):
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    for t in t_nodes:
        if not t.firstChild: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)(Câu\s*)(\d+)([\.:])?', txt, re.IGNORECASE)
        if m:
            t.firstChild.nodeValue = f"{m.group(1)}Câu {new_number}{m.group(4) or '.'}{txt[m.end():]}"
            break

def get_text_with_formatting(block):
    texts = []
    is_correct = False
    runs = block.getElementsByTagNameNS(W_NS, "r")
    for r in runs:
        t_nodes = r.getElementsByTagNameNS(W_NS, "t")
        for t in t_nodes:
            if t.firstChild and t.firstChild.nodeValue:
                texts.append(t.firstChild.nodeValue)
                if is_answer_marked(r) and t.firstChild.nodeValue.strip(): is_correct = True
    return "".join(texts).strip(), is_correct

def update_mcq_label(paragraph, new_label):
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    new_letter = new_label[0].upper()
    for t in t_nodes:
        if not t.firstChild: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)([A-D])([\.\)])?', txt, re.IGNORECASE)
        if m:
            t.firstChild.nodeValue = f"{m.group(1)}{new_letter}{m.group(3) or '.'}{txt[m.end():]}"
            break

def extract_part3_answer(block):
    runs = block.getElementsByTagNameNS(W_NS, "r")
    full_text = ""
    for r in runs:
        t_nodes = r.getElementsByTagNameNS(W_NS, "t")
        for t in t_nodes:
            if t.firstChild and t.firstChild.nodeValue: full_text += t.firstChild.nodeValue
    match = re.search(r'ĐS\s*[:\.]\s*(.+)', full_text, re.IGNORECASE)
    has_red = False
    for r in runs:
        if is_answer_marked(r): has_red = True; break
    if match and has_red: return match.group(1).strip()
    return None

def process_mcq_question(q_blocks):
    header, options = [], []
    for block in q_blocks:
        if re.match(r'^\s*[A-D][\.\)]', get_pure_text(block), re.IGNORECASE): options.append(block)
        else: header.append(block)
    if len(options) < 2: return q_blocks, ""
    orig_idx = -1
    for idx, opt in enumerate(options):
        _, is_cor = get_text_with_formatting(opt)
        if is_cor: orig_idx = idx; break
    perm = list(range(len(options)))
    random.shuffle(perm)
    shuffled_opts = [options[i] for i in perm]
    new_char = ""
    if orig_idx != -1:
        try: new_char = ["A", "B", "C", "D", "E", "F"][perm.index(orig_idx)]
        except: pass
    letters = ["A", "B", "C", "D", "E", "F"]
    for idx, opt in enumerate(shuffled_opts):
        l = letters[idx] if idx < len(letters) else "Z"
        update_mcq_label(opt, f"{l}.")
    return header + shuffled_opts, new_char

def parse_questions(blocks):
    questions, cur_q = [], []
    for block in blocks:
        if re.match(r'^Câu\s*\d+', get_pure_text(block), re.IGNORECASE):
            if cur_q: questions.append(cur_q)
            cur_q = [block]
        else:
            if cur_q: cur_q.append(block)
    if cur_q: questions.append(cur_q)
    return questions

def find_part_index(blocks, p_num):
    pat = re.compile(rf'PHẦN\s*{p_num}\b', re.IGNORECASE)
    for i, b in enumerate(blocks):
        if pat.search(get_pure_text(b)): return i
    return -1

def process_document_final(file_bytes, num_versions, filename_prefix, auto_fix_img, shuffle_mode="auto"):
    input_buffer = io.BytesIO(file_bytes)
    zip_in = zipfile.ZipFile(input_buffer, 'r')
    doc_xml = zip_in.read("word/document.xml").decode('utf-8')
    if auto_fix_img: doc_xml, _ = fix_floating_images_in_xml(doc_xml)
    
    dom = minidom.parseString(doc_xml)
    all_keys = []
    zip_out_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_out_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_final:
        for i in range(num_versions):
            v_name = f"{101 + i}"
            dom_v = minidom.parseString(doc_xml)
            body_v = dom_v.getElementsByTagNameNS(W_NS, "body")[0]
            blocks_v = [n for n in body_v.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName in ["p", "tbl"]]
            
            parts = {"intro": [], "p1": [], "p2": [], "p3": []}
            if shuffle_mode == "auto":
                p1_i = find_part_index(blocks_v, 1)
                p2_i = find_part_index(blocks_v, 2)
                p3_i = find_part_index(blocks_v, 3)
                cur = 0
                if p1_i != -1:
                    parts["intro"] = blocks_v[cur:p1_i+1]
                    cur = p1_i+1
                    e1 = p2_i if p2_i!=-1 else (p3_i if p3_i!=-1 else len(blocks_v))
                    parts["p1"] = blocks_v[cur:e1]
                    cur = e1
                else:
                    parts["p1"] = blocks_v
                    cur = len(blocks_v)
                if p2_i != -1:
                    e2 = p3_i if p3_i!=-1 else len(blocks_v)
                    parts["p2"] = blocks_v[cur:e2]
                    cur = e2
                if p3_i != -1: parts["p3"] = blocks_v[cur:]
            elif shuffle_mode == "mcq":
                parts["p1"] = blocks_v
            elif shuffle_mode == "tf":
                parts["p2"] = blocks_v
                
            final_layout = []
            final_layout.extend(parts["intro"])
            ans_key = {"Mã đề": v_name}
            g_idx = 1
            if parts["p1"]:
                qs = parse_questions(parts["p1"])
                random.shuffle(qs)
                for q in qs:
                    update_question_label(q[0], g_idx)
                    pb, ans = process_mcq_question(q)
                    final_layout.extend(pb)
                    if ans: ans_key[f"Câu {g_idx}"] = ans
                    g_idx += 1
            if parts["p2"]:
                final_layout.extend([parts["p2"][0]] if parts["p2"] else [])
                qs = parse_questions(parts["p2"][1:] if len(parts["p2"])>1 else [])
                random.shuffle(qs)
                for q in qs:
                    update_question_label(q[0], g_idx)
                    final_layout.extend(q)
                    g_idx += 1
            if parts["p3"]:
                final_layout.extend([parts["p3"][0]] if parts["p3"] else [])
                qs = parse_questions(parts["p3"][1:] if len(parts["p3"])>1 else [])
                random.shuffle(qs)
                for q in qs:
                    update_question_label(q[0], g_idx)
                    val = None
                    for b in q: 
                        val = extract_part3_answer(b)
                        if val: break
                    final_layout.extend(q)
                    if val: ans_key[f"Câu {g_idx}"] = val
                    g_idx += 1
            while body_v.hasChildNodes(): body_v.removeChild(body_v.firstChild)
            for b in final_layout: body_v.appendChild(b)
            ver_io = io.BytesIO()
            with zipfile.ZipFile(ver_io, 'w', zipfile.ZIP_DEFLATED) as z_ver:
                for item in zip_in.infolist():
                    if item.filename == "word/document.xml": z_ver.writestr(item, dom_v.toxml().encode('utf-8'))
                    else: z_ver.writestr(item, zip_in.read(item.filename))
            zip_final.writestr(f"{filename_prefix}_{v_name}.docx", ver_io.getvalue())
            all_keys.append(ans_key)
    df = pd.DataFrame(all_keys)
    cols = list(df.columns)
    if "Mã đề" in cols: cols.remove("Mã đề")
    q_cols = sorted(cols, key=lambda s: int(re.search(r'(\d+)', s).group(1)) if re.search(r'(\d+)', s) else 0)
    df = df.reindex(columns=["Mã đề"] + q_cols)
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='DapAn')
    return zip_out_buffer.getvalue(), excel_buf.getvalue()

# ==================== MAIN UI ====================

def main():
    # 1. HEADER (70% size)
    st.markdown("""
    <div class="main-header">
        <h1>TRƯỜNG THPT MINH ĐỨC</h1>
        <p>APP TRỘN ĐỀ 2025</p>
    </div>
    """, unsafe_allow_html=True)

    # LAYOUT 2 CỘT: Cột Trái (Upload + Check) | Cột Phải (Cấu hình)
    # Gap="small" để tối ưu không gian
    col_left, col_right = st.columns([1, 1], gap="small")

    # --- CỘT TRÁI ---
    with col_left:
        # THẺ HƯỚNG DẪN THU GỌN (50% size, Collapsible)
        with st.expander("ℹ️ Tải mẫu & Hướng dẫn (Bấm để xem)", expanded=False):
            c_d1, c_d2 = st.columns([1, 1])
            with c_d1:
                st.info("Quy định: Câu 1, Đáp án tô đỏ/gạch chân.")
            with c_d2:
                st.link_button("📥 Tải File Mẫu", "https://drive.google.com/file/d/1_2zhqxwoMQ-AINMfCqy6QbZyGU4Skg3n/view?usp=sharing")

        # THẺ BƯỚC 1: UPLOAD & KIỂM TRA
        st.markdown('<div class="step-card">', unsafe_allow_html=True)
        st.markdown('<div class="card-header"><div class="step-number">1</div><div class="step-title">Upload & Kiểm tra</div></div>', unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader("Chọn file (.docx)", type=["docx"], label_visibility="collapsed")
        
        if uploaded_file:
            st.session_state['file_bytes'] = uploaded_file.getvalue()
            
            # NÚT KIỂM TRA FILE
            if st.button("🔍 Kiểm tra lỗi & Cấu trúc"):
                try:
                    input_buffer = io.BytesIO(st.session_state['file_bytes'])
                    zip_in = zipfile.ZipFile(input_buffer, 'r')
                    doc_xml = zip_in.read("word/document.xml").decode('utf-8')
                    dom = minidom.parseString(doc_xml)
                    body = dom.getElementsByTagNameNS(W_NS, "body")[0]
                    blocks = [n for n in body.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName in ["p", "tbl"]]
                    
                    errors, warnings = validate_document(blocks)
                    
                    # LOGIC HIỂN THỊ LỖI / SỬA LỖI
                    st.markdown('<div class="validation-box">', unsafe_allow_html=True)
                    
                    if not errors and not warnings:
                        st.success("✅ File Tuyệt Vời! Chuẩn cấu trúc.")
                        st.session_state['is_valid'] = True
                        st.session_state['auto_fix_img'] = False
                    else:
                        if errors:
                            st.error(f"Phát hiện {len(errors)} lỗi nghiêm trọng (Cần sửa trong Word):")
                            for e in errors: st.write(e)
                            st.session_state['is_valid'] = False
                        else:
                            st.session_state['is_valid'] = True # Chỉ có warning thì vẫn cho trộn

                        if warnings:
                            st.warning(f"⚠️ {len(warnings)} hình ảnh bị trôi (Floating).")
                            st.write("👉 Hệ thống sẽ TỰ ĐỘNG SỬA khi bấm nút Trộn bên phải.")
                            st.session_state['auto_fix_img'] = True
                    
                    st.markdown('</div>', unsafe_allow_html=True)

                except Exception as e:
                    st.error("Lỗi đọc file. File bị hỏng.")

        st.markdown('</div>', unsafe_allow_html=True)

    # --- CỘT PHẢI ---
    with col_right:
        # BƯỚC 2: CẤU HÌNH (GỘP BƯỚC 2 & 3 vào 1 thẻ để tiết kiệm chỗ)
        st.markdown('<div class="step-card">', unsafe_allow_html=True)
        st.markdown('<div class="card-header"><div class="step-number">2</div><div class="step-title">Cấu hình & Xử lý</div></div>', unsafe_allow_html=True)
        
        c_cfg1, c_cfg2 = st.columns(2)
        with c_cfg1:
            mode = st.selectbox("Kiểu trộn", ["auto", "mcq", "tf"], format_func=lambda x: {"auto":"Tự động", "mcq":"Trắc nghiệm", "tf":"Đúng/Sai"}[x])
        with c_cfg2:
            num_mix = st.number_input("Số đề", 1, 50, 4)
            
        st.write("")
        # NÚT XỬ LÝ CHÍNH
        if st.button("🚀 TRỘN ĐỀ & TẢI XUỐNG"):
            if 'file_bytes' in st.session_state and st.session_state.get('is_valid', True):
                with st.spinner("Đang xử lý..."):
                    # Tự động fix ảnh nếu có warning
                    do_fix = st.session_state.get('auto_fix_img', True)
                    try:
                        z_data, e_data = process_document_final(
                            st.session_state['file_bytes'], num_mix, "KiemTra", do_fix, mode
                        )
                        st.success("Xong!")
                        d1, d2 = st.columns(2)
                        with d1:
                            st.download_button("📥 File Đề", z_data, "De_Tron.zip", "application/zip", use_container_width=True)
                        with d2:
                            st.download_button("📊 Đáp án", e_data, "Dap_An.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                    except Exception as e:
                        st.error(f"Lỗi: {e}")
            else:
                st.warning("Vui lòng Upload & Kiểm tra file trước (Cột trái).")
                
        st.markdown('</div>', unsafe_allow_html=True)

    # Footer
    st.markdown('<div style="text-align:center; color: #aaa; font-size: 0.8rem; margin-top:10px;">© 2025 Phan Trường Duy</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
