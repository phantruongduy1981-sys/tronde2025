"""
Trộn Đề Word Online - AIOMT Premium (UI Update 3 Steps)
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

# ==================== CẤU HÌNH TRANG & CSS ====================

st.set_page_config(
    page_title="App Trộn Đề 2025",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS: Teal Theme & 3-Step Layout
st.markdown("""
<style>
    /* 1. Header Nhỏ gọn (50% kích thước cũ) */
    .main-header {
        text-align: center;
        padding: 0.8rem 0; /* Giảm padding */
        background: linear-gradient(135deg, #009688 0%, #00796b 100%);
        border-radius: 8px;
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    .main-header h1 {
        font-family: 'Arial', sans-serif;
        font-weight: 700;
        font-size: 1.4rem; /* Font nhỏ hơn */
        margin: 0;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .main-header p {
        font-size: 0.8rem;
        margin: 0;
        opacity: 0.9;
        font-style: italic;
    }

    /* 2. Step Badges (Số 1, 2, 3 tròn) */
    .step-container {
        display: flex;
        align-items: center;
        margin-bottom: 10px;
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
        font-weight: bold;
        margin-right: 10px;
        font-size: 0.9rem;
    }
    .step-title {
        font-weight: 700;
        color: #333;
        font-size: 1rem;
    }

    /* 3. Instruction Box (Giống hình mẫu) */
    .instruction-box {
        background-color: #e0f2f1; /* Xanh rất nhạt */
        border: 1px solid #b2dfdb;
        border-radius: 8px;
        padding: 15px;
        margin-bottom: 20px;
        font-size: 0.9rem;
    }
    .instruction-part {
        margin-bottom: 5px;
        color: #00695c;
        font-weight: 600;
    }
    .warning-box {
        background-color: #fff8e1; /* Vàng nhạt */
        border: 1px solid #ffe082;
        border-radius: 6px;
        padding: 10px;
        margin-top: 10px;
        color: #795548;
        font-size: 0.85rem;
    }
    .highlight-red { color: #d32f2f; font-weight: bold; }
    .highlight-code { background: #eee; padding: 2px 4px; border-radius: 3px; font-family: monospace; }

    /* 4. Card Style */
    .css-card {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        border: 1px solid #f0f0f0;
        height: 100%;
    }

    /* 5. Nút bấm Download File Mẫu */
    .btn-sample {
        display: inline-block;
        background-color: #00897b;
        color: white !important;
        padding: 6px 12px;
        border-radius: 4px;
        text-decoration: none;
        font-size: 0.8rem;
        font-weight: bold;
        float: right;
    }
    
    /* 6. Main Action Button */
    .stButton > button {
        background: linear-gradient(to right, #009688, #00796b);
        color: white;
        border: none;
        font-weight: bold;
        width: 100%;
        padding: 0.7rem;
        border-radius: 6px;
        font-size: 1.1rem;
    }
    .stButton > button:hover {
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        color: white;
    }
    
    /* File Uploader Custom */
    .stFileUploader {
        border: 2px dashed #009688;
        background: #f0fdfa;
        border-radius: 8px;
        padding: 10px;
    }
</style>
""", unsafe_allow_html=True)

# ==================== CORE LOGIC (GIỮ NGUYÊN) ====================

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
        for b in q_blocks:
            if b.getElementsByTagName("wp:anchor"):
                warnings.append(f"{q_label}: Chứa hình ảnh trôi nổi (Floating).")
        q_text = " ".join([get_pure_text(b) for b in q_blocks])
        if re.search(r'\bA[\.\)]', q_text) and re.search(r'\bD[\.\)]', q_text):
            missing, has_correct = check_mcq_options(q_blocks)
            if missing: errors.append(f"❌ {q_label}: Thiếu phương án {', '.join(missing)}")
            if not has_correct: errors.append(f"❌ {q_label}: Chưa chọn đáp án đúng (Tô đỏ/Gạch chân)")
        elif "ĐS" in q_text or "đs" in q_text:
            has_red_ds = False
            for b in q_blocks:
                runs = b.getElementsByTagNameNS(W_NS, "r")
                for r in runs:
                    if is_answer_marked(r): has_red_ds = True; break
            if not has_red_ds: errors.append(f"❌ {q_label}: Đáp án 'ĐS:...' chưa được tô đỏ.")
    return errors, warnings

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
            
            # Logic chia phần dựa trên Mode
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
                    parts["p1"] = blocks_v # Mặc định nếu ko tìm thấy phần nào -> P1 (MCQ)
                    cur = len(blocks_v)
                if p2_i != -1:
                    e2 = p3_i if p3_i!=-1 else len(blocks_v)
                    parts["p2"] = blocks_v[cur:e2]
                    cur = e2
                if p3_i != -1: parts["p3"] = blocks_v[cur:]
            
            elif shuffle_mode == "mcq":
                parts["p1"] = blocks_v # Coi toàn bộ là MCQ
            
            elif shuffle_mode == "tf":
                parts["p2"] = blocks_v # Coi toàn bộ là Đúng Sai
                
            final_layout = []
            final_layout.extend(parts["intro"])
            ans_key = {"Mã đề": v_name}
            g_idx = 1
            
            # P1 Mix
            if parts["p1"]:
                qs = parse_questions(parts["p1"])
                random.shuffle(qs)
                for q in qs:
                    update_question_label(q[0], g_idx)
                    pb, ans = process_mcq_question(q)
                    final_layout.extend(pb)
                    if ans: ans_key[f"Câu {g_idx}"] = ans
                    g_idx += 1
            # P2 Mix
            if parts["p2"]:
                final_layout.extend([parts["p2"][0]] if parts["p2"] else [])
                qs = parse_questions(parts["p2"][1:] if len(parts["p2"])>1 else [])
                random.shuffle(qs)
                for q in qs:
                    update_question_label(q[0], g_idx)
                    final_layout.extend(q)
                    g_idx += 1
            # P3 Mix
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
    # HEADER NHỎ GỌN (50% so với trước)
    st.markdown("""
    <div class="main-header">
        <h1>TRƯỜNG THPT MINH ĐỨC</h1>
        <p>APP TRỘN ĐỀ 2025</p>
    </div>
    """, unsafe_allow_html=True)

    # LAYOUT 2 CỘT (Trái: Upload - Phải: Cấu hình)
    col_left, col_right = st.columns([1, 1], gap="medium")

    # --- CỘT TRÁI ---
    with col_left:
        st.markdown('<div class="css-card">', unsafe_allow_html=True)
        # Bước 1
        st.markdown("""
        <div class="step-container">
            <div class="step-badge">1</div>
            <div class="step-title">Chọn file đề Word (*.docx)</div>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader("Kéo thả file hoặc click để chọn", type=["docx"], label_visibility="collapsed")
        
        # Logic Upload
        if uploaded_file:
            st.session_state['file_bytes'] = uploaded_file.getvalue()
            st.success(f"✅ Đã nhận: {uploaded_file.name}")
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Preview hoặc Log lỗi (nếu cần) hiển thị ở đây
        if 'file_bytes' in st.session_state:
            try:
                input_buffer = io.BytesIO(st.session_state['file_bytes'])
                zip_in = zipfile.ZipFile(input_buffer, 'r')
                doc_xml = zip_in.read("word/document.xml").decode('utf-8')
                dom = minidom.parseString(doc_xml)
                body = dom.getElementsByTagNameNS(W_NS, "body")[0]
                blocks = [n for n in body.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName in ["p", "tbl"]]
                errors, warnings = validate_document(blocks)
                
                if errors:
                    st.error("🚫 File lỗi cấu trúc:")
                    for e in errors: st.write(e)
                    st.session_state['is_valid'] = False
                else:
                    st.session_state['is_valid'] = True
                    if warnings:
                        with st.expander("⚠️ Cảnh báo hình ảnh"):
                            for w in warnings: st.write(w)
            except: pass

    # --- CỘT PHẢI ---
    with col_right:
        st.markdown('<div class="css-card">', unsafe_allow_html=True)
        
        # HƯỚNG DẪN (Full Style như hình)
        sample_url = "https://drive.google.com/file/d/1_2zhqxwoMQ-AINMfCqy6QbZyGU4Skg3n/view?usp=sharing"
        st.markdown(f"""
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
            <div style="font-weight:bold; color:#00796b;">📝 Hướng dẫn & Cấu trúc</div>
            <a href="{sample_url}" target="_blank" class="btn-sample">📥 Tải mẫu</a>
        </div>
        
        <div class="instruction-box">
            <div class="instruction-part">PHẦN 1: <span style="font-weight:400; color:#333;">Trắc nghiệm (A. B. C. D.)</span></div>
            <div class="instruction-part">PHẦN 2: <span style="font-weight:400; color:#333;">Đúng/Sai (a) b) c) d))</span></div>
            <div class="instruction-part">PHẦN 3: <span style="font-weight:400; color:#333;">Trả lời ngắn</span></div>
            
            <div class="warning-box">
                <div style="font-weight:bold; margin-bottom:5px;">⚠️ Lưu ý quan trọng:</div>
                <div>• Mỗi câu hỏi bắt đầu bằng <span class="highlight-code">Câu 1.</span>, <span class="highlight-code">Câu 2.</span></div>
                <div>• Đáp án trắc nghiệm: <span class="highlight-code">A.</span> <span class="highlight-code">B.</span> (viết hoa + dấu chấm)</div>
                <div>• Đáp án đúng/sai: <span class="highlight-code">a)</span> <span class="highlight-code">b)</span> (viết thường + dấu ngoặc)</div>
                <div>• <b>Đáp án đúng:</b> Phải <span style="text-decoration:underline;">gạch chân</span> hoặc <span class="highlight-red">tô màu đỏ</span>.</div>
                <div style="margin-top:5px; border-top:1px dashed #ccc; padding-top:5px;">
                    • <b>Đáp án Phần 3 (Mới):</b> Ghi <span class="highlight-red">ĐS: Kết quả</span> và tô màu đỏ.
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # BƯỚC 2: Chọn kiểu trộn
        st.markdown("""
        <div class="step-container">
            <div class="step-badge">2</div>
            <div class="step-title">Chọn kiểu trộn</div>
        </div>
        """, unsafe_allow_html=True)
        
        # Radio button ngang
        shuffle_mode = st.radio(
            "Chọn kiểu trộn",
            options=["auto", "mcq", "tf"],
            format_func=lambda x: {
                "auto": "🔄 Tự động (Theo PHẦN 1,2,3)", 
                "mcq": "📝 Toàn bộ Trắc nghiệm", 
                "tf": "✅ Toàn bộ Đúng/Sai"
            }[x],
            label_visibility="collapsed"
        )
        
        st.write("") # Spacer

        # BƯỚC 3: Cấu hình số lượng
        st.markdown("""
        <div class="step-container">
            <div class="step-badge">3</div>
            <div class="step-title">Số mã đề cần tạo</div>
        </div>
        """, unsafe_allow_html=True)
        
        c1, c2 = st.columns([1, 2])
        with c1:
            num_versions = st.number_input("Số lượng", 1, 50, 4, label_visibility="collapsed")
        with c2:
            st.caption("● 1 mã → File Word  |  ● Nhiều mã → File ZIP")
            
        st.write("")
        auto_fix = st.checkbox("Sửa lỗi hình ảnh", value=True)

        # NÚT XỬ LÝ
        if st.button("🚀 Trộn đề & Tải xuống"):
            if st.session_state.get('is_valid') and 'file_bytes' in st.session_state:
                with st.spinner("Đang xử lý..."):
                    try:
                        z_data, e_data = process_document_final(
                            st.session_state['file_bytes'], 
                            num_versions, 
                            "KiemTra", 
                            auto_fix,
                            shuffle_mode
                        )
                        st.success("Thành công!")
                        d1, d2 = st.columns(2)
                        with d1:
                            st.download_button("📥 Tải Đề (ZIP)", z_data, "De_Tron.zip", "application/zip", use_container_width=True)
                        with d2:
                            st.download_button("📊 Đáp án (Excel)", e_data, "Dap_An.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                    except Exception as e:
                        st.error(f"Lỗi: {e}")
            else:
                st.error("Vui lòng tải file hợp lệ ở Bước 1")

        st.markdown('</div>', unsafe_allow_html=True)

    # Footer
    st.markdown('<div style="text-align:center; color: #aaa; margin-top:20px; font-size:0.8rem;">© 2025 Phan Trường Duy - THPT Minh Đức</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
