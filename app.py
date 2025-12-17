"""
Trộn Đề Word Online - AIOMT Premium (Vibrant 2025 UI)
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
    page_icon="📝", # Thay icon xúc sắc bằng icon giấy bút cho nghiêm túc hơn
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom CSS: Màu sắc nổi bật, bố cục sát lề
st.markdown("""
<style>
    /* 1. Đẩy nội dung lên sát mép trên */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
        max-width: 900px;
    }

    /* 2. Header mới: Màu Gradient Đỏ Cam Nổi Bật */
    .main-header {
        text-align: center;
        padding: 1.5rem 0;
        background: linear-gradient(135deg, #FF416C 0%, #FF4B2B 100%); /* Màu đỏ cam rực rỡ */
        border-radius: 12px;
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 15px rgba(255, 75, 43, 0.3);
    }
    .main-header h1 {
        margin: 0;
        font-family: 'Arial', sans-serif;
        font-weight: 800; /* Chữ đậm */
        font-size: 2.2rem;
        text-transform: uppercase;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        letter-spacing: 1px;
    }
    .main-header h2 {
        margin: 5px 0 0 0;
        font-family: 'Arial', sans-serif;
        font-size: 1.5rem;
        font-weight: 600;
        opacity: 0.95;
    }
    
    /* 3. Style cho các "Thẻ" (Cards) */
    .card-box {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border: 1px solid #eee;
        margin-bottom: 20px;
    }
    
    /* 4. Nút bấm đẹp hơn */
    .stButton > button {
        width: 100%;
        background: linear-gradient(to right, #FF416C, #FF4B2B);
        color: white;
        font-weight: bold;
        border: none;
        padding: 0.8rem;
        border-radius: 8px;
        transition: all 0.3s;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(255, 65, 108, 0.4);
        color: white;
    }

    /* 5. Style cho Expander (Hướng dẫn) */
    .streamlit-expanderHeader {
        background-color: #fff0f0;
        color: #d63031;
        font-weight: bold;
        border-radius: 5px;
    }
    
    /* 6. Nền trang web */
    .stApp {
        background-color: #f8f9fa;
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

def process_document_final(file_bytes, num_versions, filename_prefix, auto_fix_img):
    # CHỈ DÙNG BytesIO từ file_bytes đã đọc, KHÔNG gọi .read() lại
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
            
            # Logic trộn đề như cũ
            dom_v = minidom.parseString(doc_xml)
            body_v = dom_v.getElementsByTagNameNS(W_NS, "body")[0]
            blocks_v = [n for n in body_v.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName in ["p", "tbl"]]
            
            p1_i = find_part_index(blocks_v, 1)
            p2_i = find_part_index(blocks_v, 2)
            p3_i = find_part_index(blocks_v, 3)
            
            parts = {"intro": [], "p1": [], "p2": [], "p3": []}
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
    # 1. HEADER CHÍNH
    st.markdown("""
    <div class="main-header">
        <h1>TRƯỜNG THPT MINH ĐỨC</h1>
        <h2>APP TRỘN ĐỀ 2025</h2>
    </div>
    """, unsafe_allow_html=True)

    # 2. KHU VỰC HƯỚNG DẪN & FILE MẪU
    st.markdown('<div class="card-box">', unsafe_allow_html=True)
    st.markdown("### 📋 Hướng dẫn sử dụng")
    with st.expander("Xem quy định file & Tải mẫu", expanded=False):
        c_inst1, c_inst2 = st.columns([2, 1])
        with c_inst1:
            st.info("""
            **Quy định soạn thảo Word:**
            1. **Câu hỏi:** Bắt đầu bằng `Câu 1`, `Câu 2`...
            2. **Trắc nghiệm:** Có đủ 4 đáp án A. B. C. D.
            3. **Đáp án đúng:** Tô **màu đỏ** hoặc **gạch chân**.
            4. **Tự luận:** Đáp án ghi `ĐS: kết quả` (Tô đỏ).
            """)
        with c_inst2:
            st.write("📥 **Tải file mẫu chuẩn:**")
            sample_url = "https://drive.google.com/file/d/1_2zhqxwoMQ-AINMfCqy6QbZyGU4Skg3n/view?usp=sharing"
            st.link_button("Download File Mẫu", sample_url, type="secondary")
    st.markdown('</div>', unsafe_allow_html=True)

    # 3. KHU VỰC UPLOAD
    st.markdown('<div class="card-box">', unsafe_allow_html=True)
    st.markdown("### 📂 Bước 1: Tải đề gốc")
    uploaded_file = st.file_uploader("Kéo thả file Word (.docx) vào đây", type=["docx"])
    
    # 4. KHU VỰC CẤU HÌNH & XỬ LÝ
    if uploaded_file:
        file_bytes = uploaded_file.getvalue() # ĐỌC FILE 1 LẦN DUY NHẤT
        
        # Validate sơ bộ
        try:
            input_buffer = io.BytesIO(file_bytes)
            zip_in = zipfile.ZipFile(input_buffer, 'r')
            doc_xml = zip_in.read("word/document.xml").decode('utf-8')
            dom = minidom.parseString(doc_xml)
            body = dom.getElementsByTagNameNS(W_NS, "body")[0]
            blocks = [n for n in body.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName in ["p", "tbl"]]
            errors, warnings = validate_document(blocks)
            
            if errors:
                st.error("🚫 File chưa đúng quy định:")
                for e in errors: st.write(e)
                st.stop()
            
            if warnings:
                with st.expander("⚠️ Cảnh báo hình ảnh (Click để xem)", expanded=False):
                    for w in warnings: st.warning(w)
            
            st.success("✅ File hợp lệ! Mời cấu hình bên dưới.")
            st.markdown('</div>', unsafe_allow_html=True) # Đóng thẻ card upload
            
            # Mở thẻ card mới cho cấu hình
            st.markdown('<div class="card-box">', unsafe_allow_html=True)
            st.markdown("### ⚙️ Bước 2: Cấu hình trộn")
            
            c1, c2 = st.columns(2)
            with c1:
                num_versions = st.number_input("Số lượng đề cần tạo", min_value=1, max_value=50, value=4)
                filename_prefix = st.text_input("Mã đề (VD: KiemTra)", value="KiemTra")
            with c2:
                st.write("Tùy chọn nâng cao:")
                auto_fix = st.checkbox("Tự động sửa lỗi hình ảnh trôi nổi", value=True)
                st.info("💡 Tính năng này giúp hình ảnh không bị chạy lung tung khi trộn.")
            
            st.markdown("---")
            if st.button("🚀 TRỘN ĐỀ NGAY", type="primary"):
                with st.spinner("Đang xử lý... Hệ thống đang trộn câu hỏi và đáp án..."):
                    try:
                        # Truyền file_bytes đã đọc
                        z_data, e_data = process_document_final(file_bytes, num_versions, filename_prefix, auto_fix)
                        
                        st.balloons()
                        st.markdown("""
                        <div style="background-color: #d4edda; color: #155724; padding: 15px; border-radius: 5px; text-align: center; margin-bottom: 20px;">
                            <h3>🎉 Xử lý thành công!</h3>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        col_d1, col_d2 = st.columns(2)
                        with col_d1:
                            st.download_button("📥 Tải File Đề (ZIP)", z_data, f"{filename_prefix}_All_Exams.zip", "application/zip", use_container_width=True)
                        with col_d2:
                            st.download_button("📊 Tải Đáp Án (Excel)", e_data, f"{filename_prefix}_Keys.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                            
                    except Exception as e:
                        st.error(f"Lỗi hệ thống: {str(e)}")
            
        except Exception as e:
            st.error(f"File lỗi hoặc bị hỏng: {str(e)}")
            
    st.markdown('</div>', unsafe_allow_html=True) # Đóng thẻ card cấu hình

    # Footer
    st.markdown('<div style="text-align:center; color: #888; margin-top: 20px;">© 2025 Phan Trường Duy - THPT Minh Đức</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
