"""
Trộn Đề Word Online - AIOMT Premium (Core Old + New UI + Excel Fix)
Author: Phan Trường Duy - THPT Minh Đức
"""

import streamlit as st
import re
import random
import zipfile
import io
import pandas as pd
from xml.dom import minidom

# ==================== CẤU HÌNH TRANG ====================

st.set_page_config(
    page_title="App Trộn Đề 2025",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==================== CSS GIAO DIỆN (GIỮ NGUYÊN) ====================
st.markdown("""
<style>
    /* HEADER */
    .main-header {
        text-align: center;
        padding: 2rem 0;
        background: linear-gradient(to right, #009688, #00796b);
        color: white;
        border-radius: 0 0 15px 15px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    }
    .main-header h1 {
        font-family: 'Arial', sans-serif;
        font-weight: 900;
        font-size: 2.8rem;
        text-transform: uppercase;
        margin: 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    .main-header p { font-size: 1.1rem; margin-top: 10px; opacity: 0.9; letter-spacing: 2px; }

    /* STEP CARDS */
    .step-label {
        font-size: 1.1rem;
        font-weight: 700;
        color: #263238;
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
    }

    /* INSTRUCTION */
    .instruction-card {
        background-color: #e0f2f1;
        border-radius: 10px;
        padding: 20px;
        color: #004d40;
        font-size: 0.95rem;
        border: 1px solid #b2dfdb;
    }
    .part-title { font-weight: bold; color: #00796b; display: inline-block; width: 70px; }
    .warning-box {
        background-color: #fff8e1;
        border: 1px solid #ffe082;
        border-radius: 8px;
        padding: 15px;
        margin-top: 15px;
        color: #5d4037;
    }
    .code-tag {
        background-color: #fff;
        padding: 2px 6px;
        border-radius: 4px;
        border: 1px solid #e0e0e0;
        font-family: monospace;
        color: #d84315;
        font-weight: bold;
    }

    /* VERTICAL RADIO */
    div[role="radiogroup"] { display: flex; flex-direction: column; gap: 12px; }
    div[role="radiogroup"] > label {
        width: 100%;
        background-color: white;
        border: 1px solid #cfd8dc;
        border-radius: 8px;
        padding: 15px;
        display: flex;
        align-items: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 0px !important;
    }
    
    /* UPLOAD & BUTTON */
    .stFileUploader { border: 2px dashed #009688; border-radius: 10px; padding: 20px; background-color: white; text-align: center; }
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
    }
    .stButton > button:hover { background: #00796b; }
    .block-container { padding-top: 1rem !important; }
</style>
""", unsafe_allow_html=True)

# ==================== LÕI XỬ LÝ (TỪ CODE CŨ + BỔ SUNG EXCEL) ====================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_text(block):
    """Lấy text từ một block (Code cũ)"""
    texts = []
    t_nodes = block.getElementsByTagNameNS(W_NS, "t")
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue:
            texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def is_answer_marked(node):
    """Kiểm tra tô đỏ hoặc gạch chân (Bổ sung mới để lấy đáp án Excel)"""
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
    """Kiểm tra block có chứa đáp án đúng không"""
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
    """Tô xanh đậm (Code cũ)"""
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
    color_el.setAttributeNS(W_NS, "w:val", "0000FF")
    
    b_list = rPr.getElementsByTagNameNS(W_NS, "b")
    if not b_list:
        b_el = doc.createElementNS(W_NS, "w:b")
        rPr.appendChild(b_el)

def update_mcq_label(paragraph, new_label):
    """Cập nhật nhãn A. (Code cũ)"""
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
    """Cập nhật nhãn a) (Code cũ)"""
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
    """Cập nhật Câu 1. (Code cũ)"""
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    for i, t in enumerate(t_nodes):
        if not t.firstChild: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)(Câu\s*)(\d+)(\.)?', txt, re.IGNORECASE)
        if m:
            t.firstChild.nodeValue = m.group(1) + new_label + txt[m.end():]
            run = t.parentNode
            if run and run.localName == "r": style_run_blue_bold(run)
            break

def find_part_index(blocks, part_number):
    pattern = re.compile(rf'PHẦN\s*{part_number}\b', re.IGNORECASE)
    for i, block in enumerate(blocks):
        if pattern.search(get_text(block)): return i
    return -1

def parse_questions_in_range(blocks, start, end):
    part_blocks = blocks[start:end]
    intro, questions = [], []
    i = 0
    while i < len(part_blocks):
        if re.match(r'^Câu\s*\d+\b', get_text(part_blocks[i])): break
        intro.append(part_blocks[i])
        i += 1
    while i < len(part_blocks):
        text = get_text(part_blocks[i])
        if re.match(r'^Câu\s*\d+\b', text):
            group = [part_blocks[i]]
            i += 1
            while i < len(part_blocks):
                t2 = get_text(part_blocks[i])
                if re.match(r'^Câu\s*\d+\b', t2) or re.match(r'^PHẦN\s*\d\b', t2, re.IGNORECASE): break
                group.append(part_blocks[i])
                i += 1
            questions.append(group)
        else:
            intro.append(part_blocks[i])
            i += 1
    return intro, questions

def shuffle_array(arr):
    out = arr.copy()
    for i in range(len(out) - 1, 0, -1):
        j = random.randint(0, i)
        out[i], out[j] = out[j], out[i]
    return out

# --- CẢI TIẾN: HÀM TRỘN VÀ TRẢ VỀ ĐÁP ÁN ---

def process_mcq_question_with_key(question_blocks):
    """Trộn MCQ và trả về đáp án đúng mới"""
    indices = []
    for i, block in enumerate(question_blocks):
        if re.match(r'^\s*[A-D][\.\)]', get_text(block), re.IGNORECASE): indices.append(i)
    
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
    
    # Xác định đáp án mới
    new_correct_char = ""
    letters = ["A", "B", "C", "D", "E", "F"]
    if original_correct_idx != -1:
        if original_correct_idx in perm: # Safety check
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
    # Logic cũ của bạn cho phần Đúng Sai
    option_indices = {}
    for i, block in enumerate(question_blocks):
        m = re.match(r'^\s*([a-d])\)', get_text(block), re.IGNORECASE)
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
    if "d" in option_indices: middle.append(question_blocks[option_indices["d"]]) # Giữ d
    
    for idx, block in enumerate(middle):
        if idx < 3: update_tf_label(block, f"{['a','b','c'][idx]})")
        
    return before + middle + after

def extract_part3_answer(blocks):
    """Tìm đáp án ĐS: ..."""
    full_text = ""
    has_red = False
    for b in blocks:
        full_text += get_text(b)
        # Check red
        runs = b.getElementsByTagNameNS(W_NS, "r")
        for r in runs:
            if is_answer_marked(r): has_red = True
            
    match = re.search(r'ĐS\s*[:\.]\s*(.+)', full_text, re.IGNORECASE)
    if match and has_red:
        return match.group(1).strip()
    return None

def shuffle_docx_and_get_key(file_bytes, shuffle_mode, version_name):
    """Trộn và trả về (bytes, answer_key)"""
    input_buffer = io.BytesIO(file_bytes)
    # Check zip
    if not zipfile.is_zipfile(input_buffer): raise Exception("File không hợp lệ!")
        
    with zipfile.ZipFile(input_buffer, 'r') as zin:
        doc_xml = zin.read("word/document.xml").decode('utf-8')
        dom = minidom.parseString(doc_xml)
        body = dom.getElementsByTagNameNS(W_NS, "body")[0]
        
        # Save sectPr
        sectPr = None
        if body.lastChild and body.lastChild.localName == 'sectPr':
            sectPr = body.lastChild.cloneNode(True)
            
        blocks = [n for n in body.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName in ["p", "tbl"]]
        
        part1_idx = find_part_index(blocks, 1)
        part2_idx = find_part_index(blocks, 2)
        part3_idx = find_part_index(blocks, 3)
        
        # Cắt phần (Logic cũ)
        parts = {"intro": [], "p1": [], "p2": [], "p3": []}
        cur = 0
        if shuffle_mode == "auto":
            if part1_idx >= 0:
                parts["intro"] = blocks[cur:part1_idx+1]
                cur = part1_idx+1
                end1 = part2_idx if part2_idx >= 0 else (part3_idx if part3_idx >= 0 else len(blocks))
                parts["p1"] = blocks[cur:end1]
                cur = end1
            else:
                parts["p1"] = blocks # Default
                cur = len(blocks)
            if part2_idx >= 0:
                end2 = part3_idx if part3_idx >= 0 else len(blocks)
                parts["p2"] = blocks[cur:end2]
                cur = end2
            if part3_idx >= 0:
                parts["p3"] = blocks[cur:]
        elif shuffle_mode == "mcq":
            parts["p1"] = blocks
        elif shuffle_mode == "tf":
            parts["p2"] = blocks

        final_blocks = []
        final_blocks.extend(parts["intro"])
        answer_key = {"Mã đề": version_name}
        g_idx = 1
        
        # Process P1
        if parts["p1"]:
            intro1, qs1 = parse_questions_in_range(parts["p1"], 0, len(parts["p1"]))
            final_blocks.extend(intro1)
            random.shuffle(qs1)
            for q in qs1:
                update_question_label(q[0], f"Câu {g_idx}.")
                new_q, ans = process_mcq_question_with_key(q)
                final_blocks.extend(new_q)
                if ans: answer_key[f"Câu {g_idx}"] = ans
                g_idx += 1
                
        # Process P2
        if parts["p2"]:
            intro2, qs2 = parse_questions_in_range(parts["p2"], 0, len(parts["p2"]))
            final_blocks.extend(intro2)
            random.shuffle(qs2)
            for q in qs2:
                update_question_label(q[0], f"Câu {g_idx}.")
                new_q = process_tf_question(q)
                final_blocks.extend(new_q)
                g_idx += 1
                
        # Process P3
        if parts["p3"]:
            intro3, qs3 = parse_questions_in_range(parts["p3"], 0, len(parts["p3"]))
            final_blocks.extend(intro3)
            random.shuffle(qs3)
            for q in qs3:
                update_question_label(q[0], f"Câu {g_idx}.")
                # P3 không trộn nội dung, chỉ trộn câu
                final_blocks.extend(q)
                ans = extract_part3_answer(q)
                if ans: answer_key[f"Câu {g_idx}"] = ans
                g_idx += 1
        
        # Rebuild Body
        while body.hasChildNodes(): body.removeChild(body.firstChild)
        for b in final_blocks: body.appendChild(b)
        if sectPr: body.appendChild(sectPr)
        
        new_xml = dom.toxml()
        output_buffer = io.BytesIO()
        with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_xml.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item.filename))
                    
        return output_buffer.getvalue(), answer_key

# ==================== MAIN UI ====================

def main():
    # 1. HEADER
    st.markdown("""
    <div class="main-header">
        <h1>TRƯỜNG THPT MINH ĐỨC</h1>
        <p>APP TRỘN ĐỀ 2025</p>
    </div>
    """, unsafe_allow_html=True)

    col_left, col_right = st.columns([1, 1], gap="medium")

    # --- CỘT TRÁI ---
    with col_left:
        # HƯỚNG DẪN
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
<div style="margin-top:5px;">
<span class="part-title">PHẦN 1:</span> Trắc nghiệm (A. B. C. D.)
</div>
<div>
<span class="part-title">PHẦN 2:</span> Đúng/Sai (a) b) c) d))
</div>
<div>
<span class="part-title">PHẦN 3:</span> Trả lời ngắn (ĐS:...)
</div>

<div class="warning-box">
<div style="font-weight:bold; color:#e65100; margin-bottom:5px;">⚠️ Lưu ý quan trọng:</div>
<ul style="margin-bottom: 0; padding-left: 20px;">
<li>Câu hỏi bắt đầu bằng <span class="code-tag">Câu 1.</span>, <span class="code-tag">Câu 2.</span></li>
<li>Đáp án đúng phải <span style="text-decoration:underline;">gạch chân</span> hoặc <span style="color:blue; font-weight:bold;">tô màu đỏ</span>.</li>
<li style="margin-top:5px; border-top:1px dashed #ccc; padding-top:5px;">
<b>Đáp án Phần 3:</b> Ghi <span style="color:red; font-weight:bold;">ĐS: Kết quả</span> và tô đỏ.
</li>
</ul>
</div>
</div>
""", unsafe_allow_html=True)
        
        # UPLOAD
        st.markdown('<div class="step-label"><div class="step-badge">1</div>Chọn file đề Word (*.docx)</div>', unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Kéo thả file vào đây", type=["docx"], label_visibility="collapsed")
        
        # --- FIX: LƯU FILE VÀO SESSION STATE ---
        if uploaded_file is not None:
            st.session_state['file_bytes'] = uploaded_file.getvalue()
            st.session_state['file_name'] = uploaded_file.name
            st.success(f"✅ Đã tải lên: {uploaded_file.name}")

    # --- CỘT PHẢI ---
    with col_right:
        # BƯỚC 2: KIỂU TRỘN
        st.markdown('<div class="step-label"><div class="step-badge">2</div>Chọn kiểu trộn</div>', unsafe_allow_html=True)
        mode = st.radio(
            "Mode", ["auto", "mcq", "tf"],
            format_func=lambda x: {
                "auto": "🔄 Tự động (Phát hiện 3 phần)",
                "mcq": "📝 Trắc nghiệm (Toàn bộ A.B.C.D)",
                "tf": "✅ Đúng/Sai (Toàn bộ a)b)c)d))"
            }[x],
            label_visibility="collapsed", horizontal=False
        )
        st.write("")

        # BƯỚC 3: SỐ LƯỢNG
        st.markdown('<div class="step-label"><div class="step-badge">3</div>Số mã đề cần tạo</div>', unsafe_allow_html=True)
        c_num1, c_num2 = st.columns([1, 2])
        with c_num1:
            num_mix = st.number_input("Số lượng", 1, 50, 4, label_visibility="collapsed")
        with c_num2:
            st.markdown("""
            <div style="font-size:0.9rem; color:#666; padding-top:10px;">
                ● 1 mã → File Word<br>● Nhiều mã → File ZIP
            </div>
            """, unsafe_allow_html=True)

        st.markdown("---")
        
        # NÚT XỬ LÝ
        if st.button("🚀 Trộn đề & Tải xuống"):
            if 'file_bytes' in st.session_state:
                with st.spinner("Đang xử lý..."):
                    try:
                        all_keys = []
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
                            for i in range(num_mix):
                                v_name = f"{101 + i}"
                                # Gọi hàm trộn (Core cũ + Logic mới)
                                doc_bytes, key = shuffle_docx_and_get_key(
                                    st.session_state['file_bytes'], mode, v_name
                                )
                                zout.writestr(f"De_{v_name}.docx", doc_bytes)
                                all_keys.append(key)
                        
                        # Excel Key
                        df = pd.DataFrame(all_keys)
                        cols = list(df.columns)
                        if "Mã đề" in cols: cols.remove("Mã đề")
                        q_cols = sorted(cols, key=lambda s: int(re.search(r'(\d+)', s).group(1)) if re.search(r'(\d+)', s) else 0)
                        df = df.reindex(columns=["Mã đề"] + q_cols)
                        excel_buf = io.BytesIO()
                        with pd.ExcelWriter(excel_buf, engine='xlsxwriter') as writer:
                            df.to_excel(writer, index=False, sheet_name='DapAn')
                            
                        st.success("Thành công!")
                        d1, d2 = st.columns(2)
                        with d1:
                            st.download_button("📥 Tải Đề (ZIP)", zip_buffer.getvalue(), "De_Tron.zip", "application/zip", use_container_width=True)
                        with d2:
                            st.download_button("📊 Đáp án (Excel)", excel_buf.getvalue(), "Dap_An.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                            
                    except Exception as e:
                        st.error(f"Lỗi: {e}")
            else:
                st.warning("Vui lòng tải file ở Bước 1 trước.")

    # Footer
    st.markdown('<div style="text-align:center; color: #aaa; margin-top:20px; font-size:0.8rem;">© 2025 Phan Trường Duy - THPT Minh Đức</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
