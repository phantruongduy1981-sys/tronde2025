import streamlit as st
import re
import random
import zipfile
import io
import pandas as pd
from xml.dom import minidom

# ==================== CẤU HÌNH TRANG & GIAO DIỆN ====================
st.set_page_config(
    page_title="Trộn Đề Word - THPT Minh Đức",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Màu chủ đạo theo hình ảnh bạn gửi (Xanh lá đậm)
DQ_COLOR = "#00897b" 
BG_COLOR = "#e0f2f1"

st.markdown(f"""
<style>
    .header-container {{
        background-color: {DQ_COLOR};
        padding: 2rem;
        border-radius: 10px;
        text-align: center;
        color: white;
        margin-bottom: 2rem;
    }}
    .header-container h1 {{
        font-family: 'Arial', sans-serif;
        text-transform: uppercase;
        font-weight: bold;
        margin: 0;
        font-size: 2.5rem;
        color: white;
    }}
    .header-container p {{
        margin-top: 10px;
        font-size: 1.2rem;
        opacity: 0.9;
    }}
    .stButton>button {{
        background-color: {DQ_COLOR};
        color: white;
        border-radius: 8px;
        height: 3rem;
        font-weight: bold;
        font-size: 16px;
        border: none;
        width: 100%;
    }}
    .stButton>button:hover {{
        background-color: #00695c;
        color: white;
    }}
    .upload-box {{
        border: 2px dashed {DQ_COLOR};
        border-radius: 10px;
        padding: 20px;
        text-align: center;
        background-color: #fafffa;
    }}
    .error-box {{
        background-color: #ffebee;
        color: #c62828;
        padding: 10px;
        border-radius: 5px;
        border: 1px solid #ef9a9a;
        margin-top: 10px;
    }}
    .footer {{
        text-align: center;
        margin-top: 50px;
        color: #666;
        font-size: 0.9rem;
        border-top: 1px solid #eee;
        padding-top: 20px;
    }}
</style>
""", unsafe_allow_html=True)

# ==================== CORE LOGIC: XML PARSING ====================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_text(block):
    """Lấy text thuần từ block XML"""
    texts = []
    t_nodes = block.getElementsByTagNameNS(W_NS, "t")
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue:
            texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def check_structure(blocks):
    """Kiểm tra cấu trúc file và trả về danh sách lỗi"""
    errors = []
    full_text = "\n".join([get_text(b) for b in blocks])
    
    # 1. Kiểm tra có từ khóa 'Câu 1' không
    if not re.search(r'Câu\s*1[\.:]', full_text, re.IGNORECASE):
        errors.append("❌ Không tìm thấy 'Câu 1'. Vui lòng kiểm tra lại định dạng bắt đầu câu hỏi.")
    
    # 2. Kiểm tra các phần
    if "PHẦN 1" not in full_text and "PHẦN 2" not in full_text:
        errors.append("⚠️ Cảnh báo: Không tìm thấy phân chia 'PHẦN 1', 'PHẦN 2'. App sẽ hiểu là trộn toàn bộ dạng trắc nghiệm.")

    # 3. Kiểm tra đáp án (Sơ bộ)
    # Logic: Kiểm tra xem có A. B. C. D. không
    if not re.search(r'A\.', full_text) or not re.search(r'B\.', full_text):
         errors.append("⚠️ Cảnh báo: File có thể thiếu các phương án A. B. C. D. chuẩn.")

    return errors

def is_emphasized(paragraph):
    """
    Kiểm tra xem đoạn văn (phương án) có chứa định dạng đáp án đúng không.
    Đáp án đúng: Màu Đỏ (red/FF0000) hoặc Gạch chân (underline).
    """
    runs = paragraph.getElementsByTagNameNS(W_NS, "r")
    for run in runs:
        rPr = run.getElementsByTagNameNS(W_NS, "rPr")
        if not rPr: continue
        rPr = rPr[0]
        
        # Check Color
        colors = rPr.getElementsByTagNameNS(W_NS, "color")
        for c in colors:
            val = c.getAttributeNS(W_NS, "val")
            if val and (val.upper() in ['FF0000', 'RED']):
                return True
                
        # Check Underline
        u_tags = rPr.getElementsByTagNameNS(W_NS, "u")
        if u_tags:
            # Nếu có thẻ u mà val khác none thì là gạch chân
            val = u_tags[0].getAttributeNS(W_NS, "val")
            if val != 'none':
                return True
    return False

def clean_run_formatting(paragraph):
    """Xóa màu đỏ và gạch chân sau khi đã ghi nhận đáp án"""
    runs = paragraph.getElementsByTagNameNS(W_NS, "r")
    for run in runs:
        rPr_list = run.getElementsByTagNameNS(W_NS, "rPr")
        if not rPr_list: continue
        rPr = rPr_list[0]
        
        # Xóa màu
        colors = rPr.getElementsByTagNameNS(W_NS, "color")
        for c in colors:
            rPr.removeChild(c)
            
        # Xóa gạch chân
        u_tags = rPr.getElementsByTagNameNS(W_NS, "u")
        for u in u_tags:
            rPr.removeChild(u)

def update_label(paragraph, new_label):
    """Cập nhật nhãn A. B. C. D. hoặc Câu X."""
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    
    # Tìm node text đầu tiên chứa nội dung
    found = False
    for t in t_nodes:
        if not t.firstChild: continue
        txt = t.firstChild.nodeValue
        
        # Regex bắt: A. hoặc Câu 1.
        # Xử lý cho Phương án A. B. C. D.
        if re.match(r'^\s*[A-D][\.:\)]', txt, re.IGNORECASE):
            # Giữ lại phần text phía sau dấu chấm
            m = re.match(r'^(\s*[A-D][\.:\)])(.*)', txt, re.IGNORECASE)
            if m:
                remain = m.group(2)
                t.firstChild.nodeValue = f"{new_label}{remain}"
                found = True
                break
        
        # Xử lý cho Câu X.
        elif re.match(r'^\s*Câu\s*\d+', txt, re.IGNORECASE):
             m = re.match(r'^(\s*Câu\s*\d+[\.:]?)(.*)', txt, re.IGNORECASE)
             if m:
                remain = m.group(2)
                t.firstChild.nodeValue = f"{new_label}{remain}"
                found = True
                break
    
    # Nếu text bị chia nhỏ (VD: "Câu" ở node 1, "1" ở node 2), xử lý phức tạp hơn
    # Ở đây dùng cách đơn giản: Gán label vào node đầu, các node sau nếu chỉ chứa số thứ tự cũ thì xóa đi (Logic đơn giản hóa)

def process_questions(questions_blocks, mode="MCQ"):
    """
    Trộn câu hỏi và phương án.
    Trả về: (blocks_đã_trộn, map_đáp_án)
    """
    # 1. Trộn thứ tự câu hỏi
    indices = list(range(len(questions_blocks)))
    random.shuffle(indices)
    
    shuffled_blocks = []
    answer_key = {} # { "Câu 1": "A", "Câu 2": "C" ...}
    
    # Định nghĩa label phương án
    labels_mcq = ["A.", "B.", "C.", "D."]
    labels_tf = ["a)", "b)", "c)", "d)"]
    
    current_q_num = 1
    
    for original_idx in indices:
        q_group = questions_blocks[original_idx] # Một nhóm block gồm [Câu dẫn, A, B, C, D]
        
        # Tách câu dẫn và phương án
        intro_blocks = []
        option_blocks = [] # List of tuples: (block, is_correct)
        
        # Phân loại block
        for block in q_group:
            txt = get_text(block)
            is_opt = False
            if mode == "MCQ" and re.match(r'^\s*[A-D][\.:]', txt): is_opt = True
            elif mode == "TF" and re.match(r'^\s*[a-d][\)]', txt): is_opt = True
            
            if is_opt:
                # Kiểm tra xem đây có phải đáp án đúng không
                is_right = is_emphasized(block)
                # Xóa định dạng đỏ/gạch chân để đề thi sạch
                clean_run_formatting(block) 
                option_blocks.append({'block': block, 'is_correct': is_right})
            else:
                intro_blocks.append(block)
        
        # Xử lý trộn phương án
        final_options = []
        correct_char = ""
        
        if len(option_blocks) > 0:
            if mode == "MCQ":
                random.shuffle(option_blocks)
                # Gán lại nhãn A, B, C, D
                for i, opt in enumerate(option_blocks):
                    new_lbl = labels_mcq[i] if i < 4 else "*"
                    update_label(opt['block'], new_lbl)
                    if opt['is_correct']:
                        correct_char = new_lbl.replace(".", "")
                final_options = [o['block'] for o in option_blocks]
                
            elif mode == "TF":
                # Đúng sai thường trộn a,b,c giữ d, hoặc trộn cả. Ở đây ta trộn cả
                # Nhưng logic Đúng/Sai phức tạp hơn vì mỗi ý a,b,c,d đều có Đ/S.
                # Ở đây giả định user muốn trộn thứ tự xuất hiện các ý
                random.shuffle(option_blocks)
                for i, opt in enumerate(option_blocks):
                    new_lbl = labels_tf[i] if i < 4 else "*"
                    update_label(opt['block'], new_lbl)
                    # Với đúng sai, đáp án ko phải là A/B/C/D duy nhất nên ta bỏ qua ghi key kiểu này
                    # Hoặc ghi nhận những câu nào là Đúng
                final_options = [o['block'] for o in option_blocks]

        # Cập nhật số thứ tự câu hỏi (Câu 1, Câu 2...)
        if intro_blocks:
            update_label(intro_blocks[0], f"Câu {current_q_num}.")
        
        # Lưu vào danh sách kết quả
        shuffled_blocks.extend(intro_blocks)
        shuffled_blocks.extend(final_options)
        
        # Lưu đáp án
        if mode == "MCQ" and correct_char:
            answer_key[current_q_num] = correct_char
        elif mode == "MCQ":
            answer_key[current_q_num] = "X" # Không tìm thấy đáp án tô đỏ
            
        current_q_num += 1
            
    return shuffled_blocks, answer_key

def parse_docx_and_shuffle(file_bytes, num_versions, shuffle_mode_ui):
    """Hàm chính xử lý file"""
    input_buffer = io.BytesIO(file_bytes)
    
    try:
        zin = zipfile.ZipFile(input_buffer, 'r')
        doc_xml = zin.read("word/document.xml").decode('utf-8')
    except:
        return None, None, ["File lỗi hoặc không phải định dạng .docx chuẩn"]

    dom = minidom.parseString(doc_xml)
    body = dom.getElementsByTagNameNS(W_NS, "body")[0]
    
    # Tách các block (p và tbl)
    all_blocks = []
    for child in list(body.childNodes):
        if child.nodeType == child.ELEMENT_NODE and child.localName in ["p", "tbl"]:
            all_blocks.append(child)
            body.removeChild(child) # Xóa khỏi cây DOM để tí gắn lại sau
    
    # Kiểm tra lỗi cấu trúc
    errors = check_structure(all_blocks)
    if errors:
        # Nếu có lỗi nghiêm trọng thì trả về luôn (ở đây ta chỉ warning và vẫn chạy)
        pass 

    # --- TÁCH CÁC PHẦN (PART 1, 2, 3) ---
    # Logic đơn giản: Tìm text "PHẦN 1", "PHẦN 2"... để cắt list blocks
    parts = []
    current_part = []
    
    # Regex tìm phần
    part_pattern = re.compile(r'^\s*PHẦN\s*(\d+)', re.IGNORECASE)
    
    for block in all_blocks:
        txt = get_text(block)
        if part_pattern.match(txt):
            if current_part: parts.append(current_part)
            current_part = [block]
        else:
            current_part.append(block)
    if current_part: parts.append(current_part)
    
    # Nếu không tìm thấy chữ PHẦN nào, coi như cả file là 1 phần
    if not re.search(r'PHẦN\s*\d', "\n".join([get_text(b) for b in all_blocks])):
        parts = [all_blocks]

    # --- BẮT ĐẦU TẠO CÁC MÃ ĐỀ ---
    output_zips = io.BytesIO()
    excel_data = [] # List of dicts for DataFrame
    
    with zipfile.ZipFile(output_zips, 'w', zipfile.ZIP_DEFLATED) as zout:
        
        for ver_i in range(num_versions):
            exam_code = f"10{ver_i+1}" # Mã đề 101, 102...
            
            # Copy DOM gốc để tạo file mới
            new_dom = minidom.parseString(doc_xml)
            new_body = new_dom.getElementsByTagNameNS(W_NS, "body")[0]
            # Xóa sạch con cũ
            while new_body.firstChild:
                new_body.removeChild(new_body.firstChild)

            current_exam_key = {"Mã đề": exam_code}
            global_q_idx = 1
            
            # Duyệt qua từng phần để trộn
            for p_idx, part_blocks in enumerate(parts):
                # Tách riêng các câu hỏi trong phần này
                # Logic: Câu hỏi bắt đầu bằng "Câu X"
                intro_part = []
                questions_list = []
                current_q = []
                
                is_in_question = False
                
                for b in part_blocks:
                    txt = get_text(b)
                    if re.match(r'^\s*Câu\s*\d+', txt, re.IGNORECASE):
                        if current_q: questions_list.append(current_q)
                        current_q = [b]
                        is_in_question = True
                    elif re.match(r'^\s*PHẦN', txt, re.IGNORECASE):
                        if current_q: questions_list.append(current_q)
                        current_q = []
                        intro_part.append(b)
                        is_in_question = False
                    else:
                        if is_in_question:
                            current_q.append(b)
                        else:
                            intro_part.append(b)
                if current_q: questions_list.append(current_q)
                
                # Xác định chế độ trộn cho phần này
                # Mặc định: Phần 1 là MCQ, Phần 2 là TF (nhưng ở đây làm đơn giản theo UI user chọn)
                # Nếu User chọn Auto:
                mode = "MCQ"
                part_text = get_text(intro_part[0]) if intro_part else ""
                
                if shuffle_mode_ui == "auto":
                    if "PHẦN 2" in part_text.upper(): mode = "TF"
                    elif "PHẦN 3" in part_text.upper(): mode = "NO_SHUFFLE_OPT" # Tự luận/Điền khuyết
                    else: mode = "MCQ"
                elif shuffle_mode_ui == "mcq":
                    mode = "MCQ"
                else: # true/false
                    mode = "TF"

                # Thực hiện trộn
                if mode == "NO_SHUFFLE_OPT":
                    # Chỉ trộn thứ tự câu, không trộn đáp án
                    random.shuffle(questions_list)
                    processed_blocks = intro_part
                    for q in questions_list:
                        # Cập nhật số câu
                        update_label(q[0], f"Câu {global_q_idx}.")
                        global_q_idx += 1
                        processed_blocks.extend(q)
                else:
                    # Trộn cả câu và đáp án
                    # Cần clone các node để không ảnh hưởng lần lặp sau (Python minidom node chỉ có 1 parent)
                    # Lưu ý: minidom cloneNode(True) deep copy
                    q_clones = [[node.cloneNode(True) for node in q] for q in questions_list]
                    intro_clones = [node.cloneNode(True) for node in intro_part]
                    
                    mixed_blocks, key_map = process_questions(q_clones, mode=mode)
                    
                    # Update lại số câu global cho đúng (vì process_questions reset về 1)
                    # Fix lại số câu trong block
                    real_mixed = []
                    # Append intro
                    real_mixed.extend(intro_clones)
                    
                    # Do process_questions đã gán "Câu 1", "Câu 2"... cục bộ
                    # Ta cần sửa lại theo global_q_idx
                    # Nhưng để đơn giản, ta chấp nhận process_questions trả về list block
                    # Ta duyệt lại để sửa số câu? Hơi nặng.
                    # Cách tốt nhất: process_questions nhận start_index
                    # (Code trên tôi viết process_questions reset về 1. Ở đây ta chỉnh lại Text Node thủ công chút nếu cần)
                    # Sửa nhanh: cập nhật key map vào excel global
                    for loc_idx, ans in key_map.items():
                        current_exam_key[f"{global_q_idx + loc_idx - 1}"] = ans
                    
                    # Sửa lại label câu trong XML (nếu Phần 2 bắt đầu từ câu 13 chẳng hạn)
                    q_count_in_part = len(key_map) if key_map else len(questions_list)
                    
                    # Đoạn này xử lý lại label "Câu X" cho khớp global index
                    # Tìm tất cả các node "Câu X" trong mixed_blocks và + offset
                    count_q = 0
                    for blk in mixed_blocks:
                        txt = get_text(blk)
                        if re.match(r'^\s*Câu\s*\d+', txt):
                            count_q += 1
                            update_label(blk, f"Câu {global_q_idx + count_q - 1}.")
                            
                    global_q_idx += q_count_in_part
                    real_mixed.extend(mixed_blocks)
                    
                    # Gắn vào DOM mới
                    for b in real_mixed:
                        new_body.appendChild(b)

            # --- GHI FILE DOCX MỚI ---
            new_xml = new_dom.toxml()
            docx_out = io.BytesIO()
            with zipfile.ZipFile(docx_out, 'w', zipfile.ZIP_DEFLATED) as zdoc:
                for item in zin.infolist():
                    if item.filename == "word/document.xml":
                        zdoc.writestr(item, new_xml.encode('utf-8'))
                    else:
                        zdoc.writestr(item, zin.read(item.filename))
            
            zout.writestr(f"De_Thi_{exam_code}.docx", docx_out.getvalue())
            excel_data.append(current_exam_key)

    # --- TẠO FILE EXCEL ĐÁP ÁN ---
    df = pd.DataFrame(excel_data)
    # Sắp xếp cột cho đẹp (Mã đề, 1, 2, 3...)
    cols = ["Mã đề"] + sorted([c for c in df.columns if c != "Mã đề"], key=lambda x: int(x))
    df = df[cols]
    
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='DapAn')
        # Format đẹp
        workbook = writer.book
        worksheet = writer.sheets['DapAn']
        header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#D7E4BC', 'border': 1})
        center_fmt = workbook.add_format({'align': 'center', 'border': 1})
        
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
            worksheet.set_column(col_num, col_num, 5, center_fmt)
        worksheet.set_column(0, 0, 10, center_fmt) # Cột mã đề rộng hơn

    return output_zips.getvalue(), excel_buffer.getvalue(), errors

# ==================== MAIN UI ====================

def main():
    # Header Section
    st.markdown("""
        <div class="header-container">
            <h1>TRƯỜNG THPT MINH ĐỨC</h1>
            <p>APP TRỘN ĐỀ 2025</p>
        </div>
    """, unsafe_allow_html=True)

    col_left, col_right = st.columns([1, 1.2], gap="large")

    with col_left:
        with st.expander("📄 Hướng dẫn & Cấu trúc (Bấm để xem)", expanded=True):
            st.markdown("""
            **Cấu trúc file Word chuẩn:**
            * **PHẦN 1:** Trắc nghiệm (A. B. C. D.)
            * **PHẦN 2:** Đúng/Sai (a) b) c) d))
            * **Lưu ý quan trọng:**
                * Câu hỏi bắt đầu bằng `Câu 1.`, `Câu 2.`...
                * **Đáp án đúng:** Phải được <span style='color:red'><b>TÔ ĐỎ</b></span> hoặc <u><b>GẠCH CHÂN</b></u> trong file gốc để tạo file Excel.
            """, unsafe_allow_html=True)
            # Nút tải file mẫu giả lập
            st.button("📥 Tải File Mẫu", key="btn_sample")

        st.markdown("### 1. Chọn file đề Word (*.docx)")
        uploaded_file = st.file_uploader("", type=["docx"], label_visibility="collapsed")
        
        if uploaded_file:
            st.success(f"✅ Đã tải lên: {uploaded_file.name}")
            # Kiểm tra sơ bộ
            # (Phần này xử lý trong luồng chính để tối ưu hiệu năng)

        st.markdown("""
        <div class="upload-box">
            Drag and drop file here<br>
            <small>Limit 200MB per file • DOCX</small>
        </div>
        """, unsafe_allow_html=True)

    with col_right:
        st.markdown("### 2. Chọn kiểu trộn")
        shuffle_opt = st.radio(
            "",
            ["auto", "mcq", "tf"],
            format_func=lambda x: {
                "auto": "⚙️ Tự động (Theo PHẦN 1, 2, 3)",
                "mcq": "📝 Trắc nghiệm (A, B, C, D)",
                "tf": "✅ Đúng/Sai"
            }[x],
            key="shuffle_mode"
        )
        
        st.markdown("### 3. Số mã đề cần tạo")
        num_exams = st.number_input("", min_value=1, max_value=50, value=4, step=1)
        st.caption("ℹ️ 1 mã -> File Word. Nhiều mã -> File ZIP")

        st.markdown("---")
        
        # Nút hành động chính
        if st.button("🚀 Trộn đề & Tải xuống"):
            if not uploaded_file:
                st.error("Vui lòng tải file đề lên trước!")
            else:
                with st.spinner("Đang xử lý..."):
                    try:
                        file_bytes = uploaded_file.read()
                        zip_data, excel_data, errors = parse_docx_and_shuffle(
                            file_bytes, 
                            num_exams, 
                            shuffle_opt
                        )
                        
                        if errors:
                            for err in errors:
                                st.markdown(f"<div class='error-box'>{err}</div>", unsafe_allow_html=True)
                        
                        if zip_data:
                            # 1. Nút tải Đề (ZIP)
                            st.download_button(
                                label="📦 Tải Bộ Đề (Word)",
                                data=zip_data,
                                file_name=f"Bo_De_{uploaded_file.name}.zip",
                                mime="application/zip",
                                type="primary"
                            )
                            
                            # 2. Nút tải Đáp án (Excel)
                            st.download_button(
                                label="📊 Tải Đáp Án (Excel)",
                                data=excel_data,
                                file_name=f"Dap_An_{uploaded_file.name}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            st.balloons()
                            
                    except Exception as e:
                        st.error(f"Lỗi xử lý: {str(e)}")
                        # In chi tiết lỗi để debug
                        import traceback
                        st.text(traceback.format_exc())

    # Footer
    st.markdown("""
        <div class="footer">
            © 2025 Phan Trường Duy - THPT Minh Đức<br>
            Hệ thống quản lý trộn đề thi trắc nghiệm
        </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
