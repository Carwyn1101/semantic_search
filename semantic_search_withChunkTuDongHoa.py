import streamlit as st
import os
import docx
import fitz  
import pandas as pd  

import win32com.client
import tempfile

from underthesea import sent_tokenize
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity


from transformers import AutoTokenizer
tokenizer = AutoTokenizer.from_pretrained("VoVanPhuc/sup-SimCSE-VietNamese-phobert-base")

import re

import faiss


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INDEX_DIR = os.path.join(BASE_DIR, "faiss_index")
os.makedirs(INDEX_DIR, exist_ok=True)

INDEX_PATH = os.path.join(INDEX_DIR, "index.faiss")
META_PATH = os.path.join(INDEX_DIR, "chunks.pkl")


def highlight_keywords(text, query):
    keywords = re.findall(r"\w+", query.lower())
    for kw in set(keywords):
        text = re.sub(fr"(?i)\b({re.escape(kw)})\b", r"<mark>\1</mark>", text)
    return text


def build_and_save_faiss_index(chunks, model, index_path, meta_path):
    texts = [c['text'] for c in chunks]
    doc_embs = model.encode(texts, show_progress_bar=True)
    faiss.normalize_L2(doc_embs)

    index = faiss.IndexFlatIP(doc_embs.shape[1])
    index.add(doc_embs.astype('float32'))

    faiss.write_index(index, index_path)
    with open(meta_path, "wb") as f:
        import pickle
        pickle.dump(chunks, f)
        

def load_faiss_index(index_path, meta_path):
    index = faiss.read_index(index_path)
    with open(meta_path, "rb") as f:
        import pickle
        chunks = pickle.load(f)
    return index, chunks



def preprocess_text(text):
    # 1. Xử lý ký tự điều khiển thường gặp
    text = text.replace("\x0c", " ").replace("\xa0", " ")

    # 2. Loại bỏ dòng trắng và gom các dòng ngắt câu sai cách (OCR)
    lines = text.splitlines()
    clean_lines = [line.strip() for line in lines if line.strip()]
    merged_lines = []
    buffer = ""
    for line in clean_lines:
        if buffer:
            buffer += " " + line
        else:
            buffer = line

        if re.search(r"[.!?]$", line.strip()):
            merged_lines.append(buffer.strip())
            buffer = ""
    if buffer:
        merged_lines.append(buffer.strip())
    text = "\n\n".join(merged_lines)

    # 3. Chuẩn hóa khoảng trắng, ký tự đặc biệt
    text = text.lower()
    text = re.sub(r"[^\w\s.,:%–\-–—]", "", text)
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\n+", "\n", text)

    return text.strip()

def convert_doc_to_docx(doc_path):
    if doc_path.endswith(".doc"):
        doc_path = os.path.abspath(doc_path)
        doc_path = os.path.normpath(doc_path)

        if not os.path.exists(doc_path):
            raise FileNotFoundError(f"Không tìm thấy file: {doc_path}")

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        # Tạo thư mục tạm để chứa .docx
        temp_dir = tempfile.mkdtemp()
        temp_docx_path = os.path.join(temp_dir, os.path.basename(doc_path) + "x")  # file.docx

        try:
            doc = word.Documents.Open(doc_path)
            doc.SaveAs(temp_docx_path, FileFormat=16)
            doc.Close()
            return temp_docx_path
        finally:
            word.Quit()
    return doc_path


# ---- Đọc nội dung .docx ----
def read_docx(file_path):
    doc = docx.Document(file_path)
    return "\n".join([p.text for p in doc.paragraphs])

# ---- Đọc nội dung .pdf ----
def read_pdf(file_path):
    doc = fitz.open(file_path)
    text = ""
    for page in doc:
        text += page.get_text()
    return text

# ---- Đọc nội dung .xlsx và .xls ----
def read_excel(file_path):
    if file_path.endswith(".xlsx"):
        df = pd.read_excel(file_path, engine="openpyxl")
    elif file_path.endswith(".xls"):
        df = pd.read_excel(file_path, engine="xlrd")
    else:
        return ""
    text = df.astype(str).apply(lambda row: " ".join(row), axis=1).str.cat(sep="\n")
    return text


# def chunk_by_token_limit(text, tokenizer, max_tokens):
#     sentences = sent_tokenize(text)
#     chunks = []
#     current_chunk = ""
#     current_tokens = 0

#     for sent in sentences:
#         sent_tokens = len(tokenizer.encode(sent, add_special_tokens=False))
#         if current_tokens + sent_tokens <= max_tokens:
#             current_chunk += " " + sent
#             current_tokens += sent_tokens
#         else:
#             chunks.append(current_chunk.strip())
#             current_chunk = sent
#             current_tokens = sent_tokens

#     if current_chunk:
#         chunks.append(current_chunk.strip())

#     return chunks

# def auto_chunk(text, tokenizer, max_tokens):
#     return chunk_by_token_limit(text, tokenizer, max_tokens)


# ---- Chunk tự động theo nội dung ----
def auto_chunk(text, max_len=512):
    paragraphs = text.split("\n\n")
    chunks = []
    for para in paragraphs:
        para = para.strip()
        if not para:
            continue

        num_words = len(para.split())
        num_sentences = len(sent_tokenize(para))

        if num_words < 150 or len(para) < max_len:
            chunks.append(para)  # Đoạn ngắn, giữ nguyên
        elif num_sentences < 4:
            chunks += chunk_by_sliding_window(para, window_size=max_len, stride=max_len // 2)  # Ít câu, dài → sliding
        else:
            chunks += chunk_by_sentence(para, max_len)  # Nhiều câu → chia theo câu
    return chunks
    

def chunk_by_sentence(text, max_len=512):
    sentences = sent_tokenize(text)
    chunks = []
    current = ""
    for s in sentences:
        if len(current) + len(s) < max_len:
            current += " " + s
        else:
            chunks.append(current.strip())
            current = s
    if current:
        chunks.append(current.strip())
    return chunks

def chunk_by_sliding_window(text, window_size, stride):
    words = text.split()
    chunks = []
    for i in range(0, len(words), stride):
        chunk = " ".join(words[i:i + window_size])
        if len(chunk.strip()) > 0:
            chunks.append(chunk)
        if i + window_size >= len(words):
            break
    return chunks


# ---- Load & chunk tất cả văn bản ----
@st.cache_data
# def load_chunked_documents(folder, _tokenizer, max_tokens):
#     chunks = []
#     for filename in os.listdir(folder):
#         if filename.endswith((".docx", ".pdf", ".xlsx", ".doc", ".xls")):
#             file_path = os.path.join(folder, filename)

#             if filename.endswith(".doc"):
#                 file_path = convert_doc_to_docx(file_path)

#             if filename.endswith(".docx"):
#                 # raw_text = read_docx(file_path)
#                 raw_text = preprocess_text(read_docx(file_path))
#             elif filename.endswith(".pdf"):
#                 # raw_text = read_pdf(file_path)
#                 raw_text = preprocess_text(read_pdf(file_path))
#             elif filename.endswith((".xlsx", ".xls")):
#                 # raw_text = read_excel(file_path)
#                 raw_text = preprocess_text(read_excel(file_path))
#             else:
#                 continue
#             doc_chunks = auto_chunk(raw_text, _tokenizer, max_tokens)
#             for i, chunk in enumerate(doc_chunks):
#                 chunks.append({
#                     "filename": filename,
#                     "chunk_id": i,
#                     "text": chunk,
#                     "fulltext": raw_text
#                 })
#     return chunks

def load_chunked_documents(folder, max_len=512):
    chunks = []
    for filename in os.listdir(folder):
        if filename.endswith((".docx", ".pdf", ".xlsx", ".doc", ".xls")):
            file_path = os.path.join(folder, filename)

            if filename.endswith(".doc"):
                file_path = convert_doc_to_docx(file_path)

            if filename.endswith(".docx"):
                raw_text = read_docx(file_path)
                # raw_text = preprocess_text(read_docx(file_path))
            elif filename.endswith(".pdf"):
                raw_text = read_pdf(file_path)
                # raw_text = preprocess_text(read_pdf(file_path))
            elif filename.endswith((".xlsx", ".xls")):
                raw_text = read_excel(file_path)
                # raw_text = preprocess_text(read_excel(file_path))
            else:
                continue
            doc_chunks = auto_chunk(raw_text, max_len)
            for i, chunk in enumerate(doc_chunks):
                chunks.append({
                    "filename": filename,
                    "chunk_id": i,
                    "text": chunk,
                    "fulltext": raw_text
                })
    return chunks

# ---- Tìm kiếm ngữ nghĩa ----
# def search_semantic_chunks(query, chunks, model, top_k=5):
#     query_emb = model.encode([query])
#     texts = [c['text'] for c in chunks]
#     doc_embs = model.encode(texts)
#     sims = cosine_similarity(query_emb, doc_embs).flatten()
    
#     # Gắn điểm similarity vào từng chunk
#     for i, sim in enumerate(sims):
#         chunks[i]['score'] = sim

#     # Nhóm theo filename: chỉ giữ chunk có score cao nhất trong mỗi file
#     best_chunks = {}
#     for c in chunks:
#         fname = c['filename']
#         if fname not in best_chunks or c['score'] > best_chunks[fname]['score']:
#             best_chunks[fname] = c

#     # Sắp xếp theo điểm cao nhất và lấy top_k file duy nhất
#     sorted_chunks = sorted(best_chunks.values(), key=lambda x: x['score'], reverse=True)[:top_k]

#     # Định dạng lại kết quả đầu ra đúng theo format ban đầu
#     results = []
#     for c in sorted_chunks:
#         results.append({
#             "filename": c["filename"],
#             "chunk_id": c["chunk_id"],
#             "score": c["score"],
#             "excerpt": c["text"],
#             "fulltext": c["fulltext"]
#         })
#     return results


def search_semantic_chunks(query, chunks, model, top_k=5):
    query_emb = model.encode([query])
    texts = [c['text'] for c in chunks]
    doc_embs = model.encode(texts)

    # FAISS expects float32
    index = faiss.IndexFlatIP(doc_embs.shape[1])  # Inner Product for cosine sim if vectors are normalized
    faiss.normalize_L2(doc_embs)
    faiss.normalize_L2(query_emb)
    index.add(doc_embs.astype("float32"))

    scores, indices = index.search(query_emb.astype("float32"), len(chunks))

    # Gắn điểm similarity vào từng chunk
    for i, idx in enumerate(indices[0]):
        chunks[idx]['score'] = float(scores[0][i])

    # Nhóm theo filename: chỉ giữ chunk có score cao nhất trong mỗi file
    best_chunks = {}
    for c in chunks:
        fname = c['filename']
        if fname not in best_chunks or c['score'] > best_chunks[fname]['score']:
            best_chunks[fname] = c

    # Lấy top_k file có điểm cao nhất
    sorted_chunks = sorted(best_chunks.values(), key=lambda x: x['score'], reverse=True)[:top_k]

    results = []
    for c in sorted_chunks:
        results.append({
            "filename": c["filename"],
            "chunk_id": c["chunk_id"],
            "score": c["score"],
            "excerpt": c["text"],
            "fulltext": c["fulltext"]
        })
    return results


def search_faiss_index(query, model, index, chunks, top_k=5):
    query_emb = model.encode([query])
    faiss.normalize_L2(query_emb)

    scores, indices = index.search(query_emb.astype("float32"), len(chunks))

    # Gắn điểm similarity vào từng chunk
    for i, idx in enumerate(indices[0]):
        chunks[idx]['score'] = float(scores[0][i])

    # Lọc top chunk mỗi file
    best_chunks = {}
    for c in chunks:
        fname = c['filename']
        if fname not in best_chunks or c['score'] > best_chunks[fname]['score']:
            best_chunks[fname] = c

    sorted_chunks = sorted(best_chunks.values(), key=lambda x: x['score'], reverse=True)[:top_k]

    results = []
    for c in sorted_chunks:
        results.append({
            "filename": c["filename"],
            "chunk_id": c["chunk_id"],
            "score": c["score"],
            "excerpt": c["text"],
            "fulltext": c["fulltext"]
        })
    return results


# ==== Giao diện Streamlit ====
st.set_page_config(page_title="Tìm kiếm ngữ nghĩa", layout="wide")
st.title("Tìm kiếm văn bản theo ngữ nghĩa (Chunk-based + PhoBERT)")
st.write("Nhập mô tả nội dung hoặc ý định của bạn, hệ thống sẽ tìm các đoạn văn bản có nội dung liên quan.")

query = st.text_input("Nhập nội dung tìm kiếm", placeholder="ví dụ: phần mềm quản lý thanh tra, xử lý rác thải,...")
submit = st.button("Tìm kiếm")

if submit and query:
    with st.spinner("Đang tải mô hình và tìm kiếm..."):
        model = SentenceTransformer("VoVanPhuc/sup-SimCSE-VietNamese-phobert-base")


        DATA_DIR = "File_Mau"
        # chunks = load_chunked_documents(DATA_DIR, max_len=512)

        if not os.path.exists(INDEX_PATH) or not os.path.exists(META_PATH):
            chunks = load_chunked_documents(DATA_DIR, max_len=512)
            build_and_save_faiss_index(chunks, model, INDEX_PATH, META_PATH)
            
        # Load index và metadata
        index, chunks = load_faiss_index(INDEX_PATH, META_PATH)

        
        # tokenizer = AutoTokenizer.from_pretrained("VoVanPhuc/sup-SimCSE-VietNamese-phobert-base")
        # max_len = tokenizer.model_max_length
        # chunks = load_chunked_documents(DATA_DIR, _tokenizer=tokenizer, max_tokens=max_len)

        # results = search_semantic_chunks(query, chunks, model)

        # Tìm kiếm từ index đã build
        results = search_faiss_index(query, model, index, chunks)

        if results:
            st.success(f"Tìm thấy {len(results)} đoạn văn bản phù hợp:")
            for res in results:
                st.markdown(f"### {res['filename']} — Chunk {res['chunk_id']} — Score: `{res['score']:.4f}`")
                # st.markdown(f"**Trích đoạn:**\n\n{res['excerpt']}...")

                excerpt = res['excerpt']
                short_excerpt = excerpt[:500] + "..." if len(excerpt) > 500 else excerpt
                highlighted_excerpt = highlight_keywords(short_excerpt, query)
                # Hiển thị đoạn rút gọn đã bôi sáng
                st.markdown(f"**Trích đoạn:**<br>{highlighted_excerpt}", unsafe_allow_html=True)


                # with st.expander("Xem toàn bộ nội dung gốc của văn bản"):
                #     # st.text(res["fulltext"])
                #     st.text_area("Nội dung đầy đủ", res["fulltext"], height=300)
                # # st.markdown("---")

                
                # Nút tải file gốc nếu có
                filepath = os.path.join(DATA_DIR, res['filename'])
                if os.path.exists(filepath):
                    with open(filepath, "rb") as f:
                        st.download_button("📎 Tải file gốc", f, file_name=res['filename'])


        else:
            st.warning("Không tìm thấy văn bản phù hợp.")
