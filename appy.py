import streamlit as st
import pandas as pd 
import io

from html import escape

def highlight_differences(a, b):
    a_words = a.split()
    b_words = b.split()
    matcher = difflib.SequenceMatcher(None, a_words, b_words)
    
    a_out = []
    b_out = []

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            a_out.extend([escape(word) for word in a_words[i1:i2]])
            b_out.extend([escape(word) for word in b_words[j1:j2]])
        elif tag in ("replace", "delete"):
            a_out.extend([f"<u>{escape(word)}</u>" for word in a_words[i1:i2]])
        if tag in ("replace", "insert"):
            b_out.extend([f"<u>{escape(word)}</u>" for word in b_words[j1:j2]])

    return " ".join(a_out), " ".join(b_out)

def render_html_table(results):
    table_html = """
    <style>
    table {
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 2rem;
        font-size: 15px;
    }
    th, td {
        border: 1px solid #ccc;
        padding: 0.75rem;
        vertical-align: top;
        text-align: left;
    }
    th {
        background-color: #f2f2f2;
    }
    tr.same {
        background-color: #d0f0c0;  /* new green */
    }
    tr.modified {
        background-color: #fff3cd;  /* yellow */
    }
    tr.added {
        background-color: #b3d9ff;  /* brighter blue */
    }
    tr.deleted {
        background-color: #ffcccc;  /* light red */
    }
    u {
        text-decoration: underline;
        font-weight: bold;
    }
    </style>
    """
    
    # Build table in one clean string (no line breaks inside cells)
    table_html += "<table><thead><tr><th>Status</th><th>Original</th><th>Revised</th></tr></thead><tbody>"
    
    for row in results:
        status_class = row['Status'].lower()
        table_html += (
            f"<tr class='{status_class}'>"
            f"<td><b>{row['Status']}</b></td>"
            f"<td>{row['Original']}</td>"
            f"<td>{row['Revised']}</td>"
            f"</tr>"
        )
    
    table_html += "</tbody></table>"
    return table_html



st.set_page_config(
    page_title="DOCX Change Comparison Table",
    layout="wide",
    initial_sidebar_state="auto"
)

st.title("📄 변경 대비표")
st.markdown("""
변경 대비표에 오신 것을 환영합니다. Word documents (.docx) -- **오리지널** 먼저, 그리고 **변경한** 버전 -- 두개를 밑에 업로드 하면 두 word documents의 차의/변경표가 만들어집니다. 
""")

st.header("📂 파일 업로드")

col1, col2 = st.columns(2)

with col1:
    file1 = st.file_uploader("기존 문서 업로드 (.docx)", type="docx", key="original")

with col2:
    file2 = st.file_uploader("개정 문서 업로드 (.docx)", type="docx", key="revised")

if file1 and file2:
    st.success("✅ 두 파일 모두 업로드 완료! 분석을 시작할 수 있습니다.")
else:
    st.warning("📎 두 개의 Word 파일을 모두 업로드해주세요.")

from docx import Document  # Make sure this import is at the top!

# ---- LOAD PARAGRAPHS ----
def extract_paragraphs(file):
    try:
        doc = Document(file)
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        return paragraphs
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return []

# ---- PROCESS FILES IF BOTH ARE UPLOADED ----
if file1 and file2:
    with st.spinner("Reading and extracting paragraphs..."):
        original_paragraphs = extract_paragraphs(file1)
        revised_paragraphs = extract_paragraphs(file2)

    st.write(f"📄 기존 문서에서 {len(original_paragraphs)}개 문단 추출됨")
    st.write(f"📝 개정 문서에서 {len(revised_paragraphs)}개 문단 추출됨")


import difflib  # Make sure this is imported at the top too

# ---- CLASSIFY DIFFERENCES ----
def classify_diff(old, new, threshold=0.9):
    if old == new:
        return "Same", old, new
    elif not new:
        return "Deleted", old, "<Deleted>"
    elif not old:
        return "Added", "<New>", new
    else:
        ratio = difflib.SequenceMatcher(None, old, new).ratio()
        if ratio >= threshold:
            return "Modified", old, new
        else:
            return "Modified", old, new

# ---- COMPARE PARAGRAPHS ----
def compare_documents(original_paras, revised_paras):
    sm = difflib.SequenceMatcher(None, original_paras, revised_paras)
    result = []

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for i, j in zip(range(i1, i2), range(j1, j2)):
                result.append({
                    "Status": "Same",
                    "Original": original_paras[i],
                    "Revised": revised_paras[j]
                })

        elif tag == "replace":
            # try to match pairs
            len1 = i2 - i1
            len2 = j2 - j1
            min_len = min(len1, len2)

            # line-by-line replace
            for k in range(min_len):
                orig_raw = original_paras[i1 + k]
                rev_raw = revised_paras[j1 + k]
                orig_diff, rev_diff = highlight_differences(orig_raw, rev_raw)

                result.append({
                    "Status": "Modified",
                    "Original": orig_diff,
                    "Revised": rev_diff
                })

            # anything left over in original = deleted
            for k in range(min_len, len1):
                result.append({
                    "Status": "Deleted",
                    "Original": original_paras[i1 + k],
                    "Revised": "<Deleted>"
                })

            # anything left over in revised = added
            for k in range(min_len, len2):
                result.append({
                    "Status": "Added",
                    "Original": "<New>",
                    "Revised": revised_paras[j1 + k]
                })

        elif tag == "delete":
            for i in range(i1, i2):
                result.append({
                    "Status": "Deleted",
                    "Original": original_paras[i],
                    "Revised": "<Deleted>"
                })

        elif tag == "insert":
            for j in range(j1, j2):
                result.append({
                    "Status": "Added",
                    "Original": "<New>",
                    "Revised": revised_paras[j]
                })

    return result

if file1 and file2:
    comparison_results = compare_documents(original_paragraphs, revised_paragraphs)
    df = pd.DataFrame(comparison_results)

    st.header("🔍 변경 사항 비교 결과")
    st.markdown("아래는 두 문서를 나란히 비교한 표입니다.")

    # Highlight rows based on Status
    def highlight_diff(row):
        if row.Status == "Same":
            return ['background-color: #d4edda']*3  # light green
        elif row.Status == "Deleted":
            return ['background-color: #ffcccc']*3  # light red
        elif row.Status == "Added":
            return ['background-color: #cce5ff']*3  # light blue
        elif row.Status == "Modified":
            return ['background-color: #ffe699']*3  # light orange
        return ['']*3

    styled_df = df.style.apply(highlight_diff, axis=1)
    st.subheader("📊 변경 대비표 (수정된 부분 밑줄 표시)")

    st.components.v1.html(render_html_table(comparison_results), height=800, scrolling=True)


    # ---- OPTIONAL: DOWNLOAD CSV ----
    st.markdown("### 📥 비교 결과 다운로드")
    csv_buffer = io.StringIO()
    df.to_csv(csv_buffer, index=False)
    st.download_button(
        label="CSV 파일로 다운로드",
        data=csv_buffer.getvalue(),
        file_name="변경_대비표.csv",
        mime="text/csv"
    )


