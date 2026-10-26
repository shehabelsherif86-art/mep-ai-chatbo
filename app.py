import streamlit as st
from google import genai
from google.genai import types
import tempfile
import os
import time
from io import BytesIO
from docx import Document
from docx.oxml import OxmlElement

# ---------------------------------------------------------
# الإعدادات العامة
# ---------------------------------------------------------
# اسم الموديل (لو جوجل قفلته مستقبلاً غيّره من هنا فقط)
MODEL_NAME = "gemini-3.8-flash"

st.set_page_config(page_title="المساعد الهندسي الذكي للمواصفات", layout="wide")

# ضبط اتجاه النص (عربي + إنجليزي) وشكل الجداول
st.markdown("""
<style>
.stApp { direction: rtl; }

[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
[data-testid="stMarkdownContainer"] h1,
[data-testid="stMarkdownContainer"] h2,
[data-testid="stMarkdownContainer"] h3,
[data-testid="stMarkdownContainer"] h4,
[data-testid="stMarkdownContainer"] td,
[data-testid="stMarkdownContainer"] th {
    unicode-bidi: plaintext;
    text-align: start;
}

[data-testid="stMarkdownContainer"] ul,
[data-testid="stMarkdownContainer"] ol { direction: rtl; }
[data-testid="stMarkdownContainer"] table { direction: rtl; width: 100%; }

[data-testid="stMarkdownContainer"] code {
    direction: ltr;
    unicode-bidi: isolate;
    background: transparent;
    color: inherit;
    padding: 0;
}
</style>
""", unsafe_allow_html=True)

SYSTEM_PROMPT = """
أنت استشاري هندسي كهربائي خبير. أجب على أسئلة المهندس بالاعتماد حصرياً على المستندات المرفقة.
يجب أن تكون إجاباتك هندسية دقيقة، ومنظمة في شكل نقاط (Bullet Points)، واستخدم جداول المقارنة عند استعراض الفروق الفنية أو البنود.
اذكر رقم البند أو الصفحة إن أمكن.

قواعد التنسيق (مهمة جداً):
- اكتب بالعربية، والمصطلحات الإنجليزية بين قوسين بعد المصطلح العربي، مثال: كابلات أحادية النواة (Single Core).
- لا تستخدم علامة الدولار ($) ولا صيغ LaTeX نهائياً.
- لا تستخدم علامات الاقتباس العكسية (`) نهائياً.
- اكتب كل الأرقام بالأرقام الإنجليزية (0-9).
- اكتب المقاسات والوحدات كنص عادي، مثال: 4 × 240 مم²، واستخدم جدولاً أو نقاطاً قصيرة بدل جملة عربية طويلة فيها أرقام متتالية.
"""


# ---------------------------------------------------------
# دوال مساعدة
# ---------------------------------------------------------
def clean_text(text):
    """علامة $ بتخلي ستريمليت يعرض الأرقام كمعادلات، والـ backticks بتبوّظ ترتيب العربي."""
    return (text or "").replace("$", "\\$").replace("`", "")


def wait_for_file(client, uploaded):
    """استنى لحد ما الملف المرفوع يبقى جاهز عند جوجل."""
    while uploaded.state.name == "PROCESSING":
        time.sleep(2)
        uploaded = client.files.get(name=uploaded.name)
    if uploaded.state.name == "FAILED":
        raise RuntimeError("فشلت معالجة أحد الملفات عند جوجل. جرّب ملف PDF تاني.")
    return uploaded


def upload_files_to_gemini(client, files):
    refs = []
    for file in files:
        path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f:
                f.write(file.getvalue())
                path = f.name
            uploaded = client.files.upload(file=path)
            refs.append(wait_for_file(client, uploaded))
        finally:
            if path and os.path.exists(path):
                os.unlink(path)
    return refs


def start_chat(client, file_refs):
    """بيبدأ محادثة جديدة والملفات مرفقة في أول رسالة."""
    file_parts = [
        types.Part.from_uri(file_uri=ref.uri, mime_type=ref.mime_type or "application/pdf")
        for ref in file_refs
    ]
    history = [
        types.Content(
            role="user",
            parts=file_parts + [types.Part.from_text(
                text="هذه هي المستندات الهندسية المرجعية. اعتمد عليها حصرياً في إجاباتك."
            )],
        ),
        types.Content(
            role="model",
            parts=[types.Part.from_text(
                text="لقد اطلعت على المستندات الهندسية المرفقة وأنا جاهز للإجابة على الاستفسارات الفنية."
            )],
        ),
    ]
    return client.chats.create(
        model=MODEL_NAME,
        config=types.GenerateContentConfig(system_instruction=SYSTEM_PROMPT),
        history=history,
    )


def build_word_file(text):
    """بيعمل ملف Word من آخر إجابة، والفقرات من اليمين لليسار."""
    doc = Document()
    doc.add_heading("تقرير استشاري هندسي - مواصفات الكهرباء", 0)
    for line in text.split("\n"):
        line = line.strip()
        if not line:
            continue
        if line.startswith("#"):
            para = doc.add_heading(line.lstrip("#").strip().replace("**", ""), level=2)
        else:
            para = doc.add_paragraph(line.replace("**", ""))
        para._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------
# الواجهة
# ---------------------------------------------------------
st.sidebar.title("⚙️ إعدادات النظام الهندسي")
api_key = st.sidebar.text_input("أدخل مفتاح Google AI Studio API", type="password")

st.title("⚡ مساعد مواصفات وعقود الكهرباء (MEP AI Consultant)")

uploaded_files = st.sidebar.file_uploader(
    "ارفع ملفات المواصفات أو الأكواد (PDF)", type=["pdf"], accept_multiple_files=True
)

if "messages" not in st.session_state:
    st.session_state.messages = []

if uploaded_files and api_key:
    api_key = api_key.strip()
    client = genai.Client(api_key=api_key)

    # لو الملفات أو المفتاح اتغيروا، نرفع الملفات ونبدأ محادثة جديدة
    signature = (tuple((f.name, f.size) for f in uploaded_files), api_key)
    if st.session_state.get("signature") != signature:
        try:
            with st.spinner("جاري قراءة وتحليل المستندات الهندسية..."):
                refs = upload_files_to_gemini(client, uploaded_files)
                st.session_state.chat = start_chat(client, refs)
            st.session_state.signature = signature
            st.session_state.messages = []
        except Exception as e:
            st.session_state.pop("signature", None)
            st.error(f"حدث خطأ أثناء رفع المستندات: {e}")
            st.stop()

    # عرض المحادثة السابقة
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(clean_text(msg["text"]))

    # واجهة المحادثة
    user_query = st.chat_input("اكتب استفسارك الفني أو اطلب جدول مقارنة للمواصفات...")
    if user_query:
        st.session_state.messages.append({"role": "user", "text": user_query})
        with st.chat_message("user"):
            st.markdown(user_query)

        with st.chat_message("assistant"):
            try:
                with st.spinner("جاري التحليل..."):
                    response = st.session_state.chat.send_message(user_query)
                answer = response.text or "لم أتمكن من توليد إجابة، جرّب إعادة صياغة السؤال."
            except Exception as e:
                answer = f"حدث خطأ أثناء المعالجة: {e}"
            st.markdown(clean_text(answer))
        st.session_state.messages.append({"role": "assistant", "text": answer})

    # زر تصدير آخر استشارة إلى Word
    st.sidebar.markdown("---")
    st.sidebar.subheader("📄 تصدير المخرجات")
    assistant_msgs = [m for m in st.session_state.messages if m["role"] == "assistant"]
    if assistant_msgs:
        st.sidebar.download_button(
            label="📥 تحميل آخر استشارة كملف Word",
            data=build_word_file(assistant_msgs[-1]["text"]),
            file_name="MEP_Technical_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    else:
        st.sidebar.caption("اسأل سؤالاً أولاً، وبعدها هيظهر زر التحميل.")
else:
    st.info("الرجاء إدخال مفتاح الـ API ورفع مستندات الـ PDF من القائمة الجانبية لبدء الاستشارات الفنية.")
