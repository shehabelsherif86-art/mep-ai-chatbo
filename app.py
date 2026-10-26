import streamlit as st
from google import genai
from google.genai import types
import tempfile
import os
import time
from io import BytesIO
from datetime import date
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ---------------------------------------------------------
# الإعدادات العامة
# ---------------------------------------------------------
# اسم الموديل (لو جوجل قفلته مستقبلاً غيّره من هنا فقط)
MODEL_NAME = "gemini-3.8-flash"

# موديلات احتياطية لو الأساسي عليه ضغط (خطأ 503). كلها موديلات مختلفة عن الأساسي
# عشان لو واحد عليه ضغط يجرب اللي بعده (Flash-Lite أخف وسعته غالباً أكبر)
FALLBACK_MODELS = ["gemini-3.5-flash-lite", "gemini-3.1-flash-lite"]

# موديلات تحويل الصوت إلى نص (بالترتيب)
TRANSCRIBE_MODELS = [MODEL_NAME, "gemini-3.5-flash-lite", "gemini-3.1-flash-lite"]

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
@st.cache_resource
def get_client(key):
    """
    بنحتفظ بنفس الـ client طول فترة تشغيل التطبيق.
    لو اتعمل client جديد مع كل إعادة تشغيل للصفحة، النسخة القديمة بتتقفل تلقائياً
    والمحادثة المحفوظة بتفضل تستخدمها، وده بيطلّع خطأ: Cannot send a request, as the client has been closed.
    """
    return genai.Client(api_key=key, http_options=types.HttpOptions(timeout=120_000))


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


def is_temporary_error(error):
    """أخطاء مؤقتة من جوجل (ضغط على الموديل أو حد الاستخدام أو تأخر الرد)."""
    msg = str(error).lower()
    kind = type(error).__name__.lower()
    keywords = (
        "503", "unavailable", "429", "resource_exhausted", "500", "internal",
        "overloaded", "504", "deadline", "timed out", "timeout",
    )
    return any(k in msg for k in keywords) or "timeout" in kind


def send_with_retry(client, query):
    """بيعيد المحاولة كذا مرة لو الموديل عليه ضغط، وبعدها بيجرب موديل احتياطي."""
    st.session_state.fallback_used = None
    for attempt in range(3):
        try:
            return st.session_state.chat.send_message(query).text
        except Exception as e:
            if not is_temporary_error(e):
                raise
            time.sleep(2 ** (attempt + 1))  # 2 ثم 4 ثم 8 ثواني

    # كل المحاولات فشلت: جرّب الموديلات الاحتياطية بنفس سجل المحادثة
    config = types.GenerateContentConfig(system_instruction=SYSTEM_PROMPT)
    history = st.session_state.chat.get_history()
    last_error = None
    for model in FALLBACK_MODELS:
        try:
            fallback_chat = client.chats.create(model=model, config=config, history=history)
            answer = fallback_chat.send_message(query).text
            # نرجع للموديل الأساسي في الأسئلة الجاية مع الاحتفاظ بسجل المحادثة
            st.session_state.chat = client.chats.create(
                model=MODEL_NAME, config=config, history=fallback_chat.get_history()
            )
            st.session_state.fallback_used = model
            return answer
        except Exception as e:
            if not is_temporary_error(e):
                raise
            last_error = e
    raise last_error


def transcribe_audio(client, audio_bytes, mime_type="audio/wav"):
    """بيحوّل التسجيل الصوتي إلى نص عربي مكتوب بالحروف العربية."""
    prompt = (
        "هذا تسجيل صوتي لمهندس كهرباء مصري يسأل سؤالاً بالعربية (اللهجة المصرية). "
        "فرّغ الكلام إلى نص مكتوب بالحروف العربية كما قيل بالضبط. "
        "لا تترجم ولا تفسّر ولا تغيّر الكلمات ولا تضف شيئاً. "
        "اكتب الكلمات العربية بحروف عربية حتى لو بدت لك قريبة من كلمات إنجليزية. "
        "لا تكتب كلمة بالإنجليزية إلا إذا نطقها المتحدث فعلاً بالإنجليزية بوضوح "
        "كاختصار أو مصطلح هندسي (مثل MCCB أو XLPE). "
        "إذا لم يكن هناك كلام واضح اكتب فقط: [غير واضح]. "
        "اكتب نص الكلام فقط."
    )
    contents = [types.Part.from_bytes(data=audio_bytes, mime_type=mime_type), prompt]
    config = types.GenerateContentConfig(
        temperature=0.0,
        http_options=types.HttpOptions(timeout=45_000),
    )
    last_error = None
    for model in TRANSCRIBE_MODELS:
        try:
            result = client.models.generate_content(model=model, contents=contents, config=config)
            return (result.text or "").strip()
        except Exception as e:
            if not is_temporary_error(e):
                raise
            last_error = e
    raise last_error


def _rtl_paragraph(paragraph):
    """الفقرة من اليمين لليسار."""
    paragraph._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    return paragraph


def _add_formatted_text(paragraph, text):
    """بيضيف النص للفقرة، وبيحوّل **النص** إلى خط عريض."""
    parts = text.replace("`", "").split("**")
    for i, part in enumerate(parts):
        if not part:
            continue
        bold = i % 2 == 1
        run = paragraph.add_run(part)
        run.bold = bold
        run.font.cs_bold = bold
        run.font.rtl = True


def _is_table_line(line):
    return line.startswith("|") and line.count("|") >= 2


def _is_separator_row(cells):
    return all(c.strip() and set(c.strip()) <= set("-: ") for c in cells)


def _add_table(doc, rows):
    cols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=cols)
    table.style = "Table Grid"
    # الجدول من اليمين لليسار (العنصر لازم ييجي قبل tblW عشان ترتيب XML يبقى صح)
    tbl_pr = table._tbl.tblPr
    bidi = OxmlElement("w:bidiVisual")
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is not None:
        tbl_w.addprevious(bidi)
    else:
        tbl_pr.append(bidi)
    for r_idx, row in enumerate(rows):
        for c_idx in range(cols):
            text = row[c_idx] if c_idx < len(row) else ""
            para = _rtl_paragraph(table.cell(r_idx, c_idx).paragraphs[0])
            _add_formatted_text(para, ("**" + text + "**") if r_idx == 0 and text else text)
    doc.add_paragraph()


def _add_markdown(doc, text):
    """بيحوّل رد الموديل (Markdown) إلى عناوين وقوائم وجداول في Word."""
    lines = (text or "").split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        if _is_table_line(line):
            rows = []
            while i < len(lines) and _is_table_line(lines[i].strip()):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not _is_separator_row(cells):
                    rows.append(cells)
                i += 1
            if rows:
                _add_table(doc, rows)
            continue
        if line.startswith("#"):
            para = doc.add_heading(line.lstrip("#").strip().replace("**", ""), level=3)
            _rtl_paragraph(para)
            for run in para.runs:
                run.font.rtl = True
        elif line.startswith(("- ", "* ", "• ")):
            para = _rtl_paragraph(doc.add_paragraph(style="List Bullet"))
            _add_formatted_text(para, line[2:].strip())
        else:
            para = _rtl_paragraph(doc.add_paragraph())
            _add_formatted_text(para, line)
        i += 1


def build_word_file(messages, file_names=None):
    """بيعمل ملف Word فيه المحادثة كلها (كل الأسئلة وكل الإجابات) بالترتيب."""
    doc = Document()
    title = _rtl_paragraph(doc.add_heading("سجل المحادثة الاستشارية - مواصفات الكهرباء", 0))
    for run in title.runs:
        run.font.rtl = True

    info = _rtl_paragraph(doc.add_paragraph())
    _add_formatted_text(info, "تاريخ التصدير: " + date.today().isoformat())
    if file_names:
        docs_line = _rtl_paragraph(doc.add_paragraph())
        _add_formatted_text(docs_line, "المستندات المرجعية: " + "، ".join(file_names))

    q_no = 0
    for msg in messages:
        if msg["role"] == "user":
            q_no += 1
            heading = _rtl_paragraph(doc.add_heading("السؤال " + str(q_no), level=1))
            for run in heading.runs:
                run.font.rtl = True
            para = _rtl_paragraph(doc.add_paragraph())
            _add_formatted_text(para, msg["text"])
        else:
            heading = _rtl_paragraph(doc.add_heading("الإجابة", level=2))
            for run in heading.runs:
                run.font.rtl = True
            _add_markdown(doc, msg["text"])

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
    client = get_client(api_key)

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

    # الإدخال الصوتي (ميكروفون)
    st.sidebar.markdown("---")
    st.sidebar.subheader("🎤 اسأل بالصوت")
    if "audio_n" not in st.session_state:
        st.session_state.audio_n = 0
    audio = st.sidebar.audio_input(
        "اضغط على الميكروفون، اتكلم، وبعدين اضغط إيقاف",
        key=f"audio_{st.session_state.audio_n}",
    )
    if audio is not None:
        st.sidebar.caption(
            "🎧 اسمع تسجيلك. لو مش واضح اضغط مسح وسجّل تاني، ولو تمام اضغط تحويل لنص."
        )
        col_go, col_del = st.sidebar.columns(2)
        convert_clicked = col_go.button("تحويل لنص ✍️", use_container_width=True)
        delete_clicked = col_del.button("مسح 🗑️", use_container_width=True)

        if delete_clicked:
            # بيمسح التسجيل ويصفّر الميكروفون عشان تسجل من جديد
            st.session_state.pending_voice = None
            st.session_state.audio_n += 1
            st.rerun()

        # مفيش تفريغ تلقائي: التفريغ بيبدأ بس لما تضغط الزر
        if convert_clicked:
            audio_bytes = audio.getvalue()
            try:
                if len(audio_bytes) < 4000:
                    heard = None  # تسجيل فاضي أو بايظ، مفيش داعي نبعته لجوجل
                else:
                    with st.spinner("جاري تحويل صوتك إلى نص... (بحد أقصى دقيقة)"):
                        heard = transcribe_audio(client, audio_bytes, audio.type or "audio/wav")
                if heard and "غير واضح" not in heard:
                    st.session_state.pending_voice = heard
                    st.session_state.pending_voice_box = heard
                else:
                    st.sidebar.warning("ماسمعتش كلام واضح في التسجيل، امسحه وسجّل تاني.")
            except Exception as e:
                if is_temporary_error(e):
                    st.sidebar.error("خوادم جوجل عليها ضغط دلوقتي. استنى دقيقة وجرّب تاني.")
                else:
                    st.sidebar.error(f"حدث خطأ أثناء تحويل الصوت: {e}")

    # مراجعة النص المفرّغ قبل الإرسال (عشان تصحح أي كلمة اتفهمت غلط)
    voice_query = None
    if st.session_state.get("pending_voice"):
        st.sidebar.caption("✏️ ده اللي فهمته من صوتك. عدّله لو فيه غلط، وبعدين اضغط إرسال:")
        edited = st.sidebar.text_area(
            "نص سؤالك", key="pending_voice_box", label_visibility="collapsed"
        )
        col_send, col_cancel = st.sidebar.columns(2)
        if col_send.button("إرسال ✅", use_container_width=True):
            voice_query = edited.strip() or None
            st.session_state.pending_voice = None
            st.session_state.audio_n += 1  # بيصفّر الميكروفون للتسجيل الجاي
        if col_cancel.button("إلغاء ❌", use_container_width=True):
            st.session_state.pending_voice = None
            st.session_state.audio_n += 1
            st.rerun()

    # واجهة المحادثة
    text_query = st.chat_input("اكتب استفسارك الفني أو اطلب جدول مقارنة للمواصفات...")
    user_query = text_query or voice_query
    if user_query:
        st.session_state.messages.append({"role": "user", "text": user_query})
        with st.chat_message("user"):
            st.markdown(user_query)

        with st.chat_message("assistant"):
            try:
                with st.spinner("جاري التحليل... (لو الموديل عليه ضغط هعيد المحاولة تلقائياً)"):
                    answer = send_with_retry(client, user_query)
                answer = answer or "لم أتمكن من توليد إجابة، جرّب إعادة صياغة السؤال."
                st.markdown(clean_text(answer))
                if st.session_state.get("fallback_used"):
                    st.caption("ℹ️ الموديل الأساسي كان عليه ضغط، فالإجابة دي من موديل أخف. راجع الأرقام المهمة بنفسك.")
                st.session_state.messages.append({"role": "assistant", "text": answer})
                if voice_query:
                    st.rerun()
            except Exception as e:
                # مانخزنش الخطأ كإجابة، ونشيل السؤال عشان تقدر تسأله تاني
                st.session_state.messages.pop()
                if is_temporary_error(e):
                    st.error("خوادم جوجل عليها ضغط عالي دلوقتي. استنى دقيقة واسأل السؤال تاني.")
                else:
                    st.error(f"حدث خطأ أثناء المعالجة: {e}")

    # زر تصدير المحادثة كاملة إلى Word
    st.sidebar.markdown("---")
    st.sidebar.subheader("📄 تصدير المخرجات")
    if st.session_state.messages:
        questions_count = sum(1 for m in st.session_state.messages if m["role"] == "user")
        st.sidebar.caption("عدد الأسئلة في المحادثة: " + str(questions_count))
        st.sidebar.download_button(
            label="📥 تحميل المحادثة كاملة (Word)",
            data=build_word_file(st.session_state.messages, [f.name for f in uploaded_files]),
            file_name="MEP_Chat_History.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    else:
        st.sidebar.caption("اسأل سؤالاً أولاً، وبعدها هيظهر زر التحميل.")
else:
    st.info("الرجاء إدخال مفتاح الـ API ورفع مستندات الـ PDF من القائمة الجانبية لبدء الاستشارات الفنية.")
