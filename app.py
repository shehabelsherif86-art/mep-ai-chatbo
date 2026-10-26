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

# موديلات احتياطية لو الأساسي عليه ضغط (خطأ 503)
FALLBACK_MODELS = ["gemini-3.5-flash-lite", "gemini-3.1-flash-lite"]

# موديلات تحويل الصوت إلى نص (بالترتيب)
TRANSCRIBE_MODELS = [MODEL_NAME, "gemini-3.5-flash-lite", "gemini-3.1-flash-lite"]

PAGE_REVIEW = "📑 مراجعة العروض الفنية"
PAGE_CHAT = "💬 الشات الهندسي"

st.set_page_config(page_title="منصة مراجعة العروض الفنية ومواصفات الكهرباء", layout="wide")

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

FORMAT_RULES = """
قواعد التنسيق (مهمة جداً):
- اكتب بالعربية، والمصطلحات الإنجليزية بين قوسين بعد المصطلح العربي، مثال: كابلات أحادية النواة (Single Core).
- لا تستخدم علامة الدولار ($) ولا صيغ LaTeX نهائياً.
- لا تستخدم علامات الاقتباس العكسية (`) نهائياً.
- اكتب كل الأرقام بالأرقام الإنجليزية (0-9).
- اكتب المقاسات والوحدات كنص عادي، مثال: 4 × 240 مم²، واستخدم جدولاً أو نقاطاً قصيرة بدل جملة عربية طويلة فيها أرقام متتالية.
"""

CHAT_SYSTEM_PROMPT = """
أنت استشاري هندسي كهربائي خبير. أجب على أسئلة المهندس بالاعتماد حصرياً على المستندات المرفقة.
المستندات نوعان: مواصفات المشروع، وعروض فنية مقدمة من الموردين (إن وُجدت).
عند الحديث عن عرض مورد اذكر اسم ملفه، وعند المقارنة بين عرض والمواصفات حدد الفروق بدقة.
يجب أن تكون إجاباتك هندسية دقيقة، ومنظمة في شكل نقاط (Bullet Points)، واستخدم جداول المقارنة عند استعراض الفروق الفنية أو البنود.
اذكر رقم البند أو الصفحة إن أمكن.
""" + FORMAT_RULES


def build_review_prompt(spec_names, sub_name):
    return (
        "أنت استشاري هندسي كهربائي محترف وخبير في مراجعة العروض الفنية (Technical Submittals) "
        "ومطابقة المواصفات القياسية لمشاريع البنية التحتية والمنشآت.\n"
        "المستندات المرفقة: مواصفات المشروع (" + "، ".join(spec_names) + ") "
        "ثم العرض الفني المقدم من المورد (" + sub_name + ").\n"
        "قارن العرض الفني بمواصفات المشروع بدقة هندسية تامة، واعتمد على المستندات المرفقة فقط.\n\n"
        "قم بإعداد تقرير هندسي احترافي باللغة العربية يتضمن العناصر التالية:\n"
        "1. **ملخص تنفيذي:** نظرة عامة عن البند محل المقارنة (كابلات، لوحات، محولات، قواطع، إنارة، إلخ).\n"
        "2. **جدول المطابقة الفنية (Technical Compliance Table):** جدول يوضح (المواصفات المطلوبة، "
        "المواصفات المقدمة من المورد، حالة المطابقة: مطابق / غير مطابق / مطابقة بشروط).\n"
        "3. **الانحرافات والاختلافات الجوهرية (Deviations & Discrepancies):** رصد أي قصور أو فروق فنية "
        "عن المواصفات الاستشارية.\n"
        "4. **التوصية النهائية والقرار الهندسي:** (قبول، قبول مع ملاحظات، أو رفض مع إبداء الأسباب الهندسية).\n"
        + FORMAT_RULES
    )


# ---------------------------------------------------------
# دوال مساعدة عامة
# ---------------------------------------------------------
@st.cache_resource
def get_client(key):
    """
    بنحتفظ بنفس الـ client طول فترة تشغيل التطبيق.
    لو اتعمل client جديد مع كل إعادة تشغيل للصفحة، النسخة القديمة بتتقفل تلقائياً
    والمحادثة المحفوظة بتفضل تستخدمها، وده بيطلّع خطأ: Cannot send a request, as the client has been closed.
    """
    return genai.Client(api_key=key, http_options=types.HttpOptions(timeout=240_000))


def clean_text(text):
    """علامة $ بتخلي ستريمليت يعرض الأرقام كمعادلات، والـ backticks بتبوّظ ترتيب العربي."""
    return (text or "").replace("$", "\\$").replace("`", "")


def is_temporary_error(error):
    """أخطاء مؤقتة من جوجل (ضغط على الموديل أو حد الاستخدام أو تأخر الرد)."""
    msg = str(error).lower()
    kind = type(error).__name__.lower()
    keywords = (
        "503", "unavailable", "429", "resource_exhausted", "500", "internal",
        "overloaded", "504", "deadline", "timed out", "timeout",
    )
    return any(k in msg for k in keywords) or "timeout" in kind


def is_timeout_error(error):
    msg = str(error).lower()
    return "timed out" in msg or "timeout" in msg or "timeout" in type(error).__name__.lower()


def file_sig(f):
    """بصمة بسيطة للملف (الاسم + الحجم) عشان مانرفعوش مرتين."""
    return (f.name, f.size)


# ---------------------------------------------------------
# رفع الملفات إلى جوجل
# ---------------------------------------------------------
def wait_for_file(client, uploaded):
    """استنى لحد ما الملف المرفوع يبقى جاهز عند جوجل."""
    while uploaded.state.name == "PROCESSING":
        time.sleep(2)
        uploaded = client.files.get(name=uploaded.name)
    if uploaded.state.name == "FAILED":
        raise RuntimeError("فشلت معالجة أحد الملفات عند جوجل. جرّب ملف PDF تاني.")
    return uploaded


def ensure_uploaded(client, files):
    """بيرفع الملفات الجديدة بس، وبيحتفظ بمرجعها في الجلسة. بيرجع {بصمة الملف: مرجعه}."""
    store = st.session_state.gem_files
    for f in files:
        sig = file_sig(f)
        if sig in store:
            continue
        path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(f.getvalue())
                path = tmp.name
            store[sig] = wait_for_file(client, client.files.upload(file=path))
        finally:
            if path and os.path.exists(path):
                os.unlink(path)
    return store


def labeled_file_parts(label, name, ref):
    """نص صغير بيعرّف الموديل الملف ده إيه، وبعده الملف نفسه."""
    return [
        types.Part.from_text(text=label + ": " + name),
        types.Part.from_uri(file_uri=ref.uri, mime_type=ref.mime_type or "application/pdf"),
    ]


# ---------------------------------------------------------
# مراجعة العروض الفنية
# ---------------------------------------------------------
def generate_with_fallback(client, contents):
    """بيجرب الموديل الأساسي، ولو عليه ضغط أو تأخر بيجرب الموديلات الاحتياطية."""
    last_error = None
    for model in [MODEL_NAME] + FALLBACK_MODELS:
        tries = 2 if model == MODEL_NAME else 1
        for _ in range(tries):
            try:
                result = client.models.generate_content(model=model, contents=contents)
                return (result.text or ""), model
            except Exception as e:
                if not is_temporary_error(e):
                    raise
                last_error = e
                if is_timeout_error(e):
                    break  # مانعيدش المحاولة على موديل اتأخر، ننتقل للي بعده
                time.sleep(3)
    raise last_error


def review_submittal(client, spec_items, sub_name, sub_ref):
    contents = []
    for name, ref in spec_items:
        contents += labeled_file_parts("مواصفات المشروع", name, ref)
    contents += labeled_file_parts("العرض الفني للمورد", sub_name, sub_ref)
    contents.append(types.Part.from_text(text=build_review_prompt([n for n, _ in spec_items], sub_name)))
    return generate_with_fallback(client, contents)


def run_reviews(client, pending, spec_items):
    """بيفحص كل العروض اللي لسه ماتفحصتش واحد ورا التاني."""
    progress = st.progress(0.0)
    for i, sig in enumerate(pending):
        item = st.session_state.submittals[sig]
        progress.progress(
            i / len(pending),
            text="جاري فحص: " + item["name"] + " (" + str(i + 1) + " من " + str(len(pending)) + ")",
        )
        try:
            text, model = review_submittal(client, spec_items, item["name"], item["ref"])
            st.session_state.reports[sig] = {
                "text": text or "لم أتمكن من توليد التقرير، جرّب إعادة الفحص.",
                "model": model,
            }
        except Exception as e:
            if is_temporary_error(e):
                st.error("تعذر فحص " + item["name"] + ": خوادم جوجل عليها ضغط. جرّب تاني بعد دقيقة.")
            else:
                st.error("تعذر فحص " + item["name"] + ": " + str(e))
    progress.progress(1.0)


# ---------------------------------------------------------
# الشات الهندسي
# ---------------------------------------------------------
def prior_history(messages):
    """بيحوّل المحادثة السابقة لسجل جاهز للموديل (من غير سؤال معلّق من غير إجابة)."""
    msgs = list(messages)
    while msgs and msgs[-1]["role"] == "user":
        msgs.pop()
    history = []
    for m in msgs:
        role = "user" if m["role"] == "user" else "model"
        history.append(types.Content(role=role, parts=[types.Part.from_text(text=m["text"])]))
    return history


def start_chat(client, spec_items, sub_items, reports, messages):
    """بيبدأ محادثة جديدة والمستندات (المواصفات + العروض + التقارير) مرفقة في أول رسالة."""
    parts = []
    for name, ref in spec_items:
        parts += labeled_file_parts("مواصفات المشروع", name, ref)
    for name, ref in sub_items:
        parts += labeled_file_parts("عرض فني من مورد", name, ref)

    intro = "هذه هي المستندات المرجعية. اعتمد عليها حصرياً في إجاباتك."
    if reports:
        intro += "\n\nتقارير المراجعة التي تمت بالفعل على العروض الفنية:\n"
        intro += "\n\n".join("### " + name + "\n" + text for name, text in reports)
    parts.append(types.Part.from_text(text=intro))

    history = [
        types.Content(role="user", parts=parts),
        types.Content(
            role="model",
            parts=[types.Part.from_text(
                text="لقد اطلعت على المستندات المرفقة وأنا جاهز للإجابة على الاستفسارات الفنية."
            )],
        ),
    ] + prior_history(messages)

    return client.chats.create(
        model=MODEL_NAME,
        config=types.GenerateContentConfig(system_instruction=CHAT_SYSTEM_PROMPT),
        history=history,
    )


def send_with_retry(client, query):
    """بيعيد المحاولة كذا مرة لو الموديل عليه ضغط، وبعدها بيجرب موديلات احتياطية."""
    st.session_state.fallback_used = None
    for attempt in range(3):
        try:
            return st.session_state.chat.send_message(query).text
        except Exception as e:
            if not is_temporary_error(e):
                raise
            time.sleep(2 ** (attempt + 1))  # 2 ثم 4 ثم 8 ثواني

    # كل المحاولات فشلت: جرّب الموديلات الاحتياطية بنفس سجل المحادثة
    config = types.GenerateContentConfig(system_instruction=CHAT_SYSTEM_PROMPT)
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


# ---------------------------------------------------------
# تصدير ملف Word (تقارير العروض + المحادثة كاملة)
# ---------------------------------------------------------
def _rtl_paragraph(paragraph):
    """الفقرة من اليمين لليسار."""
    paragraph._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    return paragraph


def _rtl_heading(doc, text, level):
    heading = _rtl_paragraph(doc.add_heading(text, level))
    for run in heading.runs:
        run.font.rtl = True
    return heading


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
            _rtl_heading(doc, line.lstrip("#").strip().replace("**", ""), 4)
        elif line.startswith(("- ", "* ", "• ")):
            para = _rtl_paragraph(doc.add_paragraph(style="List Bullet"))
            _add_formatted_text(para, line[2:].strip())
        else:
            para = _rtl_paragraph(doc.add_paragraph())
            _add_formatted_text(para, line)
        i += 1


def build_word_file(messages, spec_names, submittal_names, reports):
    """ملف Word واحد فيه تقارير مراجعة العروض ثم المحادثة كلها بالترتيب."""
    doc = Document()
    _rtl_heading(doc, "تقرير منصة مراجعة العروض الفنية ومواصفات الكهرباء", 0)

    info = _rtl_paragraph(doc.add_paragraph())
    _add_formatted_text(info, "تاريخ التصدير: " + date.today().isoformat())
    if spec_names:
        para = _rtl_paragraph(doc.add_paragraph())
        _add_formatted_text(para, "مواصفات المشروع: " + "، ".join(spec_names))
    if submittal_names:
        para = _rtl_paragraph(doc.add_paragraph())
        _add_formatted_text(para, "العروض الفنية المرفوعة: " + "، ".join(submittal_names))

    if reports:
        _rtl_heading(doc, "أولاً: تقارير مراجعة العروض الفنية", 1)
        for name, text in reports:
            _rtl_heading(doc, "تقرير مراجعة: " + name, 2)
            _add_markdown(doc, text)

    if messages:
        _rtl_heading(doc, "ثانياً: سجل المحادثة الاستشارية", 1)
        q_no = 0
        for msg in messages:
            if msg["role"] == "user":
                q_no += 1
                _rtl_heading(doc, "السؤال " + str(q_no), 2)
                para = _rtl_paragraph(doc.add_paragraph())
                _add_formatted_text(para, msg["text"])
            else:
                _rtl_heading(doc, "الإجابة", 3)
                _add_markdown(doc, msg["text"])

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------
# الحالة (session state)
# ---------------------------------------------------------
def reset_workspace():
    """بيمسح كل حاجة مرتبطة بالمفتاح أو المواصفات القديمة."""
    st.session_state.gem_files = {}
    st.session_state.submittals = {}   # {بصمة الملف: {"name":..., "ref":...}}
    st.session_state.reports = {}      # {بصمة الملف: {"text":..., "model":...}}
    st.session_state.messages = []
    st.session_state.chat = None
    st.session_state.chat_sig = None
    st.session_state.spec_sig = None


if "gem_files" not in st.session_state:
    reset_workspace()
    st.session_state.key_used = None
    st.session_state.sub_up_n = 0
    st.session_state.audio_n = 0

# ---------------------------------------------------------
# الواجهة: القائمة الجانبية
# ---------------------------------------------------------
st.sidebar.title("⚙️ إعدادات المنصة")
api_key = st.sidebar.text_input("أدخل مفتاح Google AI Studio API", type="password").strip()

spec_files = st.sidebar.file_uploader(
    "1️⃣ ارفع مواصفات المشروع (PDF)",
    type=["pdf"],
    accept_multiple_files=True,
    key="spec_uploader",
)

page = st.sidebar.radio("القسم", [PAGE_REVIEW, PAGE_CHAT], key="page")

st.title("⚡ منصة مراجعة العروض الفنية ومواصفات الكهرباء (MEP AI Platform)")

if not api_key or not spec_files:
    st.info("الرجاء إدخال مفتاح الـ API ورفع مواصفات المشروع من القائمة الجانبية لبدء العمل.")
    st.stop()

client = get_client(api_key)

# لو المفتاح اتغير، الملفات المرفوعة بالمفتاح القديم مبقتش صالحة
if st.session_state.key_used != api_key:
    reset_workspace()
    st.session_state.key_used = api_key

# رفع المواصفات (مرة واحدة بس). لو اتغيّرت نبدأ من الأول
spec_sig = tuple(sorted(file_sig(f) for f in spec_files))
if st.session_state.spec_sig != spec_sig:
    try:
        with st.spinner("جاري رفع وقراءة مواصفات المشروع..."):
            ensure_uploaded(client, spec_files)
    except Exception as e:
        st.error("حدث خطأ أثناء رفع المواصفات: " + str(e))
        st.stop()
    st.session_state.reports = {}
    st.session_state.messages = []
    st.session_state.chat = None
    st.session_state.spec_sig = spec_sig

spec_items = [(f.name, st.session_state.gem_files[file_sig(f)]) for f in spec_files]
spec_names = [name for name, _ in spec_items]
subs = st.session_state.submittals
reports = st.session_state.reports

# ---------------------------------------------------------
# صفحة مراجعة العروض الفنية
# ---------------------------------------------------------
if page == PAGE_REVIEW:
    st.header(PAGE_REVIEW)
    st.caption("ارفع عرضاً فنياً أو أكثر. كل عرض بيتقارن مع مواصفات المشروع اللي رفعتها في القائمة الجانبية.")

    new_files = st.file_uploader(
        "2️⃣ ارفع العروض الفنية للموردين (PDF)",
        type=["pdf"],
        accept_multiple_files=True,
        key="sub_uploader_" + str(st.session_state.sub_up_n),
    )
    if new_files:
        try:
            with st.spinner("جاري رفع العروض..."):
                ensure_uploaded(client, new_files)
            for f in new_files:
                sig = file_sig(f)
                subs[sig] = {"name": f.name, "ref": st.session_state.gem_files[sig]}
            st.session_state.sub_up_n += 1  # بيفضّي خانة الرفع
            st.rerun()
        except Exception as e:
            st.error("حدث خطأ أثناء رفع العروض: " + str(e))

    if not subs:
        st.info("لسه مفيش عروض مرفوعة.")
    else:
        pending = [sig for sig in subs if sig not in reports]
        if pending:
            if st.button("🚀 فحص العروض الجديدة (" + str(len(pending)) + ")", type="primary"):
                run_reviews(client, pending, spec_items)

        st.markdown("### العروض المرفوعة")
        for idx, (sig, item) in enumerate(list(subs.items())):
            col_name, col_status, col_del = st.columns([6, 3, 1])
            col_name.markdown("📄 " + item["name"])
            col_status.markdown("✅ تم الفحص" if sig in reports else "⏳ لم يُفحص بعد")
            if col_del.button("🗑️", key="del_" + str(idx), help="حذف هذا العرض"):
                subs.pop(sig, None)
                reports.pop(sig, None)
                st.rerun()

        if reports:
            st.markdown("### 📋 تقارير المراجعة")
            for sig, item in subs.items():
                if sig not in reports:
                    continue
                with st.expander("📋 " + item["name"], expanded=(len(reports) == 1)):
                    st.markdown(clean_text(reports[sig]["text"]))
                    if reports[sig]["model"] != MODEL_NAME:
                        st.caption("ℹ️ الموديل الأساسي كان عليه ضغط، فالتقرير ده من موديل أخف. راجع الأرقام المهمة بنفسك.")

# ---------------------------------------------------------
# صفحة الشات الهندسي
# ---------------------------------------------------------
else:
    st.header(PAGE_CHAT)
    st.caption(
        "الشات بيعرف المواصفات (" + str(len(spec_items)) + " ملف) "
        "والعروض الفنية المرفوعة (" + str(len(subs)) + " ملف) وتقارير المراجعة."
    )

    sub_items = [(item["name"], item["ref"]) for item in subs.values()]
    report_items = [(subs[sig]["name"], reports[sig]["text"]) for sig in subs if sig in reports]

    # لو المستندات أو التقارير اتغيرت، نبدأ محادثة جديدة بنفس السجل السابق
    chat_sig = (spec_sig, tuple(sorted(subs.keys())), tuple(sorted(reports.keys())))
    if st.session_state.chat is None or st.session_state.chat_sig != chat_sig:
        try:
            with st.spinner("جاري تجهيز الشات..."):
                st.session_state.chat = start_chat(
                    client, spec_items, sub_items, report_items, st.session_state.messages
                )
            st.session_state.chat_sig = chat_sig
        except Exception as e:
            st.error("حدث خطأ أثناء تجهيز الشات: " + str(e))
            st.stop()

    # عرض المحادثة السابقة
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(clean_text(msg["text"]))

    # الإدخال الصوتي (ميكروفون)
    st.sidebar.markdown("---")
    st.sidebar.subheader("🎤 اسأل بالصوت")
    audio = st.sidebar.audio_input(
        "اضغط على الميكروفون، اتكلم، وبعدين اضغط إيقاف",
        key="audio_" + str(st.session_state.audio_n),
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
                    st.sidebar.error("حدث خطأ أثناء تحويل الصوت: " + str(e))

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
    text_query = st.chat_input("اكتب استفسارك الفني أو اطلب جدول مقارنة بين العرض والمواصفات...")
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
                    st.error("حدث خطأ أثناء المعالجة: " + str(e))

# ---------------------------------------------------------
# تصدير كل حاجة إلى Word (تقارير العروض + المحادثة كاملة)
# ---------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("📄 تصدير المخرجات")
export_reports = [(subs[sig]["name"], reports[sig]["text"]) for sig in subs if sig in reports]
if export_reports or st.session_state.messages:
    st.sidebar.caption(
        "تقارير العروض: " + str(len(export_reports))
        + " | أسئلة الشات: " + str(sum(1 for m in st.session_state.messages if m["role"] == "user"))
    )
    st.sidebar.download_button(
        label="📥 تحميل التقرير الكامل (Word)",
        data=build_word_file(
            st.session_state.messages,
            spec_names,
            [item["name"] for item in subs.values()],
            export_reports,
        ),
        file_name="MEP_Platform_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
else:
    st.sidebar.caption("افحص عرضاً أو اسأل سؤالاً أولاً، وبعدها هيظهر زر التحميل.")
