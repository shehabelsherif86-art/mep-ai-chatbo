import streamlit as st
import google.generativeai as genai
import tempfile
import os
from docx import Document

st.set_page_config(page_title="المساعد الهندسي الذكي للمواصفات", layout="wide")

st.sidebar.title("⚙️ إعدادات النظام الهندسي")
api_key = st.sidebar.text_input("أدخل مفتاح Google AI Studio API", type="password")

st.title("⚡ مساعد مواصفات وعقود الكهرباء (MEP AI Consultant)")

# رفع المستندات الهندسية
uploaded_files = st.sidebar.file_uploader(
    "ارفع ملفات المواصفات أو الأكواد (PDF)", type=["pdf"], accept_multiple_files=True
)

if uploaded_files and api_key:
    genai.configure(api_key=api_key.strip())
    
    # رفع الملفات لجيميني
    @st.cache_resource
    def upload_files_to_gemini(files):
        file_refs = []
        for file in files:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f:
                f.write(file.getvalue())
                path = f.name
            file_refs.append(genai.upload_file(path))
        return file_refs

    with st.spinner("جاري قراءة وتحليل المستندات الهندسية..."):
        file_references = upload_files_to_gemini(uploaded_files)

    # إعداد النموذج مع البرومبت الهندسي الصارم
    system_prompt = """
    أنت استشاري هندسي كهربائي خبير. أجب على أسئلة المهندس بالاعتماد حصرياً على المستندات المرفقة.
    يجب أن تكون إجاباتك هندسية دقيقة، ومنظمة في شكل نقاط (Bullet Points)، واستخدم جداول المقارنة عند استعراض الفروق الفنية أو البنود.
    ذكر رقم البند أو الصفحة إن أمكن.
    """
    
    model = genai.GenerativeModel(
        model_name="gemini-1.5-flash",
        system_instruction=system_prompt
    )

    if "chat_session" not in st.session_state:
        st.session_state.chat_session = model.start_chat(history=[{
            "role": "user",
            "parts": file_references + ["لقد اطلعت على المستندات الهندسية المرفقة وأنا جاهز للإجابة على الاستفسارات الفنية بصياغة هندسية وجداول دقيقة."]
        }])

    # واجهة المحادثة
    user_query = st.chat_input("اكتب استفسارك الفني أو اطلب جدول مقارنة للمواصفات...")
    if user_query:
        st.chat_message("user").write(user_query)
        response = st.session_state.chat_session.send_message(user_query)
        st.chat_message("assistant").write(response.text)

    # زر تصدير التقرير إلى ملف Word
    st.sidebar.markdown("---")
    st.sidebar.subheader("📄 تصدير المخرجات")
    if st.sidebar.button("تصدير آخر استشارة إلى مستند Word"):
        doc = Document()
        doc.add_heading("تقرير استشاري هندسي - مواصفات الكهرباء", 0)
        
        # إضافة آخر رد من الشات كمحتوى للتقرير
        last_response = st.session_state.chat_session.history[-1].parts[0].text
        doc.add_paragraph(last_response)
        
        word_path = "MEP_Technical_Report.docx"
        doc.save(word_path)
        
        with open(word_path, "rb") as f:
            st.sidebar.download_button(
                label="📥 تحميل ملف الورد الآن",
                data=f,
                file_name="MEP_Technical_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
else:
    st.info("الرجاء إدخال مفتاح الـ API ورفع مستندات الـ PDF من القائمة الجانبية لبدء الاستشارات الفنية.")