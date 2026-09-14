import streamlit as st
import pandas as pd
from docx import Document
import os
import io
import datetime
import time
import google.generativeai as genai

# --- הגדרות עיצוב וממשק ---
st.set_page_config(page_title="אילן פלדמן - משרד עורכי דין", page_icon="🏢", layout="centered")

st.markdown("""
    <style>
    .stApp {
        background-color: #fafafa;
        background-image: 
            linear-gradient(90deg, rgba(0,0,0,0.03) 1px, transparent 1px),
            linear-gradient(0deg, rgba(0,0,0,0.03) 1px, transparent 1px);
        background-size: 40px 40px;
    }
    .main .block-container, p, span, label, div {
        direction: rtl;
        text-align: right;
    }
    .stButton button {
        width: 100%;
        border-radius: 8px;
        font-weight: bold;
    }
    [data-testid="stSidebar"] {
        direction: rtl;
    }
    [data-testid="stSidebar"] * {
        text-align: right;
    }
    </style>
""", unsafe_allow_html=True)

LOGO_FILE = "logo.png"

# --- הגדרת מפתח ה-AI הקלאסי ---
API_KEY = "AIzaSyD2kYYWLQH-yDWgevLHLOKPS8cLxgGwg6g"
try:
    genai.configure(api_key=API_KEY)
except Exception:
    pass

# --- ניהול מצב האפליקציה (Session State) לבחירת פרויקט ---
if 'selected_project' not in st.session_state:
    st.session_state.selected_project = None

PROJECTS = {
    "סמילנסקי": "smilanski",
    "תרנא": "tarna",
    "אקווה": "aqua",
    "ביאליק": "bialik"
}

# ==================== מסך הבית: בחירת פרויקט ====================
if st.session_state.selected_project is None:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists(LOGO_FILE):
            st.image(LOGO_FILE, use_container_width=True)
        else:
            st.markdown("<h2 style='text-align: center;'>אילן פלדמן - משרד עורכי דין</h2>", unsafe_allow_html=True)

    st.markdown("<h1 style='text-align: center;'>בחר פרויקט לניהול</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #555;'>לחץ על אחד הפרויקטים כדי להיכנס למערכת הניהול והפקת החוזים שלו</p>", unsafe_allow_html=True)
    st.markdown("---")

    for heb_name, code in PROJECTS.items():
        col_img, col_btn = st.columns([1, 3])
        
        with col_img:
            building_img_path = os.path.join("projects", code, "building.png")
            if os.path.exists(building_img_path):
                st.image(building_img_path, width=100)
            else:
                st.markdown("🏢 **[תמונת בניין]**")
                
        with col_btn:
            st.markdown(f"### פרויקט {heb_name}")
            if st.button(f"כניסה לפרויקט {heb_name}", key=f"btn_{code}"):
                st.session_state.selected_project = (heb_name, code)
                st.rerun()
        st.markdown("---")

# ==================== מסך פרויקט פעיל ====================
else:
    selected_project_heb, project_code = st.session_state.selected_project

    if st.button("⬅️ חזרה לבחירת פרויקטים"):
        st.session_state.selected_project = None
        st.rerun()

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists(LOGO_FILE):
            st.image(LOGO_FILE, use_container_width=True)

    st.markdown(f"<h1 style='text-align: center;'>פרויקט: {selected_project_heb}</h1>", unsafe_allow_html=True)
    st.markdown("---")

    # הגדרת תיקיות ונתיבים דינאמיים לפרויקט הנבחר
    PROJECT_DIR = os.path.join("projects", project_code)
    os.makedirs(PROJECT_DIR, exist_ok=True)
    os.makedirs(os.path.join(PROJECT_DIR, "templates"), exist_ok=True)

    DB_FILE = os.path.join(PROJECT_DIR, "clients.xlsx")
    TEMPLATE_FILE = os.path.join(PROJECT_DIR, "templates", "lease_template.docx")
    REG_TEMPLATE_FILE = os.path.join(PROJECT_DIR, "templates", "registration_form.docx")

    if not os.path.exists(TEMPLATE_FILE):
        doc = Document()
        doc.add_paragraph("המוכר: {{שם_המוכר}}\nהקונה: {{שם_הקונה}}\nת.ז: {{תז_קונה}}\nטלפון: {{טלפון_קונה}}\nכתובת: {{כתובת_נוכחית}}\nמחיר הנכס: {{מחיר_הנכס}}")
        doc.save(TEMPLATE_FILE)

    if not os.path.exists(REG_TEMPLATE_FILE):
        doc = Document()
        doc.add_paragraph("שם המוכר:\nשם הקונה:\nת\"ז קונה:\nמספר טלפון:\nכתובת נוכחית:\nמחיר הנכס:")
        doc.save(REG_TEMPLATE_FILE)

    def get_current_template_version():
        if os.path.exists(TEMPLATE_FILE):
            mtime = os.path.getmtime(TEMPLATE_FILE)
            return time.strftime("%Y-%m-%d %H:%M", time.localtime(mtime))
        return datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    def extract_data_smart(doc):
        """חילוץ דינאמי וחכם של כל שדה מתוך טופס הוורד"""
        data = {}
        for para in doc.paragraphs:
            text = para.text.strip()
            if ":" in text:
                parts = text.split(":", 1)
                key = parts[0].replace("שדה", "").strip()
                val = parts[1].strip()
                if key and val:
                    data[key] = val
        return data

    def ai_smartify_contract_template(file_buffer):
        """פונקציית AI המשתמשת ב-Gemini להפיכת מסמך וורד רגיל לתבנית מאסטר חכמה עם {{סוגריים}}"""
        try:
            doc = Document(file_buffer)
            full_text_to_process = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            
            model = genai.GenerativeModel('gemini-1.5-flash')
            prompt = f"""
            אתה עוזר משפטי חכם. לפניך טקסט של חוזה נדל"ן. 
            אנא החלף את הפרטים האישיים המשתנים (כגון שמות קונים, מוכרים, תעודות זהות, כתובות, מחירים ותאריכים) 
            במפתחות מתאימים בתוך סוגריים מסולסלים כפולים כמו {{שם_הקונה}}, {{תז_קונה}}, {{שם_המוכר}}, {{מחיר_הנכס}}, {{כתובת_נוכחית}}, {{טלפון_קונה}}.
            החזר את הטקסט המעודכן בלבד, מבלי לוותר על שאר סעיפי החוזה.
            
            הטקסט:
            {full_text_to_process}
            """
            response = model.generate_content(prompt)
            new_text_lines = response.text.split("\n")
            
            new_doc = Document()
            for line in new_text_lines:
                new_doc.add_paragraph(line)
            
            bio = io.BytesIO()
            new_doc.save(bio)
            return bio.getvalue()
        except Exception as e:
            st.error(f"שגיאה בהפעלת שירות ה-AI: {e}")
            return None

    def generate_contract(data):
        doc = Document(TEMPLATE_FILE)
        placeholders = {}
        for k, v in data.items():
            if k not in ["תאריך הוספה", "גרסת חוזה"]:
                clean_key = f"{{{{{k.replace(' ', '_').replace('\"', '')}}}}}"
                placeholders[clean_key] = str(v)
        
        placeholders["{{שם_המוכר}}"] = str(data.get("שם המוכר", ""))
        placeholders["{{שם_הקונה}}"] = str(data.get("שם הקונה", ""))
        placeholders["{{תז_קונה}}"] = str(data.get("ת\"ז קונה", ""))
        placeholders["{{טלפון_קונה}}"] = str(data.get("מספר טלפון", data.get("טלפון", "")))
        placeholders["{{כתובת_נוכחית}}"] = str(data.get("כתובת נוכחית", ""))
        placeholders["{{מחיר_הנכס}}"] = str(data.get("מחיר הנכס", ""))

        for para in doc.paragraphs:
            for key, val in placeholders.items():
                if key in para.text:
                    para.text = para.text.replace(key, val)
        
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()

    def process_and_download(extracted_data):
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        template_ver = get_current_template_version()
        
        extracted_data["תאריך הוספה"] = current_time
        extracted_data["גרסת חוזה"] = template_ver
        
        new_row = pd.DataFrame([extracted_data])
        id_field = 'ת"ז קונה' if 'ת"ז קונה' in extracted_data else list(extracted_data.keys())[1]

        if os.path.exists(DB_FILE):
            df_existing = pd.read_excel(DB_FILE)
            if id_field in df_existing.columns and str(extracted_data.get(id_field)) in df_existing[id_field].astype(str).values:
                for col in extracted_data.keys():
                    if col in df_existing.columns:
                        df_existing.loc[df_existing[id_field].astype(str) == str(extracted_data.get(id_field)), col] = extracted_data[col]
                df_existing.to_excel(DB_FILE, index=False)
                st.success(f"הלקוח כבר קיים במאגר – פרטיו עודכנו בהצלחה לגרסה העדכנית ({template_ver})!")
            else:
                df_updated = pd.concat([df_existing, new_row], ignore_index=True, sort=False)
                df_updated.to_excel(DB_FILE, index=False)
                st.success(f"הקונה החדש נוסף בהצלחה למאגר של {selected_project_heb}.")
        else:
            new_row.to_excel(DB_FILE, index=False)
            st.success(f"נוצר מאגר נתונים חדש עבור פרויקט {selected_project_heb}.")
        
        contract_bytes = generate_contract(extracted_data)
        buyer_name = str(extracted_data.get("שם הקונה", "לקוח_ללא_שם")).replace(" ", "_")
        filename = f"הסכם_מכר_{project_code}_{buyer_name}.docx"
        
        st.download_button(
            label="📥 הורד את הסכם המכר המוכן למחשב",
            data=contract_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    # --- תפריט צדדי (הגדרות מאסטר לפרויקט הנוכחי) ---
    with st.sidebar:
        st.header(f"⚙️ הגדרות: {selected_project_heb}")
        st.markdown("ניהול תבניות המאסטר של פרויקט זה.")
        
        st.markdown("---")
        st.subheader("טופס הרשמה")
        if os.path.exists(REG_TEMPLATE_FILE):
            with open(REG_TEMPLATE_FILE, "rb") as reg_f:
                st.download_button(
                    label="📥 הורד טופס הרשמה ריק",
                    data=reg_f,
                    file_name=f"registration_form_{project_code}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        uploaded_reg_template = st.file_uploader("החלף טופס הרשמה מאסטר", type=["docx"], key="master_reg")
        if uploaded_reg_template is not None:
            if st.button("שמור טופס הרשמה חדש"):
                try:
                    with open(REG_TEMPLATE_FILE, "wb") as f:
                        f.write(uploaded_reg_template.getbuffer())
                    st.success("טופס ההרשמה עודכן!")
                except Exception as e:
                    st.error(f"שגיאה: {e}")

        st.markdown("---")
        st.subheader("חוזה מכר (מאסטר)")
        st.caption(f"גרסת מאסטר נוכחית: {get_current_template_version()}")
        
        uploaded_template = st.file_uploader("העלה חוזה מאסטר רגיל", type=["docx"], key="master_lease")
        
        if uploaded_template is not None:
            if st.button("✨ שדרג תבנית בעזרת AI (הוסף סוגריים מסולסלים אוטומטית)"):
                with st.spinner("הבינה המלאכית מעבדת את החוזה ומוסיפה שדות חכמים..."):
                    processed_bytes = ai_smartify_contract_template(uploaded_template)
                    if processed_bytes:
                        with open(TEMPLATE_FILE, "wb") as f:
                            f.write(processed_bytes)
                        st.success("החוזה שודרג בהצלחה על ידי ה-AI ונשמר כתבנית מאסטר חכמה!")
                        time.sleep(1)
                        st.rerun()

        if st.button("שמור תבנית חוזה כרגיל (ללא AI)"):
            if uploaded_template is not None:
                try:
                    with open(TEMPLATE_FILE, "wb") as f:
                        f.write(uploaded_template.getbuffer())
                    st.success("תבנית החוזה עודכנה בהצלחה!")
                except Exception as e:
                    st.error(f"שגיאה: {e}")

    # --- ממשק ראשי לפרויקט ---
    st.subheader("📌 הוספת לקוח חדש")
    method = st.radio("בחר את אופן הוספת הלקוח:", ["העלאת טופס הרשמה", "הזנה ידנית"], horizontal=True)

    if method == "העלאת טופס הרשמה":
        uploaded_file = st.file_uploader("העלי את טופס ההרשמה של הקונה (קובץ Word)", type=["docx"], key="reg_form")
        
        if uploaded_file is not None:
            if st.button("חלץ נתונים והפק הסכם מכר"):
                try:
                    doc = Document(uploaded_file)
                    data = extract_data_smart(doc)
                    if not data:
                        st.error("לא הצלחתי למצוא נתונים בפורמט מפתח: ערך.")
                    else:
                        process_and_download(data)
                except Exception as e:
                    st.error(f"אירעה שגיאה במהלך העיבוד: {e}")

    else:
        with st.form("manual_entry_form"):
            seller_name = st.text_input("שם המוכר", value='חברת נדל"ן בע"מ')
            buyer_name = st.text_input("שם הקונה")
            buyer_id = st.text_input('ת"ז קונה')
            buyer_phone = st.text_input("מספר טלפון")
            buyer_address = st.text_input("כתובת נוכחית")
            price = st.text_input("מחיר הנכס")
            
            submitted = st.form_submit_button("שמור נתונים והפק הסכם מכר")
            
            if submitted:
                if not buyer_name or not buyer_id:
                    st.error("חובה להזין לפחות שם קונה ותעודת זהות.")
                else:
                    manual_data = {
                        "שם המוכר": seller_name,
                        "שם הקונה": buyer_name,
                        "ת\"ז קונה": buyer_id,
                        "מספר טלפון": buyer_phone,
                        "כתובת נוכחית": buyer_address,
                        "מחיר הנכס": price
                    }
                    process_and_download(manual_data)

    # --- רשימת הלקוחות בפרויקט כטבלה נקייה מסודרת ---
    st.markdown("---")
    st.markdown(f"<h2>📋 רשימת הלקוחות בפרויקט: {selected_project_heb}</h2>", unsafe_allow_html=True)

    if os.path.exists(DB_FILE):
        df = pd.read_excel(DB_FILE)
        
        if "תאריך הוספה" not in df.columns:
            df["תאריך הוספה"] = "לא ידוע"
        if "גרסת חוזה" not in df.columns:
            df["גרסת חוזה"] = "ישנה"

        st.dataframe(df, use_container_width=True)

        st.markdown("### 📥 הורדת חוזים ועדכון גרסאות ללקוחות")
        
        for idx, row in df.iterrows():
            client_data = row.to_dict()
            buyer_name_str = str(client_data.get('שם הקונה', 'לקוח'))
            buyer_id_str = str(client_data.get('ת"ז קונה', ''))
            
            cols = st.columns([3, 1, 1])
            with cols[0]:
                st.text(f"{buyer_name_str} (ת\"ז: {buyer_id_str}) | גרסה: {client_data.get('גרסת חוזה', '')}")
            
            with cols[1]:
                contract_bytes = generate_contract(client_data)
                filename = f"הסכם_מכר_{project_code}_{buyer_name_str.replace(' ', '_')}.docx"
                st.download_button(
                    label="📥 הורד חוזה",
                    data=contract_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"download_{project_code}_{idx}"
                )
                
            with cols[2]:
                if st.button("🔄 עדכן לגרסה חדשה", key=f"update_ver_{project_code}_{idx}"):
                    current_ver = get_current_template_version()
                    df.loc[df.iloc[:, 1].astype(str) == buyer_id_str, 'גרסת חוזה'] = current_ver
                    df.to_excel(DB_FILE, index=False)
                    st.success(f"החוזה עודכן לגרסה {current_ver}!")
                    st.rerun()

        st.markdown("---")
        with open(DB_FILE, "rb") as f:
            st.download_button(
                label=f"📊 הורד את טבלת האקסל המלאה של {selected_project_heb} למחשב",
                data=f,
                file_name=f"clients_{project_code}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"excel_full_{project_code}"
            )
    else:
        st.info(f"עדיין אין קונים במאגר של פרויקט {selected_project_heb}. הטבלה תיווצר אוטומטית כשתזיני את הלקוח הראשון.")