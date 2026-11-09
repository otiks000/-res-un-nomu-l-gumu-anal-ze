import streamlit as st
import google.generativeai as genai
from PIL import Image

st.set_page_config(page_title="Īres Līgumu AI Skeneris", page_icon="🏠")

st.title("🏠 Dzīvokļu Īres Līgumu AI Skeneris")
st.write("Ievadi API atslēgu, augšupielādē līguma PDF, nofotografē to vai ielīmē tekstu, lai saņemtu juridisko risku atskaiti.")

# Sānu izvēlne API atslēgai
api_key = st.sidebar.text_input("Gemini API Atslēga:", type="password")

# Faila/Foto vai teksta ievade
uploaded_file = st.file_uploader(
    "Augšupielādē līgumu (PDF, TXT, MD) vai nofotografē to (JPG, PNG)", 
    type=["txt", "md", "pdf", "jpg", "jpeg", "png"]
)

user_text = st.text_area("Vai arī ielīmē līguma tekstu šeit tieši:")

document_data = None
file_type = None

if uploaded_file is not None:
    file_extension = uploaded_file.name.split('.')[-1].lower()
    
    if file_extension in ['jpg', 'jpeg', 'png']:
        document_data = Image.open(uploaded_file)
        file_type = "image"
        st.image(document_data, caption="Augšupielādētais līguma attēls", use_column_width=True)
    elif file_extension == 'pdf':
        document_data = uploaded_file.read()
        file_type = "pdf"
        st.success(f"PDF fails pielādēts: {uploaded_file.name}")
    elif file_extension in ['txt', 'md']:
        document_data = uploaded_file.read().decode("utf-8")
        file_type = "text"

elif user_text:
    document_data = user_text
    file_type = "text"

if st.button("Analizēt līgumu"):
    if not api_key:
        st.error("Lūdzu, ievadi Gemini API atslēgu sānu izvēlnē!")
    elif document_data is None:
        st.warning("Lūdzu, augšupielādē failu, foto vai ielīmē tekstu!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-2.0-flash')
            
            system_prompt = """
            Tu esi pieredzējis Latvijas nekustamā īpašuma un īres tiesību eksperts. Tavs uzdevums ir izanalizēt īres līguma tekstu (vai attēlu/PDF) un izsniegt atskaiti latviešu valodā.
            
            Analīzes noteikumi:
            1. Pārbaudi atbilstību Latvijas Dzīvojamo telpu īres likumam un Civillikumam.
            2. Identificē slēptos riskus, nesamērīgus līgumsodus (kas pārsniedz 10%), neskaidrus depozīta atgriešanas nosacījumus, vienpusējas īres maksas mainīšanas un izlikšanas kārtību.
            
            Atskaites struktūra:
            - 🔴 Kritiski riski (Punkti, kas var radīt finansiālus vai juridiskus zaudējumus)
            - 🟡 Brīdinājumi (Punkti, kuriem jāpievērš uzmanība vai kas jāprecizē)
            - 🟢 Zaļie ieteikumi (Precīzi teikumu labojumi, ko ieteikt ierakstīt līgumā)
            """
            
            with st.spinner("AI lasa un analizē dokumentu..."):
                if file_type == "image":
                    response = model.generate_content([system_prompt, document_data])
                elif file_type == "pdf":
                    with st.spinner("Apstrādāju PDF failu..."):
                        sample_file = genai.upload_file(uploaded_file, mime_type="application/pdf")
                        response = model.generate_content([system_prompt, sample_file])
                else:
                    full_prompt = f"{system_prompt}\n\nLĪGUMA TEKSTS ANALĪZEI:\n{document_data}"
                    response = model.generate_content(full_prompt)
                    
                st.markdown(response.text)
                
        except Exception as e:
            st.error(f"Kļūda apstrādē: {e}")
