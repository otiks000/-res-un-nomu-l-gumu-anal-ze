import streamlit as st
import google.generativeai as genai
from PIL import Image
import pdfplumber
import docx
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

st.set_page_config(page_title="Īres Līgumu AI Skeneris", page_icon="🏠", layout="centered")

st.title("🏠 Dzīvokļu Īres Līgumu AI Skeneris")
st.write("Augšupielādē savu īres līgumu (Word .docx, PDF vai attēlu), lai saņemtu juridisko risku un labojumu atskaiti.")

# Sānu izvēlne API atslēgai
api_key = st.sidebar.text_input("Gemini API Atslēga:", type="password")

# Faila augšupielādes komponente ar paplašinātu tipu atbalstu
uploaded_file = st.file_uploader(
    "Izvēlies vai ievelc failu šeit (.docx, .doc, .pdf, .txt, attēls):", 
    type=["docx", "doc", "pdf", "txt", "md", "jpg", "jpeg", "png"]
)

# Inicializējam sesijas mainīgos
if 'document_text' not in st.session_state:
    st.session_state['document_text'] = None
if 'file_type' not in st.session_state:
    st.session_state['file_type'] = None
if 'original_file_bytes' not in st.session_state:
    st.session_state['original_file_bytes'] = None
if 'original_filename' not in st.session_state:
    st.session_state['original_filename'] = "ligums.docx"

if uploaded_file is not None:
    try:
        file_extension = uploaded_file.name.split('.')[-1].lower()
        st.session_state['original_filename'] = uploaded_file.name
        
        # Saglabājam faila baitus sesijā tieši lejupielādei vēlāk
        uploaded_file.seek(0)
        st.session_state['original_file_bytes'] = uploaded_file.read()
        
        if file_extension in ['jpg', 'jpeg', 'png']:
            uploaded_file.seek(0)
            st.session_state['document_text'] = Image.open(uploaded_file)
            st.session_state['file_type'] = "image"
            st.success(f"Attēls veiksmīgi pievienots: {uploaded_file.name}")
            st.image(st.session_state['document_text'], caption="Augšupielādētais attēls", use_column_width=True)
                
        elif file_extension == 'pdf':
            st.session_state['file_type'] = "pdf_text"
            with pdfplumber.open(BytesIO(st.session_state['original_file_bytes'])) as pdf:
                extracted_text = ""
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        extracted_text += text + "\n"
            st.session_state['document_text'] = extracted_text
            st.success(f"PDF fails veiksmīgi nolasīts: {uploaded_file.name}")
            
        elif file_extension in ['docx', 'doc']:
            st.session_state['file_type'] = "text"
            # Droša Word nolasīšana no baitu bufera
            doc = docx.Document(BytesIO(st.session_state['original_file_bytes']))
            extracted_text = "\n".join([paragraph.text for paragraph in doc.paragraphs if paragraph.text.strip()])
            
            if not extracted_text.strip():
                st.warning("Word fails šķiet tukšs vai nesatur nolasāmu tekstu.")
            else:
                st.session_state['document_text'] = extracted_text
                st.success(f"Word fails veiksmīgi nolasīts un sagatavots: {uploaded_file.name}")
            
        elif file_extension in ['txt', 'md']:
            st.session_state['document_text'] = st.session_state['original_file_bytes'].decode("utf-8", errors="ignore")
            st.session_state['file_type'] = "text"
            st.success(f"Teksta fails pievienots: {uploaded_file.name}")

    except Exception as read_err:
        st.error(f"Kļūda nolasot failu ({uploaded_file.name}): {read_err}")

# Palīgfunkcijas rezultātu ģenerēšanai
def create_docx_bytes(text):
    doc = docx.Document()
    doc.add_heading('Īres Līguma Juridiskā Analīze un Labojumi', 0)
    for line in text.split('\n'):
        if line.strip():
            doc.add_paragraph(line)
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

def create_pdf_bytes(text):
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    textobject = p.beginText()
    textobject.setTextOrigin(40, height - 40)
    textobject.setFont("Helvetica", 10)
    
    for line in text.split('\n'):
        safe_line = line.encode('ascii', 'ignore').decode('ascii')
        if len(safe_line) > 90:
            chunks = [safe_line[i:i+90] for i in range(0, len(safe_line), 90)]
            for chunk in chunks:
                textobject.textLine(chunk)
        else:
            textobject.textLine(safe_line)
            
        if textobject.getY() < 40:
            p.drawText(textobject)
            p.showPage()
            textobject = p.beginText()
            textobject.setTextOrigin(40, height - 40)
            textobject.setFont("Helvetica", 10)
            
    p.drawText(textobject)
    p.showPage()
    p.save()
    return buffer.getvalue()

# Analīzes poga
if st.button("Analizēt līgumu"):
    if not api_key:
        st.error("Lūdzu, ievadi Gemini API atslēgu sānu izvēlnē!")
    elif not st.session_state.get('document_text'):
        st.warning("Lūdzu, augšupielādē failu analīzei!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-3.8-flash')
            
            system_prompt = """
            Tu esi pieredzējis Latvijas nekustamā īpašuma un īres tiesību eksperts. Tavs uzdevums ir izanalizēt īres līguma tekstu un izsniegt atskaiti latviešu valodā.
            
            Analīzes noteikumi:
            1. Pārbaudi atbilstību Latvijas Dzīvojamo telpu īres likumam un Civillikumam.
            2. Identificē slēptos riskus, nesamērīgus līgumsodus (kas pārsniedz 10%), neskaidrus depozīta atgriešanas nosacījumus, vienpusējas īres maksas mainīšanas un izlikšanas kārtību.
            
            Atskaites struktūra:
            - 🔴 Kritiski riski (Punkti, kas var radīt finansiālus vai juridiskus zaudējumus)
            - 🟡 Brīdinājumi (Punkti, kuriem jāpievērš uzmanība vai kas jāprecizē)
            - 🟢 Zaļie ieteikumi (Precīzi teikumu labojumi, ko ieteikt ierakstīt līgumā)
            """
            
            with st.spinner("AI lasa un analizē dokumentu..."):
                if st.session_state['file_type'] == "image":
                    response = model.generate_content([system_prompt, st.session_state['document_text']])
                else:
                    full_prompt = f"{system_prompt}\n\nLĪGUMA TEKSTS ANALĪZEI:\n{st.session_state['document_text']}"
                    response = model.generate_content(full_prompt)
                    
                st.session_state['analysis_result'] = response.text
                st.session_state['docx_bytes'] = create_docx_bytes(response.text)
                st.session_state['pdf_bytes'] = create_pdf_bytes(response.text)
                
        except Exception as e:
            st.error(f"Kļūda apstrādē: {e}")

# Izvadām rezultātus un lejupielādes pogas
if 'analysis_result' in st.session_state:
    st.markdown("---")
    st.subheader("📊 Juridiskās analīzes rezultāts:")
    st.markdown(st.session_state['analysis_result'])
    
    st.write("### 📥 Lejupielādes iespējas:")
    
    # Ja lietotājs augšupielādēja failu, piedāvājam to lejupielādēt atpakaļ
    if st.session_state.get('original_file_bytes') is not None:
        st.download_button(
            label="📄 Lejupielādēt oriģinālo augšupielādēto failu",
            data=st.session_state['original_file_bytes'],
            file_name=st.session_state['original_filename'],
            mime="application/octet-stream",
            key="dl_orig_file"
        )
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.download_button(
            label="📥 Labojumi TXT",
            data=st.session_state['analysis_result'],
            file_name="ires_liguma_labojumi.txt",
            mime="text/plain",
            key="dl_txt"
        )
        
    with col2:
        st.download_button(
            label="📥 Labojumi Word (.docx)",
            data=st.session_state['docx_bytes'],
            file_name="ires_liguma_labojumi.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_docx"
        )
        
    with col3:
        st.download_button(
            label="📥 Labojumi PDF",
            data=st.session_state['pdf_bytes'],
            file_name="ires_liguma_labojumi.pdf",
            mime="application/pdf",
            key="dl_pdf"
        )
