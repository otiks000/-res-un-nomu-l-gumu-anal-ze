import streamlit as st
import google.generativeai as genai
from PIL import Image
import pdfplumber
import docx
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

st.set_page_config(page_title="Īres Līgumu AI Skeneris", page_icon="🏠")

st.title("🏠 Dzīvokļu Īres Līgumu AI Skeneris")
st.write("Ievadi API atslēgu, augšupielādē līgumu (PDF, Word, attēlu) vai ielīmē tekstu, lai saņemtu juridisko risku atskaiti.")

# Sānu izvēlne API atslēgai
api_key = st.sidebar.text_input("Gemini API Atslēga:", type="password")

# Faila/Foto vai teksta ievade
uploaded_file = st.file_uploader(
    "Augšupielādē līgumu (PDF, Word, TXT, MD) vai nofotografē to (JPG, PNG)", 
    type=["txt", "md", "pdf", "docx", "jpg", "jpeg", "png"]
)

user_text = st.text_area("Vai arī ielīmē līguma tekstu šeit tieši:")

document_data = None
file_type = None

if uploaded_file is not None:
    try:
        file_extension = uploaded_file.name.split('.')[-1].lower()
        
        if file_extension in ['jpg', 'jpeg', 'png']:
            document_data = Image.open(uploaded_file)
            file_type = "image"
            st.success(f"Attēls veiksmīgi pievienots: {uploaded_file.name}")
            st.image(document_data, caption="Augšupielādētais līguma attēls", use_column_width=True)
                
        elif file_extension == 'pdf':
            file_type = "pdf_text"
            with pdfplumber.open(uploaded_file) as pdf:
                extracted_text = ""
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        extracted_text += text + "\n"
            document_data = extracted_text
            st.success(f"PDF fails veiksmīgi nolasīts: {uploaded_file.name}")
            
        elif file_extension == 'docx':
            file_type = "text"
            uploaded_file.seek(0)
            raw_bytes = uploaded_file.read()
            st.session_state['original_docx'] = raw_bytes
            
            doc = docx.Document(BytesIO(raw_bytes))
            extracted_text = "\n".join([paragraph.text for paragraph in doc.paragraphs if paragraph.text.strip()])
            document_data = extracted_text
            st.success(f"Word fails veiksmīgi nolasīts: {uploaded_file.name}")
            
        elif file_extension in ['txt', 'md']:
            document_data = uploaded_file.read().decode("utf-8")
            file_type = "text"
            st.success(f"Teksta fails pievienots: {uploaded_file.name}")

    except Exception as read_err:
        st.error(f"Kļūda nolasot failu: {read_err}")

elif user_text.strip():
    document_data = user_text
    file_type = "text"

# Palīgfunkcija Word (.docx) baiti
def create_docx_bytes(text):
    doc = docx.Document()
    doc.add_heading('Īres Līguma Juridiskā Analīze un Labojumi', 0)
    for line in text.split('\n'):
        if line.strip():
            doc.add_paragraph(line)
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# Palīgfunkcija PDF baiti
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
    elif not document_data:
        st.warning("Lūdzu, augšupielādē failu, foto vai ielīmē tekstu!")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-3.8-flash')
            
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
            
            with st.spinner("AI lasa un analizē dokumentu ar gemini-3.8-flash..."):
                if file_type == "image":
                    response = model.generate_content([system_prompt, document_data])
                else:
                    full_prompt = f"{system_prompt}\n\nLĪGUMA TEKSTS ANALĪZEI:\n{document_data}"
                    response = model.generate_content(full_prompt)
                    
                st.session_state['analysis_result'] = response.text
                st.session_state['docx_bytes'] = create_docx_bytes(response.text)
                st.session_state['pdf_bytes'] = create_pdf_bytes(response.text)
                
        except Exception as e:
            st.error(f"Kļūda apstrādē: {e}")

# Izvadām analīzi un atsevišķi sagatavotās lejupielādes pogas
if 'analysis_result' in st.session_state:
    st.markdown("---")
    st.subheader("📊 Juridiskās analīzes rezultāts:")
    st.markdown(st.session_state['analysis_result'])
    
    st.write("### 📥 Lejupielādes iespējas:")
    
    if 'original_docx' in st.session_state:
        st.download_button(
            label="📄 Lejupielādēt oriģinālo failu Word (.docx)",
            data=st.session_state['original_docx'],
            file_name="oriinalais_ligums.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="btn_orig_docx"
        )
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.download_button(
            label="📥 Labojumi TXT",
            data=st.session_state['analysis_result'],
            file_name="ires_liguma_labojumi.txt",
            mime="text/plain",
            key="btn_txt"
        )
        
    with col2:
        st.download_button(
            label="📥 Labojumi Word (.docx)",
            data=st.session_state['docx_bytes'],
            file_name="ires_liguma_labojumi.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="btn_docx"
        )
        
    with col3:
        st.download_button(
            label="📥 Labojumi PDF",
            data=st.session_state['pdf_bytes'],
            file_name="ires_liguma_labojumi.pdf",
            mime="application/pdf",
            key="btn_pdf"
        )
