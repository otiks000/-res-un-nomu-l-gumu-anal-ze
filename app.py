import streamlit as st
import google.generativeai as genai
import docx
import io

# Lapas dizaina un iestatījumu konfigurācija
st.set_page_config(
    page_title="Dzīvokļu un Nomas Līgumu AI Skeneris",
    page_icon="🏠",
    layout="centered"
)

st.title("🏠 Dzīvokļu un Nomas Līgumu AI Skeneris")
st.markdown("Augšupielādē savu īres līgumu (`.docx`) un saņem juridisko risku analīzi.")

# Sānu josla API atslēgai
st.sidebar.header("🔑 Piekļuves iestatījumi")
api_key = st.sidebar.text_input("Ievadi Gemini API atslēgu", type="password")

# Faila augšupielādes lauks klientiem
uploaded_file = st.file_uploader(
    "Ievelt vai izvēlies savu Word līguma failu (.docx):",
    type=["docx"]
)

SYSTEM_PROMPT = """
Tu esi pieredzējis Latvijas nekustamā īpašuma un īres tiesību eksperts. Tavs uzdevums ir profesionāli izanalizēt augšupielādēto īres līguma tekstu latviešu valodā, pārbaudot tā atbilstību Latvijas Dzīvojamo telpu īres likumam un Civillikumam.

Atskaites struktūra:
- 🔴 Kritiski riski: (Punkti, kas var radīt finansiālus zaudējumus, nesamērīgi līgumsodi virs 10%, neskaidri nosacījumi)
- 🟡 Brīdinājumi: (Vietas, kurām jāpievērš uzmanība)
- 🟢 Zaļie ieteikumi un labojumi: (Konkrēti labojumi un juridiski korekti formulējumi līgumam)
"""

if st.button("🔍 Analizēt līgumu", type="primary"):
    if not api_key:
        st.error("⚠️ Lūdzu, ievadi savu Gemini API atslēgu sānu joslā!")
    elif uploaded_file is None:
        st.error("⚠️ Lūdzu, augšupielādē Word (.docx) failu analīzei!")
    else:
        try:
            # Apstrādājam failu drošā veidā caur baitu plūsmu
            bytes_data = uploaded_file.read()
            file_stream = io.BytesIO(bytes_data)
            
            doc = docx.Document(file_stream)
            contract_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            
            if not contract_text.strip():
                st.warning("⚠️ Fails ir tukšs vai no tā neizdevās nolasīt tekstu.")
            else:
                with st.spinner("Notiek līguma juridiskā analīze... Uzgaidi mirkli."):
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel('gemini-1.5-flash')
                    
                    full_prompt = f"{SYSTEM_PROMPT}\n\nLĪGUMA TEKSTS ANALĪZEI:\n{contract_text}"
                    response = model.generate_content(full_prompt)
                    
                    st.success("✅ Analīze veiksmīgi pabeigta!")
                    st.markdown("---")
                    st.markdown(response.text)
                    
        except Exception as e:
            st.error(f"❌ Kļūda nolasot Word failu: {str(e)}")
