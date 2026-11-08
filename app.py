import streamlit as st
import google.generativeai as genai

st.set_page_config(page_title="Īres Līgumu AI Skeneris", page_icon="🏠")

st.title("🏠 Dzīvokļu Īres Līgumu AI Skeneris")
st.write("Ievadi savu API atslēgu, ielādē līguma tekstu un saņem atskaiti par juridiskiem riskiem.")

# Sānu izvēlne API atslēgai
api_key = st.sidebar.text_input("Gemini API Atslēga:", type="password")

# Faila vai teksta ievade
uploaded_file = st.file_uploader("Augšupielādē līguma failu (.txt vai .md)", type=["txt", "md"])
user_text = st.text_area("Vai arī ielīmē līguma tekstu šeit tieši:")

contract_content = ""
if uploaded_file:
    contract_content = uploaded_file.read().decode("utf-8")
elif user_text:
    contract_content = user_text

if st.button("Analizēt līgumu"):
    if not api_key:
        st.error("Lūdzu, ievadi Gemini API atslēgu sānu izvēlnē!")
    elif not contract_content:
        st.warning("Lūdzu, augšupielādē līguma failu vai ielīmē tekstu!")
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
            
            with st.spinner("AI analizē līguma punktus..."):
                full_prompt = f"{system_prompt}\n\nLĪGUMA TEKSTS ANALĪZEI:\n{contract_content}"
                response = model.generate_content(full_prompt)
                st.markdown(response.text)
        except Exception as e:
            st.error(f"Kļūda apstrādē: {e}")
