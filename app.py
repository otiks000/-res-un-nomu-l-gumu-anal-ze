import streamlit as st
import google.generativeai as genai
import docx
import io

# Lapas vizuālā konfigurācija
st.set_page_config(
    page_title="Dzīvokļu un Nomas Līgumu AI Skeneris",
    page_icon="🏠",
    layout="centered"
)

# Pielāgots dizains (CSS), lai atskaite izskatītos profesionāli un simpātiski
st.markdown("""
    <style>
    .main-title {
        font-size: 2.2rem;
        color: #1E293B;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }
    .sub-title {
        color: #64748B;
        font-size: 1.05rem;
        margin-bottom: 2rem;
    }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        height: 3em;
        font-weight: 600;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🏠 Dzīvokļu Īres Līgumu AI Skeneris</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Augšupielādē savu Word (.docx) līgumu un saņem tūlītēju juridisko risku un labojumu analīzi.</div>', unsafe_allow_html=True)

# Sānu josla API atslēgai un informācijai
st.sidebar.header("🔑 Piekļuves iestatījumi")
api_key = st.sidebar.text_input("Ievadi Gemini API atslēgu", type="password")

# Faila augšupielāde
uploaded_file = st.file_uploader(
    "Izvēlies vai ievelc savu līguma Word failu (.docx):",
    type=["docx"]
)

SYSTEM_PROMPT = """
Tu esi pieredzējis Latvijas nekustamā īpašuma un īres tiesību eksperts. Tavs uzdevums ir profesionāli izanalizēt īres vai nomas līguma tekstu atbilstoši Latvijas Dzīvojamo telpu īres likumam un Civillikumam.

Izveido ļoti glītu, strukturētu un viegli lasāmu atskaiti, izmantojot šādu formatējumu:

## 📊 Kopējais novērtējums
Noraidāms / Pieņemams ar labojumiem / Drošs (ar īsu 1-2 teikumu kopsavilkumu).

---

### 🔴 Kritiski riski un bīstamie panti
(Norādi punktus, kas var radīt finansiālus zaudējumus, nesamērīgus līgumsodus virs 10%, neiespējamus nosacījumus vai pārkāpj likumu)

### 🟡 Brīdinājumi un slēptie āķi
(Norādi neskaidrus formulējumus, īrniekam/izīrētājam nelabvēlīgas vietas, kurām jāpievērš uzmanība)

### 🟢 Zaļie ieteikumi un konkrēti labojumi
(Piedāvā precīzus, juridiski korektus tekstus/teikumus, kurus ielikt līgumā esošo kļūdu vietā)
"""

if st.button("🔍 Analizēt līgumu tagad", type="primary"):
    if not api_key:
        st.error("⚠️ Lūdzu, ievadi savu Gemini API atslēgu sānu joslā!")
    elif uploaded_file is None:
        st.error("⚠️ Lūdzu, augšupielādē Word (.docx) failu analīzei!")
    else:
        try:
            bytes_data = uploaded_file.getvalue()
            
            if len(bytes_data) == 0:
                st.error("⚠️ Augšupielādētais fails ir tukšs.")
            else:
                target_stream = io.BytesIO(bytes_data)
                doc = docx.Document(target_stream)
                contract_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
                
                if not contract_text.strip():
                    st.warning("⚠️ Fails ir tukšs vai no tā neizdevās nolasīt tekstu.")
                else:
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel('gemini-1.5-flash')
                    full_prompt = f"{SYSTEM_PROMPT}\n\nLĪGUMA TEKSTS ANALĪZEI:\n{contract_text}"
                    
                    st.markdown("### 📋 Analīzes rezultāts:")
                    
                    # 1. TIEŠRAIDES DRAUKĀŠANA / STRAUMĒŠANA: AI sāk drukāt atbildi uzreiz!
                    response = model.generate_content(full_prompt, stream=True)
                    
                    # st.write_stream pakāpeniski attēlo katru AI ģenerēto vārdu reāllaikā
                    st.write_stream(response)
                    
                    st.success("✅ Analīze veiksmīgi pabeigta!")
                        
        except Exception as e:
            st.error(f"❌ Kļūda apstrādājot failu: {str(e)}")
