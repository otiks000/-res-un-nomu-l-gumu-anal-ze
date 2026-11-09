import streamlit as st
from google import genai
import docx
import io

st.set_page_config(
    page_title="Dzīvokļu un Nomas Līgumu AI Skeneris",
    page_icon="🏠",
    layout="centered"
)

st.title("🏠 Dzīvokļu un Nomas Līgumu AI Skeneris")
st.markdown("Augšupielādē savu īres līgumu (`.docx`) un saņem juridisko risku analīzi.")

api_key = st.sidebar.text_input("Ievadi Gemini API atslēgu", type="password")

uploaded_file = st.file_uploader("Ievietojiet savu Word līguma failu (.docx):", type=["docx"])

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
            bytes_data = uploaded_file.getvalue()

            if len(bytes_data) == 0:
                st.error("⚠️ Augšupielādētais fails ir tukšs.")
            else:
                doc = docx.Document(io.BytesIO(bytes_data))

                # Rindkopas
                paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]

                # Tabulas
                for table in doc.tables:
                    for row in table.rows:
                        cells = [c.text.strip() for c in row.cells if c.text.strip()]
                        if cells:
                            paragraphs.append(" | ".join(cells))

                contract_text = "\n".join(paragraphs)

                if not contract_text.strip():
                    st.warning("⚠️ Fails ir tukšs vai no tā neizdevās nolasīt tekstu.")
                else:
                    with st.spinner("Notiek līguma juridiskā analīze... Uzgaidi mirkli."):
                        client = genai.Client(api_key=api_key)

                        full_prompt = f"{SYSTEM_PROMPT}\n\nLĪGUMA TEKSTS ANALĪZEI:\n{contract_text}"
                        response = client.models.generate_content(
                            model="gemini-3.8-flash",
                            contents=full_prompt,
                        )

                        st.success("✅ Analīze veiksmīgi pabeigta!")
                        st.markdown("---")
                        st.markdown(response.text)

        except Exception as e:
            st.error(f"❌ Kļūda apstrādājot failu: {str(e)}")
