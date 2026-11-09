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
st.markdown("Augšupielādē īres līgumu (`.docx`) vai ielīmē tā tekstu un saņem juridisko risku analīzi.")

api_key = st.sidebar.text_input("Ievadi Gemini API atslēgu", type="password")

uploaded_file = st.file_uploader("Ievietojiet Word līguma failu (.docx):")
pasted_text = st.text_area("Vai ielīmē līguma tekstu šeit:", height=200)

SYSTEM_PROMPT = """
Tu esi pieredzējis Latvijas nekustamā īpašuma un īres tiesību eksperts. Tavs uzdevums ir profesionāli izanalizēt augšupielādēto īres līguma tekstu latviešu valodā, pārbaudot tā atbilstību Latvijas Dzīvojamo telpu īres likumam un Civillikumam.

Atskaites struktūra:
- 🔴 Kritiski riski: (Punkti, kas var radīt finansiālus zaudējumus, nesamērīgi līgumsodi virs 10%, neskaidri nosacījumi)
- 🟡 Brīdinājumi: (Vietas, kurām jāpievērš uzmanība)
- 🟢 Zaļie ieteikumi un labojumi: (Konkrēti labojumi un juridiski korekti formulējumi līgumam)
"""


def read_docx(file_bytes):
    doc = docx.Document(io.BytesIO(file_bytes))
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                parts.append(" | ".join(cells))
    return "\n".join(parts)


if st.button("🔍 Analizēt līgumu", type="primary"):
    if not api_key:
        st.error("⚠️ Lūdzu, ievadi savu Gemini API atslēgu sānu joslā!")
    elif uploaded_file is None and not pasted_text.strip():
        st.error("⚠️ Augšupielādē .docx failu vai ielīmē līguma tekstu!")
    else:
        try:
            if uploaded_file is not None:
                if not uploaded_file.name.lower().endswith(".docx"):
                    st.error("⚠️ Atbalstīti tikai .docx faili.")
                    st.stop()
                contract_text = read_docx(uploaded_file.getvalue())
            else:
                contract_text = pasted_text

            if not contract_text.strip():
                st.warning("⚠️ Neizdevās nolasīt tekstu.")
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
            st.error(f"❌ Kļūda: {str(e)}")
