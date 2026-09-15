import streamlit as st
import os
import re
import time
import tempfile
import unicodedata
import pandas as pd
from docx import Document
from pptx import Presentation
from google import genai
from google.genai import errors, types

# --- CONFIGURATION PAGE ---
st.set_page_config(page_title="Convertisseur IA Markdown", page_icon="📄", layout="centered")

st.title("📄 Convertisseur Intelligent vers Markdown")
st.write("Transformez vos documents (PDF, Word, Excel, PPT) en format Markdown optimisé pour l'IA.")

# --- BARRE LATÉRALE ---
with st.sidebar:
    st.header("⚙️ Configuration")
    api_key = st.text_input("Clé API Gemini", type="password", help="Obtenez-la sur Google AI Studio")
    st.markdown("---")
    st.write("**Formats supportés :**")
    st.write("✅ PDF (.pdf)\n✅ Word (.docx)\n✅ Excel (.xlsx)\n✅ PowerPoint (.pptx)")

MODEL_NAME = "gemini-flash-latest"
MAX_TENTATIVES = 4

CONFIG_GENERATION = types.GenerateContentConfig(
    max_output_tokens=65536,
    safety_settings=[
        types.SafetySetting(category='HARM_CATEGORY_DANGEROUS_CONTENT', threshold='BLOCK_ONLY_HIGH'),
    ],
)

PROMPT_PDF = """
Tu es un assistant expert en analyse de documents de recherche et données scientifiques.
Analyse ce document. Convertis l'intégralité de son contenu en format Markdown (.md).
Règles strictes :
1. Conserve la hiérarchie d'origine (Titres en #, ##, etc.).
2. Reproduis fidèlement les tableaux.
3. Décris textuellement les schémas, graphiques ou formules complexes.
4. Ne génère aucune introduction ni conclusion.
"""

PROMPT_TEXTE = """
Tu es un expert en structuration de documents Markdown à partir de texte brut.
Reconstruis le texte fourni en Markdown propre et structuré.
- Utilise les balises de titres (#, ##), les listes et les tableaux.
- Ne modifie jamais le contenu, les chiffres, ou les noms propres.
- Si un passage est illisible, signale-le par [illisible].
- Aucun texte introductif ou de clôture.
TEXTE À CONVERTIR :
"""

def nom_ascii_securise(nom: str) -> str:
    sans_accents = unicodedata.normalize('NFKD', nom).encode('ascii', 'ignore').decode('ascii')
    return re.sub(r'[^A-Za-z0-9_.-]+', '_', sans_accents)

def extraire_excel(filepath):
    dfs = pd.read_excel(filepath, sheet_name=None)
    return "".join([f"\n\n[FEUILLE EXCEL : {nom}]\n" + df.to_string(index=False) for nom, df in dfs.items()])

def extraire_word(filepath):
    return "\n".join([p.text for p in Document(filepath).paragraphs if p.text.strip() != ""])

def extraire_ppt(filepath):
    return "".join([f"\n\n[DIAPOSITIVE {i+1}]\n" + "".join([shape.text + "\n" for shape in slide.shapes if hasattr(shape, "text")]) for i, slide in enumerate(Presentation(filepath).slides)])

# --- INTERFACE PRINCIPALE ---
uploaded_file = st.file_uploader("Déposez votre fichier ici", type=['pdf', 'docx', 'xlsx', 'pptx'])

if uploaded_file and api_key:
    if st.button("🚀 Lancer la conversion", type="primary"):
        try:
            client = genai.Client(api_key=api_key)
            ext = uploaded_file.name.split('.')[-1].lower()
            nom_original = uploaded_file.name
            nom_safe = nom_ascii_securise(nom_original)
            
            with st.status("Traitement en cours...", expanded=True) as status:
                with tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}") as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = tmp.name

                response = None
                
                if ext == 'pdf':
                    st.write(f"⏳ Upload de '{nom_safe}' vers Gemini...")
                    gemini_file = client.files.upload(file=tmp_path, config={'display_name': nom_safe})
                    
                    st.write("🧠 Analyse visuelle et conversion...")
                    for tentative in range(1, MAX_TENTATIVES + 1):
                        try:
                            response = client.models.generate_content(
                                model=MODEL_NAME, contents=[gemini_file, PROMPT_PDF], config=CONFIG_GENERATION
                            )
                            break
                        except errors.APIError as e:
                            if e.code in [429, 503] and tentative < MAX_TENTATIVES:
                                time.sleep(10 * tentative)
                            else: raise
                    
                    client.files.delete(name=gemini_file.name)

                else:
                    st.write("⚙️ Extraction locale du texte...")
                    texte_brut = extraire_word(tmp_path) if ext == 'docx' else extraire_excel(tmp_path) if ext == 'xlsx' else extraire_ppt(tmp_path)

                    st.write("🧠 Formatage IA en Markdown...")
                    for tentative in range(1, MAX_TENTATIVES + 1):
                        try:
                            response = client.models.generate_content(
                                model=MODEL_NAME, contents=f"{PROMPT_TEXTE}\n\n{texte_brut}", config=CONFIG_GENERATION
                            )
                            break
                        except errors.APIError as e:
                            if e.code in [429, 503] and tentative < MAX_TENTATIVES:
                                time.sleep(10 * tentative)
                            else: raise

                os.remove(tmp_path)

                if response and response.text:
                    status.update(label="Conversion terminée !", state="complete", expanded=False)
                    st.success("🎉 Fichier prêt !")
                    st.download_button("⬇️ Télécharger le fichier .md", data=response.text, file_name=f"{nom_original.rsplit('.', 1)[0]}.md", mime="text/markdown")
                    with st.expander("👀 Aperçu du Markdown"):
                        st.markdown(response.text)
                else:
                    status.update(label="Échec de la conversion", state="error")
                    st.error("L'IA n'a renvoyé aucun texte.")

        except Exception as e:
            st.error(f"Erreur : {e}")

elif uploaded_file and not api_key:
    st.warning("👈 Entrez votre clé API Gemini dans la barre latérale pour commencer.")
