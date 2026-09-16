import streamlit as st
import os
import re
import time
import tempfile
import unicodedata
import pandas as pd
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from pptx import Presentation
from google import genai
from google.genai import errors, types

# --- CONFIGURATION PAGE ---
st.set_page_config(page_title="Convertisseur IA Markdown", page_icon="📄", layout="centered")

st.title("📄 Convertisseur Intelligent vers Markdown")
st.write("Transformez vos documents (PDF, Word, Excel, PPT) en format Markdown optimisé pour l'IA.")

# --- RÉCUPÉRATION AUTOMATIQUE DE LA CLÉ API ---
try:
    api_key_secret = st.secrets.get("GEMINI_API_KEY", "")
except Exception:
    api_key_secret = ""

# --- BARRE LATÉRALE ---
with st.sidebar:
    st.header("⚙️ Configuration")
    if api_key_secret:
        st.success("✅ Clé API chargée automatiquement !")
        api_key = api_key_secret
    else:
        api_key = st.text_input("Clé API Gemini", type="password", help="Obtenez-la sur Google AI Studio")
    st.markdown("---")
    st.write("**Formats supportés :**")
    st.write("✅ PDF (.pdf)\n✅ Word (.docx)\n✅ Excel (.xlsx)\n✅ PowerPoint (.pptx)")

# Modèles essayés dans l'ordre : le standard, puis le "lite" en secours si
# le premier est saturé (503) ou si son quota personnel est épuisé (429).
# Les deux sont gratuits et tournent sur des pools de serveurs distincts.
MODELES_A_ESSAYER = ["gemini-flash-latest", "gemini-flash-lite-latest"]
TENTATIVES_PAR_MODELE = 2

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
TEXTE A CONVERTIR :
"""

# --- FONCTIONS UTILITAIRES ---

def nom_ascii_securise(nom: str) -> str:
    sans_accents = unicodedata.normalize('NFKD', nom).encode('ascii', 'ignore').decode('ascii')
    return re.sub(r'[^A-Za-z0-9_.-]+', '_', sans_accents)


def diagnostiquer_reponse_vide(response) -> str:
    if response and response.prompt_feedback and response.prompt_feedback.block_reason:
        return f"Bloque par securite - raison : {response.prompt_feedback.block_reason}"
    if response and response.candidates:
        candidat = response.candidates[0]
        if candidat.finish_reason and str(candidat.finish_reason) != "STOP":
            return f"Interrompu - {candidat.finish_reason}"
    return "Aucune explication fournie par l'API."


def generer_avec_repli(client, contents):
    """Essaie chaque modele de la liste l'un apres l'autre. Une erreur 429
    (quota) ou 503 (surcharge) fait basculer sur le modele suivant plutot
    que d'attendre indefiniment sur le meme."""
    derniere_erreur = None
    for nom_modele in MODELES_A_ESSAYER:
        for tentative in range(1, TENTATIVES_PAR_MODELE + 1):
            try:
                st.write(f"   -> Essai avec {nom_modele} (tentative {tentative}/{TENTATIVES_PAR_MODELE})...")
                return client.models.generate_content(
                    model=nom_modele, contents=contents, config=CONFIG_GENERATION
                )
            except errors.ClientError as e:
                derniere_erreur = e
                if e.code == 429 and tentative < TENTATIVES_PAR_MODELE:
                    time.sleep(10 * tentative)
                else:
                    break
            except errors.ServerError as e:
                derniere_erreur = e
                if tentative < TENTATIVES_PAR_MODELE:
                    time.sleep(10 * tentative)
                else:
                    break
        st.write(f"   ⚠️ {nom_modele} indisponible, bascule sur le modele suivant...")
    raise derniere_erreur


def extraire_excel(filepath):
    dfs = pd.read_excel(filepath, sheet_name=None)
    morceaux = []
    for nom, df in dfs.items():
        morceaux.append(f"\n\n[FEUILLE EXCEL : {nom}]\n" + df.to_markdown(index=False))
    return "".join(morceaux)


def _iterer_blocs_word(document):
    """Parcourt le corps du document dans l'ordre reel d'apparition,
    paragraphes et tableaux confondus (au lieu de sauter les tableaux)."""
    for enfant in document.element.body.iterchildren():
        if enfant.tag.endswith('}p'):
            yield Paragraph(enfant, document)
        elif enfant.tag.endswith('}tbl'):
            yield Table(enfant, document)


def extraire_word(filepath):
    document = Document(filepath)
    morceaux = []
    for bloc in _iterer_blocs_word(document):
        if isinstance(bloc, Paragraph):
            if bloc.text.strip():
                morceaux.append(bloc.text)
        elif isinstance(bloc, Table):
            lignes = ["\t".join(cell.text.strip() for cell in row.cells) for row in bloc.rows]
            morceaux.append("\n[TABLEAU]\n" + "\n".join(lignes))
    return "\n".join(morceaux)


def extraire_ppt(filepath):
    prs = Presentation(filepath)
    morceaux = []
    for i, slide in enumerate(prs.slides):
        morceaux.append(f"\n\n[DIAPOSITIVE {i + 1}]")
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                morceaux.append(shape.text_frame.text)
            elif shape.has_table:
                lignes = ["\t".join(cell.text.strip() for cell in row.cells) for row in shape.table.rows]
                morceaux.append("[TABLEAU]\n" + "\n".join(lignes))
    return "\n".join(morceaux)


# --- INTERFACE PRINCIPALE ---
uploaded_file = st.file_uploader("Deposez votre fichier ici", type=['pdf', 'docx', 'xlsx', 'pptx'])

if uploaded_file and api_key:
    if st.button("🚀 Lancer la conversion", type="primary"):
        try:
            client = genai.Client(api_key=api_key)
            ext = uploaded_file.name.split('.')[-1].lower()
            nom_original = uploaded_file.name
            nom_safe = nom_ascii_securise(nom_original)

            with st.status("Traitement en cours...", expanded=True) as status:
                tmp_path = None
                gemini_file = None
                response = None

                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}") as tmp:
                        tmp.write(uploaded_file.getvalue())
                        tmp_path = tmp.name

                    if ext == 'pdf':
                        st.write(f"⏳ Upload de '{nom_safe}' vers Gemini...")
                        gemini_file = client.files.upload(file=tmp_path, config={'display_name': nom_safe})
                        st.write("🧠 Analyse visuelle et conversion...")
                        response = generer_avec_repli(client, [gemini_file, PROMPT_PDF])
                    else:
                        st.write("⚙️ Extraction locale du texte...")
                        if ext == 'docx':
                            texte_brut = extraire_word(tmp_path)
                        elif ext == 'xlsx':
                            texte_brut = extraire_excel(tmp_path)
                        else:
                            texte_brut = extraire_ppt(tmp_path)

                        st.write("🧠 Formatage IA en Markdown...")
                        response = generer_avec_repli(client, f"{PROMPT_TEXTE}\n\n{texte_brut}")

                finally:
                    if gemini_file:
                        try:
                            client.files.delete(name=gemini_file.name)
                        except Exception:
                            pass
                    if tmp_path and os.path.exists(tmp_path):
                        os.remove(tmp_path)

                if response and response.text:
                    status.update(label="Conversion terminee !", state="complete", expanded=False)
                    st.success("🎉 Fichier pret !")
                    st.download_button(
                        "⬇️ Telecharger le fichier .md",
                        data=response.text,
                        file_name=f"{nom_original.rsplit('.', 1)[0]}.md",
                        mime="text/markdown",
                    )
                    with st.expander("👀 Apercu du Markdown"):
                        st.markdown(response.text)
                else:
                    raison = diagnostiquer_reponse_vide(response)
                    status.update(label="Echec de la conversion", state="error")
                    st.error(f"⚠️ L'IA n'a renvoye aucun texte.\n\n**Diagnostic :** {raison}")

        except Exception as e:
            st.error(f"Erreur inattendue : {e}")

elif uploaded_file and not api_key:
    st.warning("👈 Veuillez verifier votre cle API dans les secrets ou la saisir manuellement.")
