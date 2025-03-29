import streamlit as st
import os
from generate_pdf_optimized3 import generate_pdf_with_title_and_comments, hex_to_rgb
from tempfile import NamedTemporaryFile

# Texte et couleurs personnalisés
first_page_text = (
    "Ce rapport 360° est <b>confidentiel</b> et a été conçu afin de vous fournir une analyse détaillée des informations transmises par les personnes que vous avez sollicitées.<br/><br/>"
    "Il est à la fois <b>une photographie et un outil de progression</b>.<br/><br/>"
    "Grâce à celui-ci, vous allez pouvoir mettre en parallèle votre perception et vous rendre compte qu’elle peut <b>différer avec celle de votre environnement</b>. Le 360° est une démarche constructive, orientée vers le plus d’objectivité et en toute bienveillance.<br/><br/>"
    "Les notes attribuées aux 38 caractéristiques analysées sont restituées sous la forme d’une moyenne. Dans ce rapport, vous trouverez les représentations graphiques de chaque thématique ainsi que vos 10 compétences clés reconnues par votre environnement et 5 que vous pourriez développer davantage.<br/><br/>"
    "Ce rapport reflète également les éventuelles divergences de point de vue relatives à chaque caractéristique évaluée. Principaux objectifs du rapport 360° :<br/>"
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- Prendre conscience de la façon dont votre environnement vous perçoit<br/>"
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- Mieux connaître vos points forts pour en tirer profit<br/>"
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- Identifier les éventuels changements à mettre en œuvre<br/>"
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- Mettre en place un ou plusieurs plans d’action<br/><br/><br/>"
    "<u>L’échelle de mesure utilisée est :</u><br/>"
    "Tout à fait d’accord : 5<br/>"
    "Pas du tout d’accord : 1"
)
text_encadre = (
    "Attention ! Il est important de garder à l'esprit que les personnes sollicitées n'ont pas été formées à l’évaluation. C'est pourquoi, il est primordial de garder une certaine distance et de se concentrer sur les tendances et les points qui reviennent dans les données collectées."
)
custom_color1 = hex_to_rgb("#61A6AB")
custom_color2 = hex_to_rgb("#365B6D")
image_path = "./util/Logo simple_up_carriere.jpg"  # Chemin du logo pour le PDF

# Interface utilisateur
st.title("Générateur de rapport PDF 360°")
st.write("Téléchargez un fichier Excel pour générer un rapport PDF 360°.")

# Téléchargement du fichier Excel
uploaded_file = st.file_uploader("Choisissez un fichier Excel", type=["xlsx", "xls"])

# Saisie du nom de la personne
person_name = st.text_input("Nom de la personne")
# Définir le nom du fichier PDF en fonction du nom de la personne
pdf_name = f"./pdf/rapport_360_{person_name}.pdf"

# Bouton pour lancer la génération du PDF
if uploaded_file and person_name and st.button("Générer le PDF"):
    # Créer un fichier temporaire pour le PDF généré
    with NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        output_pdf_path = tmp_pdf.name

    try:
        # Exécuter la fonction de génération du PDF avec les fichiers et textes appropriés
        generate_pdf_with_title_and_comments(
            input_excel_path=uploaded_file,
            first_page_text=first_page_text,
            image_path=image_path,
            text_encadre=text_encadre,
            pdf_name=pdf_name
        )



        # Vérification que le fichier PDF est généré
        if os.path.exists(pdf_name) and os.path.getsize(pdf_name) > 0:
            with open(pdf_name, "rb") as pdf_file:
                st.download_button(
                    label="Télécharger le PDF généré",
                    data=pdf_file.read(),
                    file_name=pdf_name,
                    mime="application/pdf"
                )
        else:
            st.error("Le fichier PDF n'a pas été généré correctement.")

    except Exception as e:
        st.error(f"Erreur lors de la génération du PDF : {str(e)}")
