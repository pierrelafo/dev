import pandas as pd
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from io import BytesIO
from PIL import Image
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

arial_font_path = r'C:\Windows\Fonts\Arial.ttf'  # Chemin vers la police Arial
pdfmetrics.registerFont(TTFont('Arial', arial_font_path))

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) / 255.0 for i in (0, 2, 4))
	
def generate_pdf_with_title_and_comments(input_excel_path, first_page_text, last_page_text):

    def add_header_image(c):
        c.drawImage(image_path, (page_width-102)/2, page_height - 50, width=102, height=40)  # Ajustez les coordonnées selon vos besoins

    # Charger les données Excel
    sheet_1 = pd.read_excel(input_excel_path, sheet_name=0, header=2)
    sheet_2 = pd.read_excel(input_excel_path, sheet_name=1, header=2)

    # Lire les questions (ligne 2)
    categories = pd.read_excel(input_excel_path, sheet_name=0, header=None).iloc[0]
    questions_sheet_1 = pd.read_excel(input_excel_path, sheet_name=0, header=None).iloc[1]

    # Extraire prénom et nom depuis le premier onglet
    prenom = pd.read_excel(input_excel_path, sheet_name=0).iloc[2, 3]  # Supposition : première cellule
    nom = pd.read_excel(input_excel_path, sheet_name=0).iloc[2, 2]     # Supposition : deuxième cellule

    # Trouver les indices des colonnes pour les questions
    start_col = questions_sheet_1[questions_sheet_1.str.contains("A l'aise dans l'échange", case=False, na=False)].index[0]
    end_col = questions_sheet_1[questions_sheet_1.str.contains("Je prends des initiatives", case=False, na=False)].index[0]

    # Questions "forces" et "points d'amélioration"
    force_col = 44  # Colonne AS
    amelioration_col = 45  # Colonne AT
    #output_pdf_path = f"D:\\Drive LafontaineConsulting Microsoft\\OneDrive - Lafontaine Consulting\\Nouveau\\xls_to_pdf\\UP-CARRIERE 360 {prenom} {nom}.pdf"
    output_pdf_path = f"./UP-CARRIERE 360 {prenom} {nom}.pdf"
    
	
	# Créer le fichier PDF avec ReportLab
    c = canvas.Canvas(output_pdf_path, pagesize=A4)
    page_width, page_height = A4

    # Ajouter la première page avec le titre et un libellé personnalisé
    add_header_image(c)
	
    custom_color1 = hex_to_rgb("#61A6AB")
    custom_color2 = hex_to_rgb("#365B6D")
	
    # Définir la couleur du texte en utilisant la couleur RGB convertie
    c.setFillColorRGB(*custom_color2)
    c.setFont("Arial", 20)	
    c.drawCentredString(page_width/2, 600, f"360° de {prenom} {nom}")
    c.setFillColorRGB(*custom_color1)
    c.setFont("Arial", 12)
    c.drawCentredString(page_width/2, 400, first_page_text)
    c.showPage()
    add_header_image(c)
    c.setFillColorRGB(*custom_color2)
    # Générer les graphiques comme dans le script précédent
    questions_to_plot = questions_sheet_1[start_col-2:end_col+1]
    previous_category = None
    charts_per_page = 7
    chart_width = 400
    chart_height = 80
    margin_top = 750
    chart_counter = 0

    for idx, question in questions_to_plot.items():
        # Détecter la catégorie correspondante à la question
        current_category = categories[idx]

        # Si la catégorie change, ajouter le titre de la nouvelle catégorie
        if pd.notna(current_category) and current_category != previous_category and current_category.strip() != "" and idx>3:
            previous_category = current_category
            c.setFillColorRGB(*custom_color2)
            c.setFont("Arial", 14)
            c.drawString(40, margin_top, f"SECTION {current_category}")
            margin_top -= 20  # Ajouter un espace après le titre de catégorie
        auto_evaluation_score = pd.to_numeric(sheet_1.iloc[0, idx], errors='coerce')
        evaluators_scores = pd.to_numeric(sheet_2.iloc[:, idx+2], errors='coerce')
        average_evaluators_score = evaluators_scores.mean()

        if pd.isna(auto_evaluation_score) or pd.isna(average_evaluators_score):
            continue

        #plt.rcParams['font.family'] = font_prop.get_name()
        plt.figure(figsize=(6, 1), facecolor='#F2F1EC')
        bars = plt.barh(['Mon évaluation', 'Moyenne des évaluateurs'], [auto_evaluation_score, average_evaluators_score], height=0.5, color=['#365B6D', '#61A6AB'])
        plt.xlim(0, 10)
        plt.xlabel('Score (sur 10)')
        plt.title(question, fontsize=12)
        plt.rc('font', family='Arial')

        for bar, score in zip(bars, [auto_evaluation_score, average_evaluators_score]):
            plt.text(score + 0.1, bar.get_y() + bar.get_height()/2, f'{score:.1f}', va='center', fontsize=10,  antialiased=True)

        img_data = BytesIO()
        plt.savefig(img_data, format='png', bbox_inches='tight', dpi=300)
        plt.close()
        img_data.seek(0)

        img = Image.open(img_data)
        img_path = f"./tmp/temp_image_{idx}.png"
        img.save(img_path)

        y_position = margin_top - (chart_counter % charts_per_page) * (chart_height+20) 
        c.drawImage(img_path, 100, y_position - chart_height, width=chart_width, height=chart_height)

        if (chart_counter + 1) % charts_per_page == 0:
            c.showPage()
            add_header_image(c)
            margin_top = 750

        chart_counter += 1

    # Ajouter la dernière page avec les commentaires des répondants

    c.showPage()
    add_header_image(c)
    c.setFont("Arial", 10)
    c.setFillColorRGB(*custom_color2)
    c.drawString(40, 750, "Selon-vous, quelles sont ses forces ?")
    forces = sheet_2.iloc[:, force_col].dropna().tolist()
    y_position = 720
    for comment in forces:
        c.setFillColorRGB(*custom_color1)
        c.drawString(100, y_position, f"- {comment}")
        y_position -= 20
        if y_position < 100:  # Sauter à une nouvelle page si l'espace est insuffisant
            c.showPage()
            add_header_image(c)
            y_position = 780

    c.showPage()
    add_header_image(c)
    c.setFillColorRGB(*custom_color2)
    c.drawString(40, 750, "Selon-vous, quels sont ses points d'amélioration ?")
    y_position = 740
    c.setFont("Arial", 10)
    ameliorations = sheet_2.iloc[:, amelioration_col].dropna().tolist()
    y_position -= 20
    for comment in ameliorations:
        c.setFillColorRGB(*custom_color1)
        c.drawString(100, y_position, f"- {comment}")
        y_position -= 20
        if y_position < 100:
            c.showPage()
            y_position = 780

    # Dernière Page
    c.showPage()
    add_header_image(c)
    c.setFont("Arial", 12)
    c.drawString(100, 750, last_page_text)
	
    # Finaliser et sauvegarder le PDF
    c.save()

# Exemple d'appel
input_excel_path = r"D:\Drive LafontaineConsulting Microsoft\OneDrive - Lafontaine Consulting\Nouveau\xls_to_pdf\fichier_excel.xlsx"  # Utiliser le chemin brut
#output_pdf_path = f"D:\\Drive LafontaineConsulting Microsoft\\OneDrive - Lafontaine Consulting\\Nouveau\\xls_to_pdf\\UP-CARRIERE 360 {prenom} {nom}.pdf"  # Utiliser le chemin brut
first_page_text = "Ce texte est un libellé d'introduction que vous pouvez modifier."
last_page_text = "Ce texte est un libellé de conclusion que vous pouvez modifier."
image_path = r".\util\Logo simple_up_carriere.jpg"

generate_pdf_with_title_and_comments(input_excel_path, first_page_text, last_page_text)
