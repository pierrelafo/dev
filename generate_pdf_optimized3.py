import pandas as pd
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib import colors  # Pour des couleurs prédéfinies
import numpy as np
import os
from PIL import Image
import textwrap


# Enregistrer la police Arial
#arial_font_path = r'C:\Windows\Fonts\Arial.ttf'
#pdfmetrics.registerFont(TTFont('Arial', arial_font_path))

# Fonction pour convertir un code hexadécimal en RGB
def hex_to_rgb(hex_color):
	hex_color = hex_color.lstrip('#')
	return tuple(int(hex_color[i:i+2], 16) / 255.0 for i in (0, 2, 4))

def generate_pdf_with_title_and_comments(input_excel_path, first_page_text, image_path,text_encadre, pdf_name):
	def add_header_image(c):
		c.drawImage(image_path, (page_width - 102) / 2, page_height - 50, width=102, height=40)
	
	# Charger les données Excel
	sheet_1 = pd.read_excel(input_excel_path, sheet_name=0, header=2)
	sheet_2 = pd.read_excel(input_excel_path, sheet_name=1, header=2)
	categories = pd.read_excel(input_excel_path, sheet_name=0, header=None).iloc[0]
	categories_2 = pd.read_excel(input_excel_path, sheet_name=1, header=None).iloc[0]
	questions_sheet_1 = pd.read_excel(input_excel_path, sheet_name=0, header=None).iloc[1]

	# Récupérer le prénom et le nom de la personne
	prenom = pd.read_excel(input_excel_path, sheet_name=0).iloc[2, 3]
	nom = pd.read_excel(input_excel_path, sheet_name=0).iloc[2, 2]

	# Trouver les indices des colonnes pour les questions
	start_col = questions_sheet_1[questions_sheet_1.str.contains("Je suis à l'aise dans l'échange", case=False, na=False)].index[0]
	end_col = questions_sheet_1[questions_sheet_1.str.contains("Je prends des initiatives", case=False, na=False)].index[0]

	output_pdf_path = f"{pdf_name}"
	
	# Initialiser le canvas
	c = canvas.Canvas(output_pdf_path, pagesize=A4)
	page_width, page_height = A4
	add_header_image(c)

	#	#	#	#	#	#	#	#   Debut Titre	#	#	#	#	#	#	#	#	#	#
	# Ajouter le titre
	c.setFillColorRGB(*custom_color2)
	#c.setFont("Arial", 20)
	#c.drawCentredString(page_width/2, 680, f"360° de {prenom} {nom}")
	# Initialiser les styles pour le titre de la section PDF
	styles = getSampleStyleSheet()
	style = styles["BodyText"]
	style.alignment = TA_CENTER
	style.fontSize = 20
	style.textColor = colors.HexColor("#365B6D")
	
	# Titre pour la section agrégée
	text_open = f"<b>360° de {prenom} {nom}</b>"
	paragraph = Paragraph(text_open, style)
	paragraph_width, paragraph_height = paragraph.wrap(490, 0)
	paragraph.drawOn(c, 50, 680)                                                         
                               
                           
                            
                    
                                             
 
                                  
                                              
                                        
                                                           
                             
	
	#	#	#	#	#	#	#	#   Debut Pragraphe	#	#	#	#	#	#	#	#	#	#
	#Ajouter le paragraphe de 1ere page
	# Créer un style de paragraphe
	styles = getSampleStyleSheet()
	style = styles["BodyText"]  # Utiliser le style BodyText, ou ajustez comme vous le souhaitez
	style.fontSize = 12  # Taille de la police
	style.leading = 14  # Interligne (ajustez si nécessaire)
	style.textColor = colors.HexColor("#61A6AB") 
	
	# Créer un paragraphe avec le texte et le style
	paragraph = Paragraph(first_page_text, style)

	# Définir la largeur et la hauteur du bloc de texte
	paragraph_width = 490  # Largeur maximale en points
	paragraph_height = 200  # Hauteur maximale en points (ajustez si nécessaire)

	# Dessiner le paragraphe sur le canvas
	paragraph.wrap(paragraph_width, paragraph_height)  # Définir les limites de largeur/hauteur
	paragraph.drawOn(c, 50, 250)  # Positionner le paragraphe sur le canvas



	#	#	#	#	#	#	#	#   Debut Texte encadre	#	#	#	#	#	#	#	#	#	#

	# Créer un style de paragraphe pour le texte
	styles = getSampleStyleSheet()
	style = styles["BodyText"]
	style.fontSize = 8
	style.leading = 10
	style.textColor = colors.red  # Couleur du texte
	
	# Créer un paragraphe avec le texte et le style
	paragraph = Paragraph(text_encadre, style)
	
	# Définir la largeur et la hauteur du bloc de texte
	paragraph_width = 490  # Largeur de l'encadré
	paragraph_height = 30  # Hauteur approximative de l'encadré (ajustez si nécessaire)
	
	# Positionner le texte sur le canvas et obtenir les dimensions finales
	x, y = 50, 80  # Coordonnées du coin inférieur gauche de l'encadré
	paragraph.wrap(paragraph_width, paragraph_height)
	paragraph.drawOn(c, x + 5, y + 10)  # Décaler légèrement le texte pour qu'il soit à l'intérieur du cadre
	
	# Dessiner le cadre rouge autour du texte
	c.setStrokeColor(colors.red)  # Couleur de la bordure
	c.setLineWidth(0.5)  # Épaisseur de la bordure
	c.rect(x, y, paragraph_width + 10, paragraph_height + 10)  # Dessiner le rectangle autour du texte
	
	#	#	#	#	#	#	#	#   Début Graphique aggrégé	#	#	#	#	#	#	#	#   
	c.showPage()
	add_header_image(c)
	y_position_init = 760  # Position verticale de départ pour chaque page
	image_height = 250  # Hauteur approximative pour chaque graphique
	previous_category = None
	section_data = {}
	questions_to_plot = questions_sheet_1[start_col:end_col+1]
	# Initialiser les styles pour le titre de la section PDF
	styles = getSampleStyleSheet()
	style = styles["BodyText"]
	style.alignment = TA_CENTER
	style.fontSize = 14
	style.textColor = colors.HexColor("#365B6D")
	
	# Titre pour la section agrégée
	text_open = "<b>Représentation graphique agrégée pour toutes les catégories</b>"
	paragraph = Paragraph(text_open, style)
	paragraph_width, paragraph_height = paragraph.wrap(490, 0)
	y_position = y_position_init - paragraph_height
	paragraph.drawOn(c, 50, y_position)
	y_position -= 20
	
	# Initialiser les données agrégées
	aggregated_data = {}
	previous_category = None
	section_data = {}
	
	# Parcourir les questions et regrouper par catégorie pour calculer les moyennes
	for idx, question in questions_to_plot.items():
		current_category = categories[idx]
	
		if pd.notna(current_category) and current_category.strip() != "":
			# Si la catégorie change, calculer les moyennes et stocker dans aggregated_data
			if current_category != previous_category and section_data:
				# Calculer les moyennes pour chaque catégorie
				avg_eval = np.mean([score[0] for score in section_data.values()])
				avg_auto_eval = np.mean([score[1] for score in section_data.values()])
				aggregated_data[previous_category] = (avg_eval,avg_auto_eval)
				section_data = {}
	
			previous_category = current_category
	
		# Obtenir les scores de chaque question
                                                                              
		evaluators_scores = pd.to_numeric(sheet_2.iloc[:, idx+2], errors='coerce')
		average_evaluators_score = evaluators_scores.mean()
		auto_evaluation_score = pd.to_numeric(sheet_1.iloc[0, idx], errors='coerce')
	
		# Ajouter les scores à la catégorie si les données sont valides
		if not pd.isna(auto_evaluation_score) and not pd.isna(average_evaluators_score):
			section_data[question] = (average_evaluators_score,auto_evaluation_score)
	
	# Calculer et stocker la dernière section restante
	if section_data:
		avg_eval = np.mean([score[0] for score in section_data.values()])
		avg_auto_eval = np.mean([score[1] for score in section_data.values()])
		aggregated_data[previous_category] = (avg_eval,avg_auto_eval)
	
	# Inverser l'ordre des catégories dans aggregated_data
	aggregated_data_reversed = {key: aggregated_data[key] for key in reversed(aggregated_data)}

	# Générer le graphique avec les catégories inversées
	y_position = generate_overall_bar_chart(aggregated_data_reversed, c, page_width, y_position, image_height, custom_color2, custom_color1, add_header_image, y_position_init)

	
	#	#	#	#	#	#	#	#   Début Graphique par catégorie	#	#	#	#	#	#	#	#   
	c.showPage()
	add_header_image(c)

	# Initialisation des variables
	previous_category = None
	section_data = {}
	questions_to_plot = questions_sheet_1[start_col:end_col+1]

	# Titre Détail représentation graphique
	styles = getSampleStyleSheet()
	style = styles["BodyText"]  # Utiliser le style BodyText, ou ajustez comme vous le souhaitez
	style.alignment = TA_CENTER
	style.fontSize = 14 
	style.textColor = colors.HexColor("#365B6D") 
	paragraph_width = 490  
	text_open = "<b>Représentation graphique détaillée pour chaque thématique</b>"
	paragraph = Paragraph(text_open, style)
	paragraph_width, paragraph_height = paragraph.wrap(paragraph_width, 0)
	y_position = y_position_init-paragraph_height
	paragraph.drawOn(c, 50, y_position)
	y_position -= 20
	
	# Parcourir toutes les questions et regrouper par section
	for idx, question in questions_to_plot.items():
		current_category = categories[idx]

		if pd.notna(current_category) and current_category.strip() != "":
			# Si la catégorie change, générer le graphique pour la section précédente
			if current_category != previous_category and section_data:
				# Générer le graphique pour la section précédente avec toutes ses questions
				y_position = generate_grouped_bar_chart(section_data, previous_category, c, page_width, y_position, image_height, custom_color1, custom_color2, add_header_image,y_position_init)
				section_data = {}  # Réinitialiser pour la nouvelle section

			# Mettre à jour la catégorie actuelle
			previous_category = current_category

			# Obtenir les scores de chaque question
		auto_evaluation_score = pd.to_numeric(sheet_1.iloc[0, idx], errors='coerce')
		evaluators_scores = pd.to_numeric(sheet_2.iloc[:, idx+2], errors='coerce')
		average_evaluators_score = evaluators_scores.mean()

		# Ajouter les scores si les données sont valides
		if not pd.isna(auto_evaluation_score) and not pd.isna(average_evaluators_score):
			section_data[question] = (auto_evaluation_score, average_evaluators_score)

	# Générer un graphique pour la dernière section restante
	if section_data:
		generate_grouped_bar_chart(section_data, previous_category, c, page_width, y_position, image_height, custom_color1, custom_color2, add_header_image,y_position_init)
	#	#  Les compétences	#	#	
	c.showPage()
	add_header_image(c)
	y_position = y_position_init
	styles = getSampleStyleSheet()
	style = styles["BodyText"]  # Utiliser le style BodyText, ou ajustez comme vous le souhaitez
	style.alignment = TA_CENTER
	style.fontSize = 14 
	style.textColor = colors.HexColor("#365B6D") 
	paragraph_width = 490  
	text_open = "<b>Analyse qualitative - Les compétences</b><br/><br/>"
	paragraph = Paragraph(text_open, style)
	paragraph_width, paragraph_height = paragraph.wrap(paragraph_width, 0)
	paragraph.drawOn(c, 50, y_position_init-paragraph_height)  
	y_position=y_position_init-paragraph_height-20
			
	#	#  Générer TOP 5	#	#	
	titre='Top 5 des compétences reconnues par votre environnement'
	type='max'
	y_position=generate_top_5_competences(input_excel_path,page_width,y_position,c,y_position_init,add_header_image,custom_color2,titre,type)
 	#	#  Générer TOP 5 sous-evaluation   #	#	
	titre='Top 5 des compétences où je me sous-evalue le plus'
	type='sous'
	y_position=generate_top_5_evaluation(input_excel_path,page_width,y_position,c,y_position_init,add_header_image,custom_color1,custom_color2,titre,type)
	# Sauvegarder le PDF
	#	#  Les Améliorations	#	#	
	c.showPage()
	add_header_image(c)
	y_position = y_position_init
	styles = getSampleStyleSheet()
	style = styles["BodyText"]  # Utiliser le style BodyText, ou ajustez comme vous le souhaitez
	style.alignment = TA_CENTER
	style.fontSize = 14 
	style.textColor = colors.HexColor("#365B6D") 
	paragraph_width = 490  
	text_open = "<b>Analyse qualitative - Les améliorations</b><br/><br/>"
	paragraph = Paragraph(text_open, style)
	paragraph_width, paragraph_height = paragraph.wrap(paragraph_width, 0)
	paragraph.drawOn(c, 50, y_position_init-paragraph_height)  
	y_position=y_position_init-paragraph_height-20
			
	#	#  Générer TOP 5 améliorations   #	#	
	titre='Top 5 des axes d’amélioration suggérés par mon environnement'
	type='min'
	y_position=generate_top_5_competences(input_excel_path,page_width,y_position,c,y_position_init,add_header_image,custom_color2,titre,type)
	# Sauvegarder le PDF

	#	#  Générer TOP 5 sur-evaluation   #	#	
	titre='Top 5 des compétences où je me sur-evalue le plus'
	type='sur'
	y_position=generate_top_5_evaluation(input_excel_path,page_width,y_position,c,y_position_init,add_header_image,custom_color1,custom_color2,titre,type)

	
	#	#	#	#	#	#	#	#   Debut réponses ouvertes	#	#	#	#	#	#	#	#	#	#
	#force_col = 44  # Colonne AS
	#amelioration_col = 45  # Colonne AT
	#evolution_col = 56  # Colonne AU
	c.showPage()
	add_header_image(c)
	
	# Titre question ouvertes
	styles = getSampleStyleSheet()
	style = styles["BodyText"]  # Utiliser le style BodyText, ou ajustez comme vous le souhaitez
	style.alignment = TA_CENTER
	style.fontSize = 14 
	style.textColor = colors.HexColor("#365B6D") 
	paragraph_width = 490  
	text_open = "<b>Analyse qualitative - Questions ouvertes</b><br/><br/>"
	paragraph = Paragraph(text_open, style)
	paragraph_width, paragraph_height = paragraph.wrap(paragraph_width, 0)
	paragraph.drawOn(c, 50, y_position_init-paragraph_height)  
	y_position=y_position_init-paragraph_height-20

	# Paragraphe explication questions ouvertes
	style.alignment = TA_LEFT
	style.fontSize = 10  
	style.textColor = colors.HexColor("#61A6AB") 
	paragraph_width = 490  
	text_open = "<i>Les commentaires repris dans cette section ont été écrits tels quels par les personnes sollicitées. Ils n'ont subi aucune modification, n'ont pas été triés ou classés d'une quelconque manière. " \
				"Si certains commentaires apparaissent plusieurs fois, cela signifie que plusieurs personnes ont fait le même commentaire.</i><br/><br/><br/>"
	paragraph = Paragraph(text_open, style)
	paragraph_width, paragraph_height = paragraph.wrap(paragraph_width, 0)
	paragraph.drawOn(c, 50, y_position-paragraph_height)
	y_position -= (paragraph_height+10)

	for idx_col in range(42,51):
		y_position=generate_question_ouverte(c,idx_col,y_position,y_position_init,add_header_image,sheet_2,sheet_1,categories_2)
                                              

	# Sauvegarder le PDF	
	#Fin
	c.save()

def generate_question_ouverte(c,idx_col,y_position,y_position_init,add_header_image,sheet_2,sheet_1,categories_2):
	if y_position <180:
		c.showPage()
		add_header_image(c)
		y_position = y_position_init
	styles = getSampleStyleSheet()
	style = styles["BodyText"]  # Utiliser le style BodyText, ou ajustez comme vous le souhaitez
	style.fontSize = 12  
	style.textColor = colors.HexColor("#365B6D") 
	paragraph_width = 490  
	text_open = "<b>"+categories_2[idx_col+2]+"</b>"
	paragraph = Paragraph(text_open, style)
	paragraph_width, paragraph_height = paragraph.wrap(paragraph_width, 0)
	paragraph.drawOn(c, 50, y_position-paragraph_height) 
	y_position -= (paragraph_height+10)

	comments = sheet_2.iloc[:, idx_col+2].dropna().tolist()
	style.fontSize = 10 
	style.textColor = colors.HexColor("#61A6AB") 

	for comment in comments:
		paragraph = Paragraph("&nbsp;&nbsp;&nbsp;&nbsp;- " + str(comment), style)
		paragraph_width, paragraph_height = paragraph.wrap(paragraph_width, 0)
		if y_position - (paragraph_height+5)<50:
			c.showPage()
			add_header_image(c)
			y_position = y_position_init
		paragraph.drawOn(c, 50, y_position-paragraph_height)  
		y_position -= (paragraph_height+5)
	if idx_col != 50 : 
		paragraph = Paragraph("<b>Mon verbatim : </b>"+ str(sheet_1.iloc[0,idx_col]), style)
		paragraph_width, paragraph_height = paragraph.wrap(paragraph_width, 0)
		paragraph.drawOn(c, 50, y_position-paragraph_height) 
		y_position -= (paragraph_height+5)

	return y_position-10

def generate_overall_bar_chart(aggregated_data, c, page_width, y_position, image_height, custom_color1, custom_color2, add_header_image, y_position_init):
	# Extraire les scores moyens pour chaque catégorie
	categories = list(aggregated_data.keys())
	auto_evaluation_scores = [data[0] for data in aggregated_data.values()]
	evaluator_scores = [data[1] for data in aggregated_data.values()]

	# Configurer les barres pour chaque catégorie
	bar_width = 0.35
	spacing = 1.5
	indices = np.array([i * spacing for i in range(len(categories))]) 

	fig, ax = plt.subplots(figsize=(10, 6))
	fig.patch.set_facecolor('#F2F1EC')
	ax.grid(axis='x', color='#F2F1EC', linestyle='-', linewidth=0.7, zorder=1)

	# Tracer les barres horizontales pour chaque catégorie
	ax.barh(indices + bar_width, auto_evaluation_scores, height=bar_width, label="Moyenne évaluateurs", color=custom_color1, zorder=2)
	ax.barh(indices , evaluator_scores, height=bar_width, label="Mon évaluation", color=custom_color2, zorder=2)

	# Ajouter les scores en bout de chaque barre
	for i, (score_auto, score_eval) in enumerate(zip(auto_evaluation_scores, evaluator_scores)):
		ax.text(score_eval + 0.1, indices[i], f"{score_eval:.1f}", va='center', ha='left', fontsize=8, color=custom_color2)
		ax.text(score_auto + 0.1, indices[i] + bar_width, f"{score_auto:.1f}", va='center', ha='left', fontsize=8, color=custom_color1)


	# Personnaliser le graphique
	ax.set_yticks(indices + bar_width / 2)
	ax.set_yticklabels(categories, rotation=0, ha='right', fontsize=8)
	ax.set_title("Comparaison agrégée des scores par catégorie", fontsize=12)
	ax.set_xlabel("Score moyen")
	ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.15), ncol=2)
	ax.set_xlim(0.9, 5)
	plt.tight_layout()
	plt.xticks(np.arange(1, 6, 1))

	# Sauvegarder le graphique et l'ajouter au PDF
	img_path = "./tmp/overall_horizontal_chart.png"
	plt.savefig(img_path, format='png', bbox_inches='tight', dpi=300)
	plt.close()
	y_position = image_pdf(img_path, page_width, y_position, c, y_position_init, add_header_image)
	return y_position

	
def generate_top_5_evaluation(input_excel_path,page_width,y_position,c,y_position_init,add_header_image,custom_color1,custom_color2,titre,type):	
	#	#  Générer TOP 5 améliorations   #	#	
	 # Charger les deux onglets : celui de la personne évaluée et celui des évaluateurs
	person_df = pd.read_excel(input_excel_path, sheet_name=0, header=1)  # Onglet de la personne évaluée
	evaluators_df = pd.read_excel(input_excel_path, sheet_name=1, header=1)  # Onglet des évaluateurs

	# Sélectionner les colonnes contenant les réponses de la personne évaluée et des évaluateurs avec décalage
	person_scores = pd.to_numeric(person_df.iloc[1, 4:42], errors='coerce')  # Colonnes G à AR pour les réponses de la personne (index 4 à 41)
	questions = person_df.columns[4:42]  # Questions en en-tête sur la feuille de la personne

	# Sélectionner les colonnes décalées pour les évaluateurs (deux colonnes supplémentaires)			
	evaluator_scores = evaluators_df.iloc[:, 6:44].apply(pd.to_numeric, errors='coerce')  # Colonnes I à AT (index 6 à 43)

	# Calculer la moyenne des réponses des évaluateurs pour chaque question
	evaluator_means = evaluator_scores.mean()
	# Calculer les écarts entre la personne et la moyenne des évaluateurs
	gaps = evaluator_means - person_scores if type == 'sous' else  person_scores - evaluator_means
	# Calculer les 5 questions avec les scores les plus faibles de la personne évaluée (axes d'amélioration)
	top_5_gaps = gaps.nlargest(5)
	top_5_indices = top_5_gaps.index
	
	top_5_person_scores = person_scores[top_5_indices]
	top_5_evaluator_means = evaluator_means[top_5_indices]
	# Convertir les indices en entiers
	top_5_indices = [questions.get_loc(i) for i in top_5_indices]  # Récupère les positions des indices dans questions

	# Maintenant, cela devrait fonctionner pour sélectionner les questions correspondantes
	top_5_questions = [questions[i] for i in top_5_indices]

	# Appliquer un retour à la ligne automatique pour les labels des questions
	wrapped_labels = [textwrap.fill(label, width=45) for label in top_5_questions]

	# Créer un graphique pour comparer les scores de la personne évaluée et la moyenne des évaluateurs
	bar_width = 0.35
	spacing = 1.5
	indices = [i * spacing for i in range(len(top_5_questions))]
	fig, ax = plt.subplots(figsize=(8, 5 * 0.70))
	fig.patch.set_facecolor('#F2F1EC')
	ax.grid(axis='x', color='#F2F1EC', linestyle='-', linewidth=0.7, zorder=1)
	ax.barh(indices, top_5_evaluator_means, height=bar_width, label="Moyenne des évaluateurs" ,color=custom_color2, zorder=2)
	ax.barh([i + bar_width for i in indices], top_5_person_scores, height=bar_width, label="Mon évaluation", color=custom_color1, zorder=2)
	# Ajouter le score en bout de chaque barre
	for i, (score_eval, score_person) in enumerate(zip(top_5_evaluator_means, top_5_person_scores)):
		ax.text(score_eval + 0.1, indices[i], f"{score_eval:.1f}", va='center', ha='left', fontsize=8, color=custom_color2)
		ax.text(score_person + 0.1, indices[i] + bar_width+ 0.04, f"{score_person:.1f}", va='center', ha='left', fontsize=8, color=custom_color1)

	#ax.barh(indices, top_5_questions, height=bar_width, color=color2, zorder=2) 
	ax.set_yticks(indices)
	ax.set_yticklabels(top_5_questions)
	ax.set_xlim(0.9, 5)
	ax.set_xlabel("Score (sur 5)")
	ax.set_title(f"{titre}")
	ax.grid(axis='x', color='#F2F1EC', linestyle='-', linewidth=0.7, zorder=1)
	ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.2), ncol=2, fontsize=8)
	
	plt.yticks(indices, wrapped_labels,fontsize=8)
	plt.title(titre,fontsize=10)
	plt.xticks(np.arange(1, 6, 1))
	plt.xlabel('Score')
	plt.gca().invert_yaxis()  # Inverser l'ordre pour afficher la meilleure moyenne en haut
	plt.tight_layout()

	img_path = f"./tmp/section_top_5_evaluation_{type}.png"
	plt.savefig(img_path, format='png', bbox_inches='tight', dpi=300)
	plt.close()
	y_position=image_pdf(img_path,page_width,y_position,c,y_position_init,add_header_image)
	return y_position
	
def generate_top_5_competences(input_excel_path,page_width,y_position,c,y_position_init,add_header_image,color2,titre,type):
	df = pd.read_excel(input_excel_path, sheet_name=1, header=1)
	questions_df = df.iloc[:, 6:44]  # Colonnes G à AR (index 6 à 43)
	questions_df = questions_df.apply(pd.to_numeric, errors='coerce')
	mean_scores = questions_df.mean()
	top_5_questions = mean_scores.nlargest(5) if type == 'max' else mean_scores.nsmallest(5)
	# Appliquer le retour à la ligne automatique aux étiquettes des questions
	wrapped_labels = [textwrap.fill(label, width=45) for label in top_5_questions.index]
	# Définir les paramètres de hauteur et d'espacement
	bar_width = 0.45
	spacing = 1.5
	indices = [i * spacing for i in range(len(top_5_questions))]
	fig, ax = plt.subplots(figsize=(8, 5 * 0.70))
	fig.patch.set_facecolor('#F2F1EC')
	ax.set_xlim(0.9, 5)
	ax.grid(axis='x', color='#F2F1EC', linestyle='-', linewidth=0.7, zorder=1)
	ax.barh(indices, top_5_questions, height=bar_width, color=color2, zorder=2) 
	plt.yticks(indices, wrapped_labels,fontsize=8)
	plt.title(titre,fontsize=10)
	plt.xticks(np.arange(1, 6, 1))
	plt.xlabel('Score')
	plt.gca().invert_yaxis()  # Inverser l'ordre pour afficher la meilleure moyenne en haut
	plt.tight_layout()  

	img_path = f"./tmp/section_top_5_{type}.png"
	plt.savefig(img_path, format='png', bbox_inches='tight', dpi=300)
	plt.close()
	y_position=image_pdf(img_path,page_width,y_position,c,y_position_init,add_header_image)
	return y_position

def generate_grouped_bar_chart(section_data, section_title, c, page_width, y_position, image_height, color1, color2, add_header_image,y_position_init):
	# Récupérer les questions et les scores
	max_label_width = 50  # Nombre maximal de caractères par ligne (ajustez si nécessaire)
	questions = [textwrap.fill(question, max_label_width) for question in section_data.keys()]
	auto_scores = [score[0] for score in section_data.values()]
	evaluator_scores = [score[1] for score in section_data.values()]
	
	fig, ax = plt.subplots(figsize=(8, len(questions) * 0.70))
	fig.patch.set_facecolor('#F2F1EC')
	
	# Indices pour les barres
	bar_width = 0.35
	spacing = 1.5  # Espacement supplémentaire entre chaque question
	indices = [i * spacing for i in range(len(questions))]

	# Création des barres groupées avec une boucle
	for i, question in enumerate(questions):
		ax.barh(indices[i] + bar_width / 2, evaluator_scores[i], height=bar_width, color=color2, label="Moyenne des évaluateurs" if i == 0 else "", zorder=2)
		ax.barh(indices[i] - bar_width / 2, auto_scores[i], height=bar_width, color=color1, label="Mon évaluation" if i == 0 else "", zorder=2)
		# Ajouter la valeur au bout de chaque barre pour l'auto-évaluation
		ax.text(auto_scores[i] + 0.1, indices[i] - bar_width / 2 - 0.04, f'{auto_scores[i]:.1f}', va='center', color=color1, fontsize=8)
		ax.text(evaluator_scores[i] + 0.1, indices[i] + bar_width / 2 , f'{evaluator_scores[i]:.1f}', va='center', color=color2, fontsize=8)


	# Définir les étiquettes et le titre
	ax.set_yticks(indices)
	ax.set_yticklabels(questions)
	ax.set_xlim(0.9, 5)
	ax.set_xlabel("Score (sur 5)")
	ax.set_title(f"{section_title}")
	ax.grid(axis='x', color='#F2F1EC', linestyle='-', linewidth=0.7, zorder=1)
	#ax.legend(loc="center left", bbox_to_anchor=(1, 0.5))
	legend_y_position = -0.3 if len(questions) <= 3 else -0.15
	ax.legend(loc="upper center", bbox_to_anchor=(0.5, legend_y_position), ncol=2)
	plt.xticks(np.arange(1, 6, 1))

	# Sauvegarder chaque graphique dans un fichier image unique pour chaque section
	img_path = f"./tmp/section_{section_title.replace(' ', '_').replace('/', '').replace('?', '').replace('@', '')[:10]}.png"
	plt.savefig(img_path, format='png', bbox_inches='tight', dpi=300)
	plt.close()
	y_position=image_pdf(img_path,page_width,y_position,c,y_position_init,add_header_image)
	return y_position

def image_pdf(img_path,page_width,y_position,c,y_position_init,add_header_image):
	# Ouvrir l'image pour récupérer ses dimensions
	with Image.open(img_path) as img:
		img_width, img_height = img.size

	scale_factor = (page_width-50) / img_width
	display_width = img_width * scale_factor
	display_height = img_height * scale_factor
	
	# Vérifier si l'image tient sur la page, sinon ajouter un saut de page
	if y_position - display_height < 60:  # Si l'image ne tient pas, changer de page
		c.showPage()
		add_header_image(c)  # Ajouter l'image d'en-tête sur la nouvelle page
		y_position = y_position_init  # Réinitialiser la position verticale pour la nouvelle page

	# Ajouter l'image au PDF à la position calculée
	c.drawImage(img_path, (page_width - 545) / 2, y_position - display_height, width=display_width, height=display_height)

	# Mettre à jour la position pour le prochain graphique
	y_position -= display_height + 30  # Espace entre les graphiques
	return y_position

# Exemple d'appel
input_excel_path = r"C:\Users\Lafontaine\OneDrive - Lafontaine Consulting\Nouveau\xls_to_pdf\fichier_excel.xlsx"
first_page_text = "Ce rapport 360° est <b>confidentiel</b> et a été conçu afin de vous fournir une analyse détaillée des informations transmises par les personnes que vous avez sollicitées.<br/><br/>" \
				  "Il est à la fois <b>une photographie et un outil de progression</b>.<br/><br/>" \
				  "Grâce à celui-ci, vous allez pouvoir mettre en parallèle votre perception et vous rendre compte qu’elle peut <b>différer avec celle de votre environnement</b>. Le 360° est une démarche constructive, orientée vers le plus d’objectivité et en toute bienveillance.<br/><br/>" \
				  "Les notes attribuées aux 38 caractéristiques analysées sont restituées sous la forme d’une moyenne. Dans ce rapport, vous trouverez les représentations graphiques de chaque thématique ainsi que vos 10 compétences clés reconnues par votre environnement et 5 que vous pourriez développer davantage.<br/><br/>" \
				  "Ce rapport reflète également les éventuelles divergences de points de vue relatives à chaque caractéristique évaluée. Principaux objectifs du rapport 360° :<br/>" \
				  "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- Prendre conscience de la façon dont votre environnement vous perçoit<br/>" \
				  "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- Mieux connaître vos points forts pour en tirer profit<br/>" \
				  "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- Identifier les éventuels changements à mettre en œuvre<br/>" \
				  "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- Mettre en place un ou plusieurs plans d’action<br/><br/><br/>" \
				  "<u>L’échelle de mesure utilisée est :</u><br/>" \
				  "Tout à fait d’accord : 5<br/>" \
				  "Pas du tout d’accord : 1"
text_encadre = "Attention ! Il est important de garder à l'esprit que les personnes sollicitées n'ont pas été formées à l’évaluation. C'est pourquoi, il est primordial de garder une certaine distance et de se concentrer sur les tendances et les points qui reviennent dans les données collectées."

image_path = r".\util\Logo simple_up_carriere.jpg"
	
# Couleurs personnalisées
custom_color1 = hex_to_rgb("#61A6AB")
custom_color2 = hex_to_rgb("#365B6D")

#generate_pdf_with_title_and_comments(input_excel_path, first_page_text, image_path,text_encadre)

                                                                                                           
