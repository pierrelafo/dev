import subprocess

# Remplacez 'script.py' par le chemin du script Streamlit principal
streamlit_script = "360.py"

# Commande pour exécuter le script Streamlit
command = f"streamlit run {streamlit_script}"


try:
    # Exécuter Streamlit
    subprocess.run(command, shell=True, check=True)
except subprocess.CalledProcessError as e:
    print(f"Erreur lors de l'exécution de Streamlit : {e}")
    input("Appuyez sur Entrée pour fermer...")  # Pause pour voir l'erreur