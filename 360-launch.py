import subprocess
import sys
import os

# Remplacez 'script.py' par le chemin du script Streamlit principal
streamlit_script = "launch_streamlit.py"

# Commande pour exécuter le script Streamlit
command = f"streamlit run {streamlit_script}"

# Exécuter Streamlit avec subprocess
subprocess.run(command, shell=True)