from flask import Flask, send_from_directory, jsonify, render_template, request, send_file
from prétraitement import prétraitement
from post_traitement import clean_import_sheets  # Assure-toi que ce fichier s'appelle bien ainsi
import os
import tempfile
import uuid

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/generate_excel", methods=["POST"])
def generate_excel():
    tabs = request.get_json()
    if not tabs:
        return jsonify({"message": "Aucun onglet fourni"}), 400
    
    # Prétraitement et création du fichier Excel de base
    prétraitement(tabs)
    
    return jsonify({"message": "Fichier généré", "url": "/download_excel"})

@app.route("/download_excel")
def download_excel():
    # Sert le fichier Excel généré par le prétraitement
    return send_from_directory(os.getcwd(), "mon_fichier.xlsx", as_attachment=True)

@app.route("/upload_excel", methods=["POST"])
def upload_excel():
    print("✅ Fichier reçu ? =>", "file" in request.files)

    if "file" not in request.files:
        return "Aucun fichier reçu", 400

    file = request.files["file"]
    print("📄 Nom du fichier :", file.filename)

    if file.filename == "":
        return "Nom de fichier vide", 400

    unique_name = f"utilisateur_{uuid.uuid4().hex}.xlsx"
    temp_input_path = os.path.join(tempfile.gettempdir(), unique_name)
    output_name = f"nettoye_{uuid.uuid4().hex}.xlsx"
    temp_output_path = os.path.join(tempfile.gettempdir(), output_name)

    file.save(temp_input_path)
    print("💾 Fichier enregistré à :", temp_input_path)

    clean_import_sheets(temp_input_path, temp_output_path)

    if not os.path.exists(temp_output_path):
        print("❌ Fichier nettoyé non trouvé")
        return "Le fichier de sortie n'a pas été généré.", 500

    print("✅ Envoi du fichier nettoyé")
    return send_file(temp_output_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)