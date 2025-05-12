from flask import Flask, send_from_directory, jsonify, render_template, request
from excel_utils import prétraitement
import os

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/generate_excel", methods=["POST"])
def generate_excel():
    tabs = request.get_json()
    if not tabs:
        return jsonify({"message": "Aucun onglet fourni"}), 400
    prétraitement(tabs)
    return jsonify({"message": "Fichier généré", "url": "/download_excel"})

@app.route("/download_excel")
def download_excel():
    return send_from_directory(os.getcwd(), "mon_fichier.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)