from flask import Flask, send_from_directory, jsonify, render_template
from excel_utils import create_excel_file

import os

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/generate_excel")
def generate_excel():
    create_excel_file()
    return jsonify({"message": "Fichier généré", "url": "/download_excel"})

@app.route("/download_excel")
def download_excel():
    return send_from_directory(os.getcwd(), "mon_fichier.xlsx", as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)