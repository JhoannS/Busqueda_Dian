from flask import Flask, render_template, request, send_file
import os
from scraping import consultar_nits
from datetime import datetime

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'resultados'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    print("✔️ Cargando vista index.html")
    return render_template('index.html')

@app.route('/procesar', methods=['POST'])
def procesar():
    file = request.files['input_file']
    output_file_name = f'resultado_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'

    if file and file.filename.endswith('.xlsx'):
        input_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(input_path)

        output_path = os.path.join(RESULT_FOLDER, output_file_name)
        result_file = consultar_nits(input_path, output_path)

        return send_file(result_file, as_attachment=True)

    return "Archivo no válido. Solo se permite .xlsx"

if __name__ == "__main__":
    app.run(debug=True, port=3000)
