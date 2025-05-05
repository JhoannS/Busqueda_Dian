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
    return render_template('index.html')

@app.route('/procesar', methods=['POST'])
def procesar():
    if request.method == 'POST':
        # Obtenemos el archivo subido
        file = request.files['input_file']
        output_file_name = request.form['output_file']

        if file and file.filename.endswith('.xlsx'):
            # Guardamos el archivo subido temporalmente
            input_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(input_path)

            # Llamamos a consultar_nits pasándole el archivo y el nombre del archivo de salida
            path_output = os.path.join(RESULT_FOLDER, output_file_name)
            resultado = consultar_nits(input_path, path_output)
            
            # Enviamos el archivo de salida generado como una descarga
            return send_file(resultado, as_attachment=True)

        else:
            return "Archivo no válido. Solo se permite .xlsx"

if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
