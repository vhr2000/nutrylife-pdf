"""
Nutrylife PDF Generator
Servidor Flask que recibe parámetros del formulario
y genera el PDF personalizado usando el script de Myriam.
"""

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import json
import os
import tempfile
import sys

# Importar el generador de PDF
from pdf_generator import generar_pdf_completo

app = Flask(__name__)
CORS(app)  # Permite que el formulario HTML se conecte desde cualquier dominio

@app.route('/')
def home():
    return jsonify({
        'status': 'ok',
        'mensaje': 'Nutrylife PDF Generator funcionando',
        'version': '1.0'
    })

@app.route('/generar', methods=['POST'])
def generar():
    """
    Recibe los parámetros del formulario y devuelve el PDF.
    """
    try:
        datos = request.get_json()

        if not datos:
            return jsonify({'error': 'No se recibieron datos'}), 400

        # Validar campos mínimos
        if not datos.get('nombre'):
            return jsonify({'error': 'Falta el nombre del paciente'}), 400

        if not datos.get('kcal'):
            return jsonify({'error': 'Faltan las calorías'}), 400

        # Generar PDF en un archivo temporal
        with tempfile.NamedTemporaryFile(
            suffix='.pdf',
            delete=False,
            prefix=f"plan_{datos['nombre'].replace(' ', '_')}_"
        ) as tmp:
            output_path = tmp.name

        generar_pdf_completo(datos, output_path)

        # Nombre del archivo para la descarga
        nombre_archivo = f"Plan_{datos['nombre'].replace(' ', '_')}.pdf"

        return send_file(
            output_path,
            as_attachment=True,
            download_name=nombre_archivo,
            mimetype='application/pdf'
        )

    except Exception as e:
        print(f"Error generando PDF: {e}", file=sys.stderr)
        return jsonify({'error': str(e)}), 500


@app.route('/ping', methods=['GET'])
def ping():
    """Health check para Render."""
    return 'pong', 200


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
