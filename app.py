"""
Nutrylife PPTX Generator
Servidor Flask que recibe parámetros del formulario
y genera el plan nutricional en formato PPTX.
"""

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import tempfile
import sys

from pptx_generator import generar_pptx_completo

app = Flask(__name__)
CORS(app)

@app.route('/')
def home():
    return jsonify({
        'status': 'ok',
        'mensaje': 'Nutrylife PPTX Generator funcionando',
        'version': '2.0'
    })

@app.route('/generar', methods=['POST'])
def generar():
    try:
        datos = request.get_json()

        if not datos:
            return jsonify({'error': 'No se recibieron datos'}), 400

        if not datos.get('nombre'):
            return jsonify({'error': 'Falta el nombre del paciente'}), 400

        if not datos.get('kcal'):
            return jsonify({'error': 'Faltan las calorias'}), 400

        # Generar PPTX en archivo temporal
        nombre_limpio = datos['nombre'].replace(' ', '_')
        with tempfile.NamedTemporaryFile(
            suffix='.pptx',
            delete=False,
            prefix=f"plan_{nombre_limpio}_"
        ) as tmp:
            output_path = tmp.name

        generar_pptx_completo(datos, output_path)

        nombre_archivo = f"Plan_{nombre_limpio}.pptx"

        return send_file(
            output_path,
            as_attachment=True,
            download_name=nombre_archivo,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        print(f"Error generando PPTX: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr)
        return jsonify({'error': str(e)}), 500


@app.route('/ping', methods=['GET'])
def ping():
    return 'pong', 200


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
