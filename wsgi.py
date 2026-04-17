"""
WSGI bridge for PythonAnywhere.
Upload this file alongside api.py and set it as your WSGI file in the PythonAnywhere Web tab.
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from flask import Flask, request, jsonify, make_response
import json, traceback

# Import all the logic from api.py
from api import (parse_rainco_quote, generate_formatted_pdf,
                 _init_fonts, _init_assets, _extract_multipart)

# Pre-load fonts and assets at startup
_init_fonts()
_init_assets()

application = Flask(__name__)

@application.after_request
def cors(r):
    r.headers['Access-Control-Allow-Origin']  = '*'
    r.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    r.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return r

@application.route('/', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@application.route('/parse', methods=['POST', 'OPTIONS'])
def parse_route():
    if request.method == 'OPTIONS':
        return '', 204
    try:
        f = request.files.get('file')
        if not f:
            return jsonify({'error': 'No file uploaded'}), 400
        data = parse_rainco_quote(f.read())
        if not data['items']:
            return jsonify({'error': 'No items found — is this a RainCo quote PDF?'}), 400
        return jsonify(data)
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@application.route('/generate', methods=['POST', 'OPTIONS'])
def generate_route():
    if request.method == 'OPTIONS':
        return '', 204
    try:
        payload = request.get_json()
        pdf = generate_formatted_pdf(
            payload['quote'], payload['rooms'], payload['roomQtys'])
        r = make_response(pdf)
        r.headers['Content-Type'] = 'application/pdf'
        r.headers['Content-Disposition'] = 'attachment; filename=quote.pdf'
        return r
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500
