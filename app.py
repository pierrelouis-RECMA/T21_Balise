import io, os, traceback, uuid
from datetime import datetime, timedelta
import pandas as pd
from flask import Flask, abort, jsonify, request, send_file

# Importation du moteur de génération
from generate_pptx_v3 import build_agency_pptx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB

TEMPLATE_NAME = "T21_HK_Agencies_Glass_v13.pptx"
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), TEMPLATE_NAME)

_cache = {}

@app.route("/")
def index():
    return "Serveur NBB Opérationnel. Envoyez un POST sur /generate avec un fichier Excel."

@app.route("/generate", methods=["POST"])
def generate():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file uploaded"}), 400
        
        file = request.files["file"]
        df = pd.read_excel(file)

        # Appel du moteur V3
        pptx_data = build_agency_pptx(df, TEMPLATE_PATH)

        token = str(uuid.uuid4())
        _cache[token] = {
            "data": pptx_data,
            "filename": f"NBB_Report_{datetime.now().strftime('%Y%m%d')}.pptx",
            "expiry": datetime.now() + timedelta(minutes=10)
        }

        return jsonify({
            "status": "success",
            "download_url": f"{request.host_url.rstrip('/')}/download/{token}"
        })
    except Exception as e:
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500

@app.route("/download/<token>")
def download(token):
    entry = _cache.get(token)
    if not entry or datetime.now() > entry["expiry"]:
        abort(404, "Lien expiré ou invalide.")
    
    return send_file(
        io.BytesIO(entry["data"]),
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        as_attachment=True,
        download_name=entry["filename"]
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
