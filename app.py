import io, json, os, threading, traceback, uuid
from datetime import datetime, timedelta
import pandas as pd
from flask import Flask, abort, jsonify, render_template_string, request, send_file
from generate_pptx_v3 import build_agency_pptx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

TEMPLATE_NAME = "T21_HK_Agencies_Glass_v13.pptx"
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), TEMPLATE_NAME)

_cache = {}

# --- COPIEZ LE CODE HTML ICI ---
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>NBB Report Generator</title>
<style>
  body { background: #0A0E1A; color: #E2E8F0; font-family: sans-serif; display: flex; flex-direction: column; align-items: center; padding: 50px; }
  .dropzone { border: 2px dashed #1E293B; border-radius: 12px; padding: 60px; text-align: center; cursor: pointer; background: #111827; width: 100%; max-width: 600px; }
  .btn { background: #38BDF8; color: #0A0E1A; padding: 15px 30px; border: none; border-radius: 8px; font-weight: bold; cursor: pointer; margin-top: 20px; }
  .status { margin-top: 20px; font-family: monospace; }
</style>
</head>
<body>
  <h1>NBB Report Generator</h1>
  <div class="dropzone" id="dz">
    <input type="file" id="fi" style="display:none" accept=".xlsx,.xls">
    <p id="dzLabel">Glissez votre Excel ici ou cliquez pour choisir</p>
  </div>
  <button class="btn" id="btn">Générer le PowerPoint</button>
  <div id="status" class="status"></div>

  <script>
    const dz = document.getElementById('dz');
    const fi = document.getElementById('fi');
    dz.onclick = () => fi.click();
    fi.onchange = () => { document.getElementById('dzLabel').innerText = fi.files[0].name; };

    document.getElementById('btn').onclick = async () => {
      if(!fi.files[0]) return alert("Choisissez un fichier !");
      const fd = new FormData();
      fd.append('file', fi.files[0]);
      document.getElementById('status').innerText = "Génération en cours...";
      
      const res = await fetch('/generate', { method: 'POST', body: fd });
      const data = await res.json();
      
      if(data.status === 'success') {
        document.getElementById('status').innerHTML = `<a href="${data.download_url}" style="color:#10B981">✅ Télécharger le PPTX</a>`;
      } else {
        document.getElementById('status').innerText = "❌ Erreur : " + data.error;
      }
    };
  </script>
</body>
</html>"""

@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/generate", methods=["POST"])
def generate():
    try:
        file = request.files.get("file")
        if not file: return jsonify({"error": "No file"}), 400
        df = pd.read_excel(file)
        pptx_bytes = build_agency_pptx(df, TEMPLATE_PATH)
        token = str(uuid.uuid4())
        _cache[token] = (pptx_bytes, f"NBB_Report_{datetime.now().strftime('%Y%m%d')}.pptx", datetime.now() + timedelta(minutes=10))
        return jsonify({"status": "success", "download_url": f"{request.host_url.rstrip('/')}/download/{token}"})
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 500

@app.route("/download/<token>")
def download(token):
    entry = _cache.get(token)
    if not entry: abort(404)
    return send_file(io.BytesIO(entry[0]), as_attachment=True, download_name=entry[1])

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
