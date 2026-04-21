import io, json, os, threading, traceback, uuid
from datetime import datetime, timedelta
import pandas as pd
from flask import Flask, abort, jsonify, render_template_string, request, send_file

# Importation du moteur de génération
from generate_pptx_v3 import build_agency_pptx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024 

# Configuration du template
TEMPLATE_NAME = "T21_HK_Agencies_Glass_v13.pptx"
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), TEMPLATE_NAME)

# Gestion du cache mémoire
_cache = {}
_lock = threading.Lock()

def store_file(data, filename):
    token = str(uuid.uuid4())
    expiry = datetime.now() + timedelta(minutes=10)
    with _lock:
        # On stocke en dictionnaire pour être explicite
        _cache[token] = {"data": data, "filename": filename, "expiry": expiry}
    return token

# --- INTERFACE GRAPHIQUE ---
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>NBB Generator v13</title>
<style>
  body { background: #0A0E1A; color: #E2E8F0; font-family: sans-serif; display: flex; flex-direction: column; align-items: center; padding: 50px; }
  .card { background: #111827; border: 1px solid #1E293B; border-radius: 12px; padding: 30px; width: 100%; max-width: 500px; text-align: center; }
  .dropzone { border: 2px dashed #38BDF8; border-radius: 10px; padding: 40px; margin: 20px 0; cursor: pointer; }
  .btn { background: #38BDF8; color: #0A0E1A; padding: 12px 24px; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; width: 100%; }
  .status { margin-top: 20px; font-family: monospace; font-size: 13px; }
  a { color: #10B981; text-decoration: none; font-weight: bold; border: 1px solid #10B981; padding: 10px; border-radius: 5px; display: inline-block; }
</style>
</head>
<body>
  <div class="card">
    <h2>NBB Generator v13</h2>
    <div class="dropzone" id="dz">
      <input type="file" id="fi" style="display:none" accept=".xlsx,.xls">
      <p id="dzLabel">📁 Cliquez ou glissez l'Excel ici</p>
    </div>
    <button class="btn" id="btn">Générer le rapport</button>
    <div id="status" class="status"></div>
  </div>
  <script>
    const dz = document.getElementById('dz');
    const fi = document.getElementById('fi');
    dz.onclick = () => fi.click();
    fi.onchange = () => { if(fi.files[0]) document.getElementById('dzLabel').innerText = "✅ " + fi.files[0].name; };
    document.getElementById('btn').onclick = async () => {
      if(!fi.files[0]) return alert("Fichier manquant !");
      const fd = new FormData();
      fd.append('file', fi.files[0]);
      const status = document.getElementById('status');
      status.innerText = "⏳ Génération en cours...";
      try {
        const res = await fetch('/generate', { method: 'POST', body: fd });
        const d = await res.json();
        if(d.status === 'success') {
          status.innerHTML = `<br><a href="${d.download_url}" target="_blank">⬇️ TÉLÉCHARGER LE PPTX</a>`;
        } else {
          status.innerText = "❌ Erreur : " + d.error;
        }
      } catch(e) { status.innerText = "❌ Erreur réseau."; }
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
        if not file:
            return jsonify({"status": "error", "error": "Fichier Excel manquant"}), 400

        df = pd.read_excel(file)
        # Appel au script generate_pptx_v3.py
        pptx_bytes = build_agency_pptx(df, TEMPLATE_PATH)

        filename = f"NBB_Report_{datetime.now().strftime('%H%M')}.pptx"
        token = store_file(pptx_bytes, filename)
        
        return jsonify({
            "status": "success",
            "download_url": f"{request.host_url.rstrip('/')}/download/{token}"
        })
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 500

@app.route("/download/<token>")
def download(token):
    with _lock:
        entry = _cache.get(token)
    
    if not entry:
        abort(404, "Fichier non trouvé ou expiré.")

    # SÉCURITÉ ANTI-TUPLE : On vérifie si c'est un dictionnaire ou un tuple
    if isinstance(entry, dict):
        data = entry["data"]
        name = entry["filename"]
    else:
        # Si c'est un vieux tuple resté en mémoire
        data = entry[0]
        name = entry[1]
    
    return send_file(
        io.BytesIO(data),
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        as_attachment=True,
        download_name=name
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
