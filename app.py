import io, json, os, threading, traceback, uuid
from datetime import datetime, timedelta
import pandas as pd
from flask import Flask, abort, jsonify, render_template_string, request, send_file

# Importation du moteur de génération
from generate_pptx_v3 import build_agency_pptx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024 # 20 MB

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
        _cache[token] = {"data": data, "filename": filename, "expiry": expiry}
    return token

# --- INTERFACE GRAPHIQUE ---
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>NBB Report Generator</title>
<style>
  body { background: #0A0E1A; color: #E2E8F0; font-family: sans-serif; display: flex; flex-direction: column; align-items: center; padding: 50px; }
  .card { background: #111827; border: 1px solid #1E293B; border-radius: 12px; padding: 30px; width: 100%; max-width: 500px; text-align: center; }
  .dropzone { border: 2px dashed #38BDF8; border-radius: 10px; padding: 40px; margin: 20px 0; cursor: pointer; transition: 0.2s; }
  .dropzone:hover { background: #1C2333; }
  .btn { background: #38BDF8; color: #0A0E1A; padding: 12px 24px; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; width: 100%; }
  .status { margin-top: 20px; font-family: monospace; font-size: 13px; }
  a { color: #10B981; text-decoration: none; font-weight: bold; }
</style>
</head>
<body>
  <div class="card">
    <h2>NBB Generator v13</h2>
    <p style="color: #64748B; font-size: 14px;">Chargez votre Excel pour remplir le template Glass</p>
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
      if(!fi.files[0]) return alert("Veuillez choisir un fichier Excel !");
      const fd = new FormData();
      fd.append('file', fi.files[0]);
      
      const status = document.getElementById('status');
      status.innerText = "⏳ Génération du PPTX...";
      
      try {
        const res = await fetch('/generate', { method: 'POST', body: fd });
        const d = await res.json();
        if(d.status === 'success') {
          status.innerHTML = `✅ Terminé !<br><br><a href="${d.download_url}" target="_blank">⬇️ TÉLÉCHARGER LE PPTX</a>`;
        } else {
          status.innerText = "❌ Erreur : " + d.error;
        }
      } catch(e) {
        status.innerText = "❌ Erreur réseau.";
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
        if not file:
            return jsonify({"status": "error", "error": "Fichier manquant"}), 400

        # Lecture Excel 
        df = pd.read_excel(file)
        
        # Lancement du moteur de génération V3 
        pptx_bytes = build_agency_pptx(df, TEMPLATE_PATH)

        # Stockage propre dans le cache 
        filename = f"NBB_Report_{datetime.now().strftime('%H%M')}.pptx"
        token = store_file(pptx_bytes, filename)
        
        return jsonify({
            "status": "success",
            "download_url": f"{request.host_url.rstrip('/')}/download/{token}"
        })

    except Exception as e:
        return jsonify({
            "status": "error", 
            "error": str(e),
            "detail": traceback.format_exc()
        }), 500

@app.route("/download/<token>")
def download(token):
    with _lock:
        entry = _cache.get(token)
    
    if not entry:
        abort(404, "Lien expiré.")
    
    return send_file(
        io.BytesIO(entry["data"]),
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        as_attachment=True,
        download_name=entry["filename"]
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
