import io
import os
import uuid
import time
import traceback
import pandas as pd
from flask import Flask, jsonify, render_template_string, request, send_file, abort
from pptx import Presentation

# Import de tes scripts de génération
from fill_template import load_data_from_df, build_placeholders, replace_all_placeholders
from generate_pptx_v3 import build_agency_pptx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30 MB

# Dossier temporaire pour stocker les PPTX générés
TEMP_DIR = "/tmp/pptx_gen"
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

# Localisation du template
TEMPLATE_NAME = "T21_HK_Agencies_Glass_v13.pptx"
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), TEMPLATE_NAME)

def cleanup_old_files():
    """Supprime les fichiers de plus de 15 minutes dans /tmp/"""
    now = time.time()
    for f in os.listdir(TEMP_DIR):
        fpath = os.path.join(TEMP_DIR, f)
        if os.stat(fpath).st_mtime < now - 900: # 900 secondes = 15 min
            try:
                os.remove(fpath)
            except:
                pass

# --- Interface HTML ---
HTML = r"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>NBB Generator v13</title>
    <style>
        body { background: #0A0E1A; color: white; font-family: sans-serif; text-align: center; padding: 50px; }
        .card { background: #111827; border: 1px solid #1E293B; padding: 30px; border-radius: 12px; display: inline-block; width: 400px; }
        .btn { background: #38BDF8; color: #0A0E1A; padding: 12px 24px; border: none; border-radius: 6px; cursor: pointer; font-weight: bold; width: 100%; margin-top: 10px; }
        .btn:disabled { opacity: 0.5; }
        #dl { display: none; margin-top: 20px; background: #10B981; color: white; padding: 15px; border-radius: 8px; text-decoration: none; font-weight: bold; }
        input[type="file"] { margin-bottom: 20px; color: #64748B; }
    </style>
</head>
<body>
    <div class="card">
        <h2>📊 NBB Generator</h2>
        <p style="color: #64748B; font-size: 0.9em;">Upload Excel → PPTX</p>
        <input type="file" id="fi" accept=".xlsx">
        <button class="btn" id="btn">GÉNÉRER LE RAPPORT</button>
        <div id="status" style="margin-top: 20px; font-size: 14px;"></div>
        <a id="dl" href="#">⬇ TÉLÉCHARGER LE PPTX</a>
    </div>

    <script>
        const btn = document.getElementById('btn'), fi = document.getElementById('fi'), dl = document.getElementById('dl'), st = document.getElementById('status');
        btn.onclick = async () => {
            if(!fi.files[0]) return alert("Veuillez choisir un fichier Excel.");
            btn.disabled = true; st.innerText = "⏳ Génération en cours..."; dl.style.display = 'none';

            const fd = new FormData();
            fd.append('file', fi.files[0]);

            try {
                const res = await fetch('/generate', { method: 'POST', body: fd });
                const d = await res.json();
                if(d.status === 'success') {
                    st.innerText = "✅ Rapport prêt !";
                    dl.href = d.download_url;
                    dl.style.display = 'block';
                } else {
                    st.innerText = "❌ Erreur: " + d.error;
                }
            } catch(e) { st.innerText = "❌ Erreur de connexion au serveur."; }
            btn.disabled = false;
        };
    </script>
</body>
</html>"""

@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/generate", methods=["POST"])
def generate():
    cleanup_old_files()
    try:
        file = request.files.get("file")
        if not file:
            return jsonify({"status": "error", "error": "Fichier manquant"}), 400
        
        # 1. Charger les données
        df = pd.read_excel(file)
        
        # 2. Utiliser build_agency_pptx (v3)
        # Note : On s'assure que build_agency_pptx renvoie bien des BYTES
        pptx_bytes = build_agency_pptx(df, TEMPLATE_PATH)
        
        # 3. Sauvegarder physiquement le fichier dans /tmp/
        token = str(uuid.uuid4())
        file_filename = f"NBB_Report_{token[:8]}.pptx"
        file_path = os.path.join(TEMP_DIR, f"{token}.pptx")
        
        with open(file_path, "wb") as f:
            f.write(pptx_bytes)
        
        return jsonify({
            "status": "success",
            "download_url": f"/download/{token}"
        })
    except Exception as e:
        print(traceback.format_exc())
        return jsonify({"status": "error", "error": str(e)}), 500

@app.route("/download/<token>")
def download(token):
    file_path = os.path.join(TEMP_DIR, f"{token}.pptx")
    
    if not os.path.exists(file_path):
        return "ERREUR : Le fichier est introuvable ou a expiré. Veuillez cliquer sur 'Générer' à nouveau.", 404

    return send_file(
        file_path,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        as_attachment=True,
        download_name="NBB_Report_Generated.pptx"
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
