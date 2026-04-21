import io, os, threading, uuid, traceback
from datetime import datetime, timedelta
import pandas as pd
from flask import Flask, abort, jsonify, render_template_string, request, send_file
from pptx import Presentation

# Imports de tes scripts
from fill_template import load_data_from_df, build_placeholders, replace_all_placeholders
from generate_pptx_v2 import build_agency_pptx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30 MB

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "T21_HK_Agencies_Glass_v13.pptx")

# Cache global
_cache = {}
_lock = threading.Lock()

def store_file(data, filename):
    token = str(uuid.uuid4())
    # On garde le fichier 30 minutes
    expiry = datetime.now() + timedelta(minutes=30)
    with _lock:
        _cache[token] = {"data": data, "filename": filename, "expiry": expiry}
    return token

# --- HTML ---
HTML = r"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>NBB Generator</title>
    <style>
        body { background: #0A0E1A; color: white; font-family: sans-serif; text-align: center; padding: 50px; }
        .card { background: #111827; border: 1px solid #1E293B; padding: 30px; border-radius: 12px; display: inline-block; }
        .btn { background: #38BDF8; color: #0A0E1A; padding: 12px 24px; border: none; border-radius: 6px; cursor: pointer; font-weight: bold; }
        .btn:disabled { opacity: 0.5; }
        #dl { display: none; margin-top: 20px; background: #10B981; color: white; padding: 15px; border-radius: 8px; text-decoration: none; }
    </style>
</head>
<body>
    <div class="card">
        <h2>Générateur NBB</h2>
        <input type="file" id="fi" accept=".xlsx"><br><br>
        <button class="btn" id="btn">Générer le rapport</button>
        <br><br>
        <div id="status"></div>
        <a id="dl" href="#" download>⬇ TÉLÉCHARGER LE PPTX</a>
    </div>

    <script>
        const btn = document.getElementById('btn'), fi = document.getElementById('fi'), dl = document.getElementById('dl'), st = document.getElementById('status');
        btn.onclick = async () => {
            if(!fi.files[0]) return alert("Sélectionnez un fichier");
            btn.disabled = true; st.innerText = "⏳ Analyse et génération en cours...";
            dl.style.display = 'none';

            const fd = new FormData();
            fd.append('file', fi.files[0]);

            try {
                const res = await fetch('/generate', { method: 'POST', body: fd });
                const d = await res.json();
                if(d.status === 'success') {
                    st.innerText = "✅ Généré avec succès !";
                    dl.href = d.download_url;
                    dl.style.display = 'block';
                } else {
                    st.innerText = "❌ Erreur: " + d.error;
                }
            } catch(e) { st.innerText = "❌ Erreur réseau"; }
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
    try:
        file = request.files.get("file")
        if not file: return jsonify({"status": "error", "error": "Fichier manquant"}), 400
        
        # 1. Traitement Data
        df = pd.read_excel(file)
        data = load_data_from_df(df)
        ph = build_placeholders(data)
        
        # 2. Remplissage Slides 1-6
        prs = Presentation(TEMPLATE_PATH)
        replace_all_placeholders(prs, ph)
        
        buf_pre = io.BytesIO()
        prs.save(buf_pre)
        pre_bytes = buf_pre.getvalue()
        
        # 3. Génération Slides 7+
        final_bytes = build_agency_pptx(df, TEMPLATE_PATH, prefilled_prs_bytes=pre_bytes)
        
        # 4. Stockage
        token = store_file(final_bytes, "NBB_Report.pptx")
        
        return jsonify({
            "status": "success",
            "download_url": f"/download/{token}"
        })
    except Exception as e:
        print(traceback.format_exc())
        return jsonify({"status": "error", "error": str(e)}), 500

@app.route("/download/<token>")
def download(token):
    with _lock:
        entry = _cache.get(token)
    
    if not entry:
        return "Erreur : Le fichier a expiré ou n'existe pas sur ce serveur. Réessayez.", 404

    return send_file(
        io.BytesIO(entry["data"]),
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        as_attachment=True,
        download_name=entry["filename"]
    )

if __name__ == "__main__":
    # Local test
    app.run(host="0.0.0.0", port=5000)
