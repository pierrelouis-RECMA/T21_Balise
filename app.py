import io, json, os, re, threading, traceback, uuid, zipfile
from datetime import datetime, timedelta

import pandas as pd
from flask import Flask, abort, jsonify, render_template_string, request, send_file
from pptx import Presentation

# Importation de tes modules personnalisés
from fill_template import load_data_from_df, build_placeholders, replace_all_placeholders
from generate_pptx_v2 import build_agency_pptx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # Limite à 20 MB

# Chemin vers le template PPTX
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "T21_HK_Agencies_Glass_v13.pptx")

# ── Gestion du cache mémoire (20 min TTL) ─────────────────────
_cache = {}
_lock = threading.Lock()

def store_file(data: bytes, filename: str) -> str:
    """Stocke le fichier généré en mémoire et retourne un token unique."""
    token = str(uuid.uuid4())
    expiry = datetime.now() + timedelta(minutes=20) # Lien valide 20 minutes
    with _lock:
        _cache[token] = {"data": data, "filename": filename, "expiry": expiry}
    return token

def purge_expired_entries():
    """Supprime les fichiers expirés du cache pour libérer la mémoire."""
    now = datetime.now()
    with _lock:
        keys_to_del = [k for k, v in _cache.items() if v["expiry"] < now]
        for k in keys_to_del:
            del _cache[k]

# ── Interface HTML (Version Dark Mode) ────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NBB Report Generator</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=DM+Mono&display=swap" rel="stylesheet">
<style>
:root{--bg:#0A0E1A;--surface:#111827;--border:#1E293B;--accent:#38BDF8;--win:#10B981;--dep:#F43F5E;--text:#E2E8F0;--muted:#64748B;}
body{background:var(--bg);color:var(--text);font-family:'DM Sans',sans-serif;margin:0;display:flex;flex-direction:column;align-items:center;min-height:100vh;}
header{width:100%;padding:20px 40px;background:var(--surface);border-bottom:1px solid var(--border);display:flex;align-items:center;gap:12px;}
.dot{width:10px;height:10px;background:var(--accent);border-radius:50%;box-shadow:0 0 10px var(--accent);}
h1{font-size:16px;text-transform:uppercase;letter-spacing:1px;margin:0;}
main{max-width:600px;width:90%;padding:40px 20px;}
.card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:30px;margin-bottom:20px;}
.form-group{margin-bottom:15px;display:flex;flex-direction:column;gap:5px;}
label{font-size:12px;color:var(--muted);font-weight:bold;}
select, input{background:#1C2333;border:1px solid var(--border);color:white;padding:10px;border-radius:6px;outline:none;}
.dropzone{border:2px dashed var(--border);padding:40px;text-align:center;border-radius:10px;cursor:pointer;transition:0.3s;}
.dropzone:hover{border-color:var(--accent);background:rgba(56,189,248,0.05);}
.btn{width:100%;padding:14px;background:var(--accent);color:#000;border:none;border-radius:6px;font-weight:bold;cursor:pointer;font-size:15px;}
.btn:disabled{opacity:0.3;cursor:not-allowed;}
.status{margin-top:20px;padding:15px;border-radius:8px;display:none;font-size:14px;}
.success{background:rgba(16,185,129,0.1);border:1px solid var(--win);color:var(--win);display:block;}
.error{background:rgba(244,63,94,0.1);border:1px solid var(--dep);color:var(--dep);display:block;}
.dl-btn{display:inline-block;margin-top:10px;padding:10px 20px;background:var(--win);color:white;text-decoration:none;border-radius:5px;font-weight:bold;}
</style>
</head>
<body>
<header><div class="dot"></div><h1>NBB Generator v13</h1></header>
<main>
    <div class="card">
        <div class="form-group"><label>Marché</label><input type="text" id="market" value="HK"></div>
        <div class="form-group"><label>Année</label><input type="text" id="year" value="2025"></div>
        <div class="dropzone" id="dz"><input type="file" id="fi" style="display:none" accept=".xlsx,.xls"><p id="dzTxt">📁 Cliquez ou déposez l'Excel ici</p></div>
        <button class="btn" id="btn" style="margin-top:20px">Générer la présentation</button>
        <div id="status"></div>
    </div>
</main>
<script>
const dz=document.getElementById('dz'),fi=document.getElementById('fi'),btn=document.getElementById('btn'),st=document.getElementById('status');
dz.onclick=()=>fi.click();
fi.onchange=()=>{if(fi.files[0])document.getElementById('dzTxt').innerText="✅ "+fi.files[0].name;};
btn.onclick=async()=>{
    if(!fi.files[0])return alert("Fichier manquant");
    btn.disabled=true; st.innerHTML="⏳ Génération..."; st.className="status success";
    const fd=new FormData(); fd.append('file',fi.files[0]); fd.append('market',document.getElementById('market').value); fd.append('year',document.getElementById('year').value);
    try{
        const res=await fetch('/generate',{method:'POST',body:fd});
        const d=await res.json();
        if(d.status==='success'){
            st.innerHTML=`<strong>✅ Terminé !</strong><br><br><a class="dl-btn" href="${d.download_url}">⬇ Télécharger le PPTX</a>`;
        }else{
            st.innerHTML="❌ Erreur: "+d.error; st.className="status error";
        }
    }catch(e){st.innerHTML="❌ Erreur réseau"; st.className="status error";}
    btn.disabled=false;
};
</script>
</body>
</html>"""

# ── Routes API ───────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/generate", methods=["POST"])
def generate():
    try:
        purge_expired_entries()
        file = request.files.get("file")
        market = request.form.get("market", "Market")
        year = request.form.get("year", "2025")

        if not file:
            return jsonify({"status": "error", "error": "Fichier Excel manquant"}), 400

        # 1. Lecture Excel
        df = pd.read_excel(file)

        # 2. Remplissage slides 1-6 (fill_template.py)
        data = load_data_from_df(df)
        ph = build_placeholders(data)
        prs = Presentation(TEMPLATE_PATH)
        replace_all_placeholders(prs, ph)
        
        buf = io.BytesIO()
        prs.save(buf)
        prefilled_bytes = buf.getvalue()

        # 3. Génération cards dynamiques 7+ (generate_pptx_v2.py)
        # On passe les bytes déjà remplis pour que generate_pptx_v2 les complète
        final_pptx_bytes = build_agency_pptx(df, TEMPLATE_PATH, prefilled_prs_bytes=prefilled_bytes)

        # 4. Stockage et réponse
        filename = f"NBB_Report_{market}_{year}.pptx"
        token = store_file(final_pptx_bytes, filename)

        return jsonify({
            "status": "success",
            "download_url": f"{request.host_url.rstrip('/')}/download/{token}",
            "filename": filename
        })

    except Exception as e:
        app.logger.error(traceback.format_exc())
        return jsonify({"status": "error", "error": str(e)}), 500

@app.route("/download/<token>")
def download(token):
    # On ne fait la purge qu'au début de la route generate pour éviter de supprimer le fichier
    # qu'on essaie justement de télécharger.
    with _lock:
        entry = _cache.get(token)
    
    if not entry:
        abort(404, "Le fichier n'existe plus ou le lien a expiré (20 min).")

    if datetime.now() > entry["expiry"]:
        with _lock:
            _cache.pop(token, None)
        abort(410, "Lien expiré.")

    return send_file(
        io.BytesIO(entry["data"]),
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        as_attachment=True,
        download_name=entry["filename"]
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
