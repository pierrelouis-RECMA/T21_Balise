"""
app.py — NBB Report Generator
──────────────────────────────────────────────────────────────
Render.com deployment.

Pipeline :
  1. fill_template.py   → remplace les {{balises}} slides 1-6
  2. generate_pptx_v2.py → génère les agency cards slides 7+
     (XML direct dans le ZIP — évite le bug pptx add_slide)

Start : gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120
"""

import io, json, os, re, threading, traceback, uuid, zipfile
from datetime import datetime, timedelta

import pandas as pd
from flask import Flask, abort, jsonify, render_template_string, request, send_file
from pptx import Presentation

from fill_template   import load_data_from_df, build_placeholders, replace_all_placeholders
from generate_pptx_v2 import build_agency_pptx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "T21_HK_Agencies_Glass_v13.pptx")

# ── In-memory cache (10 min TTL) ──────────────────────────────
_cache: dict = {}
_lock = threading.Lock()

def store_file(data: bytes, filename: str) -> str:
    token  = str(uuid.uuid4())
    expiry = datetime.now() + timedelta(minutes=10)
    with _lock:
        _cache[token] = {"data": data, "filename": filename, "expiry": expiry}
    return token

def purge_cache():
    now = datetime.now()
    with _lock:
        dead = [k for k, v in _cache.items() if v["expiry"] < now]
        for k in dead:
            del _cache[k]

# ─────────────────────────────────────────────────────────────
# HTML
# ─────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NBB Report Generator</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
:root{--bg:#0A0E1A;--surface:#111827;--surface2:#1C2333;--border:#1E293B;
      --accent:#38BDF8;--win:#10B981;--dep:#F43F5E;--text:#E2E8F0;--muted:#64748B;}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:'DM Sans',sans-serif;
     min-height:100vh;display:flex;flex-direction:column;align-items:center}
header{width:100%;padding:18px 40px;border-bottom:1px solid var(--border);
       background:var(--surface);display:flex;align-items:center;gap:12px}
.dot{width:10px;height:10px;border-radius:50%;background:var(--accent);
     box-shadow:0 0 12px var(--accent)}
header h1{font-size:14px;font-weight:500;letter-spacing:.1em;text-transform:uppercase}
header span{margin-left:auto;font-family:'DM Mono',monospace;font-size:11px;color:var(--muted)}
main{width:100%;max-width:640px;padding:56px 24px 80px;display:flex;flex-direction:column;gap:28px}
.tagline{text-align:center}
.tagline h2{font-size:28px;font-weight:300;color:#fff;letter-spacing:-.02em}
.tagline h2 em{font-style:normal;color:var(--accent)}
.tagline p{margin-top:8px;font-size:13px;color:var(--muted);line-height:1.6}
.card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:26px 30px}
.card-label{font-size:11px;font-weight:500;letter-spacing:.12em;text-transform:uppercase;
            color:var(--muted);margin-bottom:16px;display:flex;align-items:center;gap:8px}
.card-label::before{content:'';width:16px;height:1px;background:var(--accent)}
.col-doc{display:flex;flex-direction:column;gap:5px}
.col-row{display:flex;align-items:center;gap:10px;padding:8px 11px;
         background:var(--surface2);border-radius:6px;font-size:12px}
.col-name{font-family:'DM Mono',monospace;color:var(--accent);min-width:185px;flex-shrink:0}
.col-desc{color:var(--muted)}
.badge{font-size:9px;padding:2px 6px;border-radius:3px;font-weight:600;
       letter-spacing:.05em;text-transform:uppercase;margin-left:auto;flex-shrink:0}
.req{background:rgba(244,63,94,.15);color:var(--dep)}
.opt{background:rgba(100,116,139,.15);color:var(--muted)}
.form-row{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:14px}
.form-group{display:flex;flex-direction:column;gap:5px}
.form-group label{font-size:11px;color:var(--muted);font-weight:500;letter-spacing:.05em}
.form-group select,.form-group input{background:var(--surface2);border:1px solid var(--border);
  border-radius:7px;padding:10px 13px;color:var(--text);font-family:'DM Sans',sans-serif;
  font-size:14px;outline:none;transition:border-color .2s;-webkit-appearance:none;appearance:none}
.form-group select:focus,.form-group input:focus{border-color:var(--accent)}
.dropzone{border:1.5px dashed var(--border);border-radius:8px;padding:32px;
          text-align:center;cursor:pointer;transition:border-color .2s,background .2s;
          position:relative;overflow:hidden}
.dropzone:hover,.dropzone.over{border-color:var(--accent);background:rgba(56,189,248,.04)}
.dropzone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.dz-icon{font-size:24px;margin-bottom:8px;opacity:.6}
.dz-label{font-size:14px;font-weight:500}
.dz-hint{font-size:12px;color:var(--muted);margin-top:4px;font-family:'DM Mono',monospace}
.dz-name{display:none;margin-top:10px;padding:6px 12px;background:rgba(16,185,129,.1);
         border:1px solid rgba(16,185,129,.3);border-radius:5px;
         font-size:12px;font-family:'DM Mono',monospace;color:var(--win)}
.btn{margin-top:20px;width:100%;padding:14px;background:var(--accent);color:#0A0E1A;
     border:none;border-radius:7px;font-size:15px;font-weight:600;cursor:pointer;
     transition:background .2s,opacity .2s;display:flex;align-items:center;justify-content:center;gap:8px}
.btn:hover:not(:disabled){background:#7DD3FC}
.btn:disabled{opacity:.4;cursor:not-allowed}
.status{margin-top:16px;padding:14px 16px;border-radius:8px;font-size:13px;display:none}
.status.loading{display:flex;align-items:center;gap:10px;
                background:rgba(56,189,248,.07);border:1px solid rgba(56,189,248,.2);color:var(--accent)}
.status.success{display:block;background:rgba(16,185,129,.07);
                border:1px solid rgba(16,185,129,.25);color:var(--win)}
.status.error{display:block;background:rgba(244,63,94,.07);
              border:1px solid rgba(244,63,94,.25);color:var(--dep)}
.spinner{width:16px;height:16px;border:2px solid rgba(56,189,248,.2);
         border-top-color:var(--accent);border-radius:50%;animation:spin .7s linear infinite;flex-shrink:0}
@keyframes spin{to{transform:rotate(360deg)}}
.dl-link{display:inline-block;margin-top:10px;padding:9px 20px;background:var(--win);
         color:#fff;text-decoration:none;border-radius:6px;font-weight:600;font-size:13px}
.dl-link:hover{opacity:.85}
.stats-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-top:18px}
.stat{background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:12px 14px}
.stat-label{font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);margin-bottom:3px}
.stat-val{font-size:20px;font-weight:600;font-family:'DM Mono',monospace}
.stat-sub{font-size:11px;color:var(--muted);margin-top:1px}
</style>
</head>
<body>
<header>
  <div class="dot"></div>
  <h1>NBB Report Generator</h1>
  <span>Powered by Render</span>
</header>
<main>
  <div class="tagline">
    <h2>Upload Excel → <em>Présentation</em></h2>
    <p>Slides 1–6 remplies automatiquement · Agency cards (slide 7+) générées dynamiquement<br>
       4 agences par slide · triées par NBB décroissant</p>
  </div>

  <div class="card">
    <div class="card-label">Colonnes Excel requises</div>
    <div class="col-doc">
      <div class="col-row"><span class="col-name">Agency</span><span class="col-desc">Nom de l'agence</span><span class="badge req">requis</span></div>
      <div class="col-row"><span class="col-name">NewBiz</span><span class="col-desc">WIN · DEPARTURE · RETENTION</span><span class="badge req">requis</span></div>
      <div class="col-row"><span class="col-name">Advertiser</span><span class="col-desc">Nom de l'annonceur / client</span><span class="badge req">requis</span></div>
      <div class="col-row"><span class="col-name">Integrated Spends</span><span class="col-desc">Budget $m — positif WIN, négatif DEPARTURE</span><span class="badge req">requis</span></div>
      <div class="col-row"><span class="col-name">Date of announcement</span><span class="col-desc">Affiché sur les agency cards</span><span class="badge opt">optionnel</span></div>
      <div class="col-row"><span class="col-name">Incumbent</span><span class="col-desc">Agence précédente (WIN)</span><span class="badge opt">optionnel</span></div>
    </div>
  </div>

  <div class="card">
    <div class="card-label">Générer le rapport</div>

    <div class="form-row">
      <div class="form-group">
        <label>Marché</label>
        <select id="market">
          <option value="Mexico">Mexico</option>
          <option value="Hong_Kong">Hong Kong</option>
          <option value="Singapore">Singapore</option>
          <option value="Indonesia">Indonesia</option>
          <option value="Other">Other</option>
        </select>
      </div>
      <div class="form-group">
        <label>Année</label>
        <input type="text" id="year" value="2025" placeholder="2025">
      </div>
    </div>

    <div class="dropzone" id="dz">
      <input type="file" id="fi" accept=".xlsx,.xls">
      <div class="dz-icon">📊</div>
      <div class="dz-label">Glissez votre Excel ici ou cliquez</div>
      <div class="dz-hint">.xlsx ou .xls · max 20 MB</div>
      <div class="dz-name" id="dzName"></div>
    </div>

    <button class="btn" id="btn" disabled>Générer la présentation</button>

    <div class="status loading" id="sLoad"><div class="spinner"></div><span>Génération en cours…</span></div>
    <div class="status success" id="sOk"></div>
    <div class="status error"   id="sErr"></div>
  </div>
</main>

<script>
const dz=document.getElementById('dz'),fi=document.getElementById('fi'),btn=document.getElementById('btn');
const sLoad=document.getElementById('sLoad'),sOk=document.getElementById('sOk'),sErr=document.getElementById('sErr');

dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('over')});
dz.addEventListener('dragleave',()=>dz.classList.remove('over'));
dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('over');
  if(e.dataTransfer.files[0]){fi.files=e.dataTransfer.files;onFile(e.dataTransfer.files[0])}});
fi.addEventListener('change',()=>{if(fi.files[0])onFile(fi.files[0])});

function onFile(f){
  document.getElementById('dzName').style.display='block';
  document.getElementById('dzName').textContent='✓ '+f.name;
  btn.disabled=false;
}

btn.addEventListener('click',async()=>{
  if(!fi.files[0])return;
  btn.disabled=true;
  sLoad.style.display='flex'; sOk.style.display='none'; sErr.style.display='none';
  const fd=new FormData();
  fd.append('file',fi.files[0]);
  fd.append('market',document.getElementById('market').value);
  fd.append('year',document.getElementById('year').value);
  try{
    const res=await fetch('/generate',{method:'POST',body:fd});
    const d=await res.json();
    sLoad.style.display='none'; btn.disabled=false;
    if(d.status==='success'){
      sOk.style.display='block';
      sOk.innerHTML=`<strong>✅ Prêt !</strong> ${d.filename}<br>
        <a class="dl-link" href="${d.download_url}" download>⬇ Télécharger le PPTX</a>
        <div class="stats-grid">
          <div class="stat"><div class="stat-label">Agences</div><div class="stat-val">${d.agencies}</div><div class="stat-sub">dans l'Excel</div></div>
          <div class="stat"><div class="stat-label">Slides</div><div class="stat-val">${d.total_slides}</div><div class="stat-sub">au total</div></div>
          <div class="stat"><div class="stat-label">Lignes</div><div class="stat-val">${d.records}</div><div class="stat-sub">W + D + R</div></div>
        </div>`;
    }else{
      sErr.style.display='block';
      sErr.innerHTML=`<strong>❌ Erreur</strong><br>${d.error}`;
    }
  }catch(e){
    sLoad.style.display='none'; btn.disabled=false;
    sErr.style.display='block';
    sErr.innerHTML=`<strong>❌ Erreur réseau</strong><br>${e.message}`;
  }
});
</script>
</body>
</html>"""

# ─────────────────────────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/generate", methods=["POST"])
def generate():
    try:
        file   = request.files.get("file")
        market = request.form.get("market", "Unknown")
        year   = request.form.get("year",   "2025")

        if not file:
            return jsonify({"status": "error", "error": "Fichier Excel manquant"}), 400

        if not os.path.exists(TEMPLATE_PATH):
            return jsonify({"status": "error",
                            "error": f"Template introuvable : {os.path.basename(TEMPLATE_PATH)}"}), 500

        # ── Lire l'Excel ──────────────────────────────────────
        df = pd.read_excel(file)
        required = ["Agency", "NewBiz", "Advertiser", "Integrated Spends"]
        missing  = [c for c in required if c not in df.columns]
        if missing:
            return jsonify({"status": "error",
                            "error": f"Colonnes manquantes : {', '.join(missing)}. "
                                     f"Disponibles : {', '.join(df.columns.tolist())}"}), 400

        # ── Étape 1 : remplir les slides statiques 1-6 ───────
        data      = load_data_from_df(df)
        ph        = build_placeholders(data)
        prs       = Presentation(TEMPLATE_PATH)
        replace_all_placeholders(prs, ph)
        buf = io.BytesIO()
        prs.save(buf)
        prefilled = buf.getvalue()

        # ── Étape 2 : générer les agency cards 7+ (XML direct)
        pptx_bytes = build_agency_pptx(df, TEMPLATE_PATH,
                                       prefilled_prs_bytes=prefilled)

        # ── Stats ─────────────────────────────────────────────
        with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
            total_slides = len([n for n in z.namelist()
                                if re.match(r"ppt/slides/slide\d+\.xml$", n)])

        ts       = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f"NBB_{market}_{year}_{ts}.pptx"
        token    = store_file(pptx_bytes, filename)
        purge_cache()

        return jsonify({
            "status":       "success",
            "download_url": f"{request.host_url.rstrip('/')}/download/{token}",
            "filename":     filename,
            "agencies":     int(df["Agency"].dropna().nunique()),
            "total_slides": total_slides,
            "records":      len(df),
        })

    except Exception as e:
        app.logger.error(traceback.format_exc())
        return jsonify({"status": "error", "error": str(e),
                        "detail": traceback.format_exc()}), 500


@app.route("/download/<token>")
def download(token):
    purge_cache()
    with _lock:
        entry = _cache.get(token)
    if not entry:
        abort(404, "Lien expiré — régénérez le rapport.")
    if datetime.now() > entry["expiry"]:
        with _lock:
            _cache.pop(token, None)
        abort(410, "Lien expiré (10 min) — régénérez le rapport.")
    return send_file(
        io.BytesIO(entry["data"]),
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        as_attachment=True,
        download_name=entry["filename"],
    )


@app.route("/health")
def health():
    return jsonify({
        "status":           "ok",
        "template_present": os.path.exists(TEMPLATE_PATH),
        "template_name":    os.path.basename(TEMPLATE_PATH),
        "cache_entries":    len(_cache),
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=False)
