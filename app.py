import io, os, threading, traceback, uuid
from datetime import datetime, timedelta
import pandas as pd
from flask import Flask, abort, jsonify, render_template_string, request, send_file

from generate_pptx_v3 import build_agency_pptx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "T21_HK_Agencies_Glass_v13.pptx")

_cache = {}
_lock  = threading.Lock()

def store_file(data, filename):
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

HTML = r"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>NBB Generator</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  :root { --bg:#0A0E1A; --surface:#111827; --surface2:#1C2333; --border:#1E293B;
          --accent:#38BDF8; --win:#10B981; --dep:#F43F5E; --text:#E2E8F0; --muted:#64748B; }
  *, *::before, *::after { box-sizing:border-box; margin:0; padding:0; }
  body { background:var(--bg); color:var(--text); font-family:'DM Sans',sans-serif;
         min-height:100vh; display:flex; flex-direction:column; align-items:center; }
  header { width:100%; padding:18px 40px; border-bottom:1px solid var(--border);
           background:var(--surface); display:flex; align-items:center; gap:12px; }
  .dot { width:10px; height:10px; border-radius:50%; background:var(--accent);
         box-shadow:0 0 12px var(--accent); }
  header h1 { font-size:14px; font-weight:500; letter-spacing:.1em; text-transform:uppercase; }
  main { width:100%; max-width:640px; padding:60px 24px; display:flex; flex-direction:column; gap:28px; }
  .tagline { text-align:center; }
  .tagline h2 { font-size:28px; font-weight:300; color:#fff; letter-spacing:-.02em; }
  .tagline h2 em { font-style:normal; color:var(--accent); }
  .tagline p { margin-top:8px; font-size:13px; color:var(--muted); }
  .card { background:var(--surface); border:1px solid var(--border); border-radius:12px; padding:28px 32px; }
  .card-label { font-size:11px; font-weight:500; letter-spacing:.12em; text-transform:uppercase;
                color:var(--muted); margin-bottom:16px; display:flex; align-items:center; gap:8px; }
  .card-label::before { content:''; width:16px; height:1px; background:var(--accent); }
  .dropzone { border:1.5px dashed var(--border); border-radius:8px; padding:32px;
              text-align:center; cursor:pointer; transition:border-color .2s,background .2s;
              position:relative; }
  .dropzone:hover, .dropzone.over { border-color:var(--accent); background:rgba(56,189,248,.04); }
  .dropzone input { position:absolute; inset:0; opacity:0; cursor:pointer; width:100%; height:100%; }
  .dz-icon { font-size:24px; margin-bottom:8px; opacity:.6; }
  .dz-label { font-size:14px; font-weight:500; }
  .dz-hint { font-size:12px; color:var(--muted); margin-top:4px; font-family:'DM Mono',monospace; }
  .dz-name { display:none; margin-top:10px; padding:6px 12px; background:rgba(16,185,129,.1);
             border:1px solid rgba(16,185,129,.3); border-radius:5px;
             font-size:12px; font-family:'DM Mono',monospace; color:var(--win); }
  .btn { margin-top:20px; width:100%; padding:14px; background:var(--accent); color:#0A0E1A;
         border:none; border-radius:7px; font-size:15px; font-weight:600; cursor:pointer;
         transition:background .2s; }
  .btn:hover:not(:disabled) { background:#7DD3FC; }
  .btn:disabled { opacity:.4; cursor:not-allowed; }
  .status { margin-top:16px; font-size:13px; }
  .status.loading { display:flex; align-items:center; gap:10px; color:var(--accent); }
  .spinner { width:16px; height:16px; border:2px solid rgba(56,189,248,.2);
             border-top-color:var(--accent); border-radius:50%; animation:spin .7s linear infinite; }
  @keyframes spin { to { transform:rotate(360deg); } }
  .dl-link { display:inline-block; margin-top:12px; padding:10px 20px; background:var(--win);
             color:#fff; text-decoration:none; border-radius:6px; font-weight:600; font-size:13px; }
  .err { color:var(--dep); margin-top:10px; font-family:'DM Mono',monospace; font-size:12px; }
  .col-doc { display:flex; flex-direction:column; gap:6px; }
  .col-row { display:flex; align-items:center; gap:10px; padding:8px 12px;
             background:var(--surface2); border-radius:6px; font-size:12px; }
  .col-name { font-family:'DM Mono',monospace; color:var(--accent); min-width:180px; }
  .col-desc { color:var(--muted); }
  .badge { font-size:10px; padding:2px 6px; border-radius:3px; font-weight:600;
           letter-spacing:.05em; text-transform:uppercase; margin-left:auto; }
  .req  { background:rgba(244,63,94,.15); color:var(--dep); }
  .opt  { background:rgba(100,116,139,.15); color:var(--muted); }
</style>
</head>
<body>
<header>
  <div class="dot"></div>
  <h1>NBB Report Generator</h1>
</header>
<main>
  <div class="tagline">
    <h2>Upload Excel → <em>Presentation</em></h2>
    <p>Les slides 1–6 sont remplies automatiquement. Les agency cards (7+) sont générées dynamiquement.</p>
  </div>

  <div class="card">
    <div class="card-label">Colonnes Excel requises</div>
    <div class="col-doc">
      <div class="col-row"><span class="col-name">Agency</span><span class="col-desc">Nom de l'agence</span><span class="badge req">requis</span></div>
      <div class="col-row"><span class="col-name">NewBiz</span><span class="col-desc">WIN / DEPARTURE / RETENTION</span><span class="badge req">requis</span></div>
      <div class="col-row"><span class="col-name">Advertiser</span><span class="col-desc">Nom de l'annonceur</span><span class="badge req">requis</span></div>
      <div class="col-row"><span class="col-name">Integrated Spends</span><span class="col-desc">Budget $m (+ win, - departure)</span><span class="badge req">requis</span></div>
      <div class="col-row"><span class="col-name">Date of announcement</span><span class="col-desc">Date d'annonce</span><span class="badge opt">optionnel</span></div>
      <div class="col-row"><span class="col-name">Incumbent</span><span class="col-desc">Agence précédente</span><span class="badge opt">optionnel</span></div>
    </div>
  </div>

  <div class="card">
    <div class="card-label">Générer le rapport</div>
    <div class="dropzone" id="dz">
      <input type="file" id="fi" accept=".xlsx,.xls">
      <div class="dz-icon">📊</div>
      <div class="dz-label">Glissez votre Excel ici ou cliquez</div>
      <div class="dz-hint">.xlsx ou .xls · max 20 MB</div>
      <div class="dz-name" id="dzName"></div>
    </div>
    <button class="btn" id="btn" disabled>Générer la présentation</button>
    <div id="status"></div>
  </div>
</main>
<script>
const dz = document.getElementById('dz');
const fi = document.getElementById('fi');
const btn = document.getElementById('btn');
const st  = document.getElementById('status');

dz.addEventListener('dragover',  e => { e.preventDefault(); dz.classList.add('over'); });
dz.addEventListener('dragleave', () => dz.classList.remove('over'));
dz.addEventListener('drop', e => {
  e.preventDefault(); dz.classList.remove('over');
  if (e.dataTransfer.files[0]) { fi.files = e.dataTransfer.files; onFile(e.dataTransfer.files[0]); }
});
fi.addEventListener('change', () => { if (fi.files[0]) onFile(fi.files[0]); });

function onFile(f) {
  document.getElementById('dzName').style.display = 'block';
  document.getElementById('dzName').textContent = '✓ ' + f.name;
  btn.disabled = false;
}

btn.addEventListener('click', async () => {
  if (!fi.files[0]) return;
  btn.disabled = true;
  st.className = 'status loading';
  st.innerHTML = '<div class="spinner"></div><span>Génération en cours…</span>';
  const fd = new FormData();
  fd.append('file', fi.files[0]);
  try {
    const res  = await fetch('/generate', { method:'POST', body:fd });
    const data = await res.json();
    if (data.status === 'success') {
      st.className = 'status';
      st.innerHTML = `<a class="dl-link" href="${data.download_url}" target="_blank">⬇ Télécharger le PPTX</a>`;
    } else {
      st.className = 'status';
      st.innerHTML = `<div class="err">❌ ${data.error}</div>`;
    }
  } catch(e) {
    st.className = 'status';
    st.innerHTML = `<div class="err">❌ Erreur réseau</div>`;
  }
  btn.disabled = false;
});
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

        required = ["Agency", "NewBiz", "Advertiser", "Integrated Spends"]
        missing  = [c for c in required if c not in df.columns]
        if missing:
            return jsonify({"status": "error",
                            "error": f"Colonnes manquantes : {', '.join(missing)}"}), 400

        pptx_bytes = build_agency_pptx(df, TEMPLATE_PATH)

        ts       = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f"NBB_Report_{ts}.pptx"
        token    = store_file(pptx_bytes, filename)
        purge_cache()

        return jsonify({
            "status":       "success",
            "download_url": f"{request.host_url.rstrip('/')}/download/{token}",
            "filename":     filename,
            "agencies":     int(df["Agency"].nunique()),
        })
    except Exception as e:
        app.logger.error(traceback.format_exc())
        return jsonify({"status": "error", "error": str(e)}), 500

@app.route("/download/<token>")
def download(token):
    purge_cache()
    with _lock:
        entry = _cache.get(token)
    if not entry:
        abort(404, "Lien expiré. Régénérez le rapport.")
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
        "cache_entries":    len(_cache),
    })

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
