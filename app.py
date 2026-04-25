import io, os, threading, traceback, uuid
from datetime import datetime, timedelta
import pandas as pd
from flask import Flask, abort, jsonify, render_template_string, request, send_file

from generate_pptx_v3    import build_agency_pptx
from generate_html       import build_report_html
from generate_compitches import build_compitches_html

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "T21_HK_Agencies_Glass_v13.pptx")

_cache = {}
_lock  = threading.Lock()

def store_file(data, filename, content_type="application/octet-stream"):
    token  = str(uuid.uuid4())
    expiry = datetime.now() + timedelta(minutes=10)
    with _lock:
        _cache[token] = {"data": data, "filename": filename,
                         "content_type": content_type, "expiry": expiry}
    return token

def purge_cache():
    now = datetime.now()
    with _lock:
        dead = [k for k, v in _cache.items() if v["expiry"] < now]
        for k in dead: del _cache[k]

# ── UI ────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>RECMA Generator</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
:root{
  --bg:#0A0E1A;--surface:#111827;--surface2:#1C2333;--border:#1E293B;
  --accent:#38BDF8;--accent-dark:#2D5C54;--win:#10B981;--dep:#F43F5E;
  --text:#E2E8F0;--muted:#64748B;
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:'DM Sans',sans-serif;min-height:100vh;display:flex;flex-direction:column;align-items:center;}
header{width:100%;padding:14px 40px;border-bottom:1px solid var(--border);background:var(--surface);display:flex;align-items:center;gap:14px;}
.logo{font-family:'Syne',sans-serif;font-size:16px;font-weight:800;color:#fff;letter-spacing:.06em;}
.logo span{color:var(--accent);}
.logo-sub{font-size:11px;color:var(--muted);font-family:'DM Mono',monospace;letter-spacing:.08em;}
main{width:100%;max-width:700px;padding:48px 24px;display:flex;flex-direction:column;gap:28px;}
.tagline{text-align:center;}
.tagline h2{font-family:'Syne',sans-serif;font-size:28px;font-weight:700;color:#fff;letter-spacing:-.02em;}
.tagline h2 em{font-style:normal;color:var(--accent);}
.tagline p{margin-top:6px;font-size:13px;color:var(--muted);}

/* TABS */
.tabs{display:grid;grid-template-columns:1fr 1fr;background:var(--surface2);border-radius:12px;padding:5px;border:1px solid var(--border);}
.tab-btn{padding:11px 16px;border-radius:8px;border:none;background:transparent;color:var(--muted);cursor:pointer;font-family:'DM Sans',sans-serif;font-size:13px;font-weight:600;transition:all .2s;display:flex;align-items:center;justify-content:center;gap:7px;}
.tab-btn:hover{color:var(--text);}
.tab-btn.active{background:var(--surface);color:#fff;box-shadow:0 2px 8px rgba(0,0,0,.3);}
.tab-icon{font-size:16px;}
.tab-badge{font-size:9px;padding:2px 6px;border-radius:10px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;background:rgba(56,189,248,.15);color:var(--accent);}

/* PANELS */
.panel{display:none;flex-direction:column;gap:16px;}
.panel.active{display:flex;}
.card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:24px 28px;}
.card-label{font-size:11px;font-weight:500;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:14px;display:flex;align-items:center;gap:8px;}
.card-label::before{content:'';width:16px;height:1px;background:var(--accent);}

/* DROPZONE */
.dropzone{border:1.5px dashed var(--border);border-radius:8px;padding:28px;text-align:center;cursor:pointer;transition:border-color .2s,background .2s;position:relative;}
.dropzone:hover,.dropzone.over{border-color:var(--accent);background:rgba(56,189,248,.04);}
.dropzone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%;}
.dz-icon{font-size:22px;margin-bottom:8px;opacity:.6;}
.dz-label{font-size:14px;font-weight:500;}
.dz-hint{font-size:12px;color:var(--muted);margin-top:4px;font-family:'DM Mono',monospace;}
.dz-name{display:none;margin-top:10px;padding:6px 12px;background:rgba(16,185,129,.1);border:1px solid rgba(16,185,129,.3);border-radius:5px;font-size:12px;font-family:'DM Mono',monospace;color:var(--win);}

/* BUTTONS */
.format-choice{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:16px;}
.fmt-btn{padding:12px;border-radius:8px;border:1.5px solid var(--border);background:var(--surface2);color:var(--text);cursor:pointer;font-family:'DM Sans',sans-serif;font-size:13px;font-weight:500;transition:border-color .2s,background .2s;text-align:center;display:flex;flex-direction:column;gap:4px;align-items:center;}
.fmt-btn:hover:not(:disabled){border-color:var(--accent);background:rgba(56,189,248,.06);}
.fmt-btn:disabled{opacity:.3;cursor:not-allowed;}
.fmt-btn .fmt-icon{font-size:20px;}
.fmt-btn .fmt-label{font-size:11px;color:var(--muted);}
.full-btn{width:100%;padding:13px;border-radius:8px;border:1.5px solid var(--border);background:var(--surface2);color:var(--text);cursor:pointer;font-family:'DM Sans',sans-serif;font-size:13px;font-weight:600;transition:border-color .2s,background .2s;margin-top:16px;display:flex;align-items:center;justify-content:center;gap:8px;}
.full-btn:hover:not(:disabled){border-color:var(--accent);background:rgba(56,189,248,.06);}
.full-btn:disabled{opacity:.3;cursor:not-allowed;}
.full-btn .fmt-icon{font-size:18px;}

/* STATUS */
.status{margin-top:14px;font-size:13px;}
.status.loading{display:flex;align-items:center;gap:10px;color:var(--accent);}
.spinner{width:16px;height:16px;border:2px solid rgba(56,189,248,.2);border-top-color:var(--accent);border-radius:50%;animation:spin .7s linear infinite;}
@keyframes spin{to{transform:rotate(360deg)}}
.results{display:flex;flex-direction:column;gap:8px;margin-top:12px;}
.dl-btn{display:flex;align-items:center;gap:8px;padding:10px 16px;border-radius:6px;text-decoration:none;font-weight:600;font-size:13px;transition:opacity .2s;}
.dl-pptx{background:#2D5C54;color:#fff;}
.dl-html{background:#0369A1;color:#fff;}
.dl-comp{background:#7C3AED;color:#fff;}
.dl-btn:hover{opacity:.85;}
.err{color:var(--dep);margin-top:10px;font-family:'DM Mono',monospace;font-size:12px;}

/* THRESHOLD */
.threshold-row{margin-top:16px;padding:12px 14px;background:var(--surface2);border:1px solid var(--border);border-radius:8px;display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;}
.thr-label{font-size:12px;font-weight:500;color:var(--text);display:flex;flex-direction:column;gap:2px;}
.thr-hint{font-size:11px;color:var(--muted);font-weight:400;}
.thr-input-wrap{display:flex;align-items:center;gap:4px;background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:4px 10px;}
.thr-prefix,.thr-suffix{font-family:'DM Mono',monospace;font-size:13px;color:var(--muted);}
.thr-input{width:60px;background:transparent;border:none;color:var(--text);font-size:16px;font-weight:600;text-align:center;font-family:'DM Mono',monospace;outline:none;}

/* COLS DOC */
.col-doc{display:flex;flex-direction:column;gap:5px;}
.col-row{display:flex;align-items:center;gap:10px;padding:7px 10px;background:var(--surface2);border-radius:6px;font-size:12px;}
.col-name{font-family:'DM Mono',monospace;color:var(--accent);min-width:175px;}
.col-desc{color:var(--muted);}
.badge{font-size:9px;padding:2px 6px;border-radius:3px;font-weight:600;letter-spacing:.05em;text-transform:uppercase;margin-left:auto;}
.req{background:rgba(244,63,94,.15);color:var(--dep);}
.opt{background:rgba(100,116,139,.15);color:var(--muted);}

/* INFO BOX */
.info-box{background:rgba(56,189,248,.06);border:1px solid rgba(56,189,248,.2);border-radius:8px;padding:14px 16px;font-size:12.5px;line-height:1.6;color:var(--muted);}
.info-box strong{color:var(--accent);}
</style>
</head>
<body>
<header>
  <div>
    <div class="logo">RECMA <span>·</span> Generator</div>
    <div class="logo-sub">NBB T21 · Compitches T18</div>
  </div>
</header>
<main>
  <div class="tagline">
    <h2>Upload Excel → <em>Report</em></h2>
    <p>Génère automatiquement vos rapports RECMA depuis un fichier Excel.</p>
  </div>

  <!-- TABS -->
  <div class="tabs">
    <button class="tab-btn active" onclick="switchTab('nbb',this)">
      <span class="tab-icon">📊</span> NBB Report <span class="tab-badge">T21</span>
    </button>
    <button class="tab-btn" onclick="switchTab('comp',this)">
      <span class="tab-icon">🏆</span> Compitches <span class="tab-badge">T18</span>
    </button>
  </div>

  <!-- NBB PANEL -->
  <div class="panel active" id="panel-nbb">
    <div class="card">
      <div class="card-label">Colonnes Excel requises</div>
      <div class="col-doc">
        <div class="col-row"><span class="col-name">Agency</span><span class="col-desc">Nom de l'agence</span><span class="badge req">requis</span></div>
        <div class="col-row"><span class="col-name">NewBiz</span><span class="col-desc">WIN / DEPARTURE / RETENTION</span><span class="badge req">requis</span></div>
        <div class="col-row"><span class="col-name">Advertiser</span><span class="col-desc">Nom de l'annonceur</span><span class="badge req">requis</span></div>
        <div class="col-row"><span class="col-name">Integrated Spends</span><span class="col-desc">Budget $m (+ win, - departure)</span><span class="badge req">requis</span></div>
        <div class="col-row"><span class="col-name">Country</span><span class="col-desc">Marché</span><span class="badge opt">optionnel</span></div>
        <div class="col-row"><span class="col-name">Years</span><span class="col-desc">Année</span><span class="badge opt">optionnel</span></div>
      </div>
    </div>
    <div class="card">
      <div class="card-label">Générer le rapport NBB</div>
      <div class="dropzone" id="dz-nbb">
        <input type="file" id="fi-nbb" accept=".xlsx,.xls">
        <div class="dz-icon">📊</div>
        <div class="dz-label">Glissez votre Excel ici ou cliquez</div>
        <div class="dz-hint">.xlsx ou .xls · max 20 MB</div>
        <div class="dz-name" id="dzName-nbb"></div>
      </div>
      <div class="threshold-row">
        <label class="thr-label" for="threshold">
          Seuil wins/deps dans le tableau
          <span class="thr-hint">Filtre les mouvements en dessous de ce montant</span>
        </label>
        <div class="thr-input-wrap">
          <span class="thr-prefix">±</span>
          <input type="number" id="threshold" class="thr-input" value="5" min="0" max="500" step="1">
          <span class="thr-suffix">$m</span>
        </div>
      </div>
      <div class="format-choice">
        <button class="fmt-btn" id="btnPptx" disabled>
          <span class="fmt-icon">📑</span><strong>PPTX</strong>
          <span class="fmt-label">Présentation PowerPoint</span>
        </button>
        <button class="fmt-btn" id="btnHtml" disabled>
          <span class="fmt-icon">🌐</span><strong>HTML éditable</strong>
          <span class="fmt-label">Rapport web · Rich Text · Export PDF</span>
        </button>
      </div>
      <div id="status-nbb" class="status"></div>
    </div>
  </div>

  <!-- COMPITCHES PANEL -->
  <div class="panel" id="panel-comp">
    <div class="card">
      <div class="card-label">Colonnes Excel requises</div>
      <div class="col-doc">
        <div class="col-row"><span class="col-name">Agency</span><span class="col-desc">Nom de l'agence</span><span class="badge req">requis</span></div>
        <div class="col-row"><span class="col-name">NewBiz</span><span class="col-desc">WIN / DEPARTURE / RETENTION</span><span class="badge req">requis</span></div>
        <div class="col-row"><span class="col-name">Advertiser</span><span class="col-desc">Nom de l'annonceur</span><span class="badge req">requis</span></div>
        <div class="col-row"><span class="col-name">Integrated Spends</span><span class="col-desc">Budget $m</span><span class="badge req">requis</span></div>
        <div class="col-row"><span class="col-name">Move ?</span><span class="col-desc">Local / Global / Regional</span><span class="badge opt">optionnel</span></div>
        <div class="col-row"><span class="col-name">Pitch participation ?</span><span class="col-desc">Yes / No</span><span class="badge opt">optionnel</span></div>
        <div class="col-row"><span class="col-name">Incumbent</span><span class="col-desc">Agence précédente</span><span class="badge opt">optionnel</span></div>
      </div>
      <div class="info-box" style="margin-top:14px">
        <strong>Mode édition inline</strong> — Le HTML généré inclut un bouton <strong>✏️ EDIT</strong>
        pour modifier directement dans le navigateur : texte, couleurs, ajouter/supprimer des lignes.
        Puis <strong>💾 Export HTML</strong> pour sauvegarder.
      </div>
    </div>
    <div class="card">
      <div class="card-label">Générer le rapport Compitches</div>
      <div class="dropzone" id="dz-comp">
        <input type="file" id="fi-comp" accept=".xlsx,.xls">
        <div class="dz-icon">🏆</div>
        <div class="dz-label">Glissez votre Excel ici ou cliquez</div>
        <div class="dz-hint">.xlsx ou .xls · max 20 MB</div>
        <div class="dz-name" id="dzName-comp"></div>
      </div>
      <button class="full-btn" id="btnComp" disabled>
        <span class="fmt-icon">🌐</span>
        <strong>Générer Compitches HTML</strong>
      </button>
      <div id="status-comp" class="status"></div>
    </div>
  </div>

</main>

<script>
// ── TABS ──────────────────────────────────────────────────────
function switchTab(tab, btn) {
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('panel-' + tab).classList.add('active');
}

// ── DROPZONES ────────────────────────────────────────────────
function setupDz(dzId, fiId, nameId, btns) {
  const dz = document.getElementById(dzId);
  const fi = document.getElementById(fiId);
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('over'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('over'));
  dz.addEventListener('drop', e => {
    e.preventDefault(); dz.classList.remove('over');
    if (e.dataTransfer.files[0]) { fi.files = e.dataTransfer.files; onFile(e.dataTransfer.files[0], nameId, btns); }
  });
  fi.addEventListener('change', () => { if (fi.files[0]) onFile(fi.files[0], nameId, btns); });
}

function onFile(f, nameId, btns) {
  const n = document.getElementById(nameId);
  n.style.display = 'block'; n.textContent = '✓ ' + f.name;
  btns.forEach(id => { const b = document.getElementById(id); if(b) b.disabled = false; });
}

setupDz('dz-nbb','fi-nbb','dzName-nbb',['btnPptx','btnHtml']);
setupDz('dz-comp','fi-comp','dzName-comp',['btnComp']);

// ── GENERATE ─────────────────────────────────────────────────
async function generate(fiId, format, statusId, extraFields) {
  const fi = document.getElementById(fiId);
  if (!fi.files[0]) return;
  const st = document.getElementById(statusId);
  st.className = 'status loading';
  st.innerHTML = '<div class="spinner"></div><span>Génération en cours…</span>';
  ['btnPptx','btnHtml','btnComp'].forEach(id => { const b=document.getElementById(id); if(b) b.disabled=true; });

  const fd = new FormData();
  fd.append('file', fi.files[0]);
  fd.append('format', format);
  if (extraFields) Object.entries(extraFields).forEach(([k,v]) => fd.append(k,v));

  try {
    const res = await fetch('/generate', { method: 'POST', body: fd });
    const d   = await res.json();
    if (d.status === 'success') {
      st.className = 'status';
      const cls  = format === 'pptx' ? 'dl-pptx' : (format === 'compitches' ? 'dl-comp' : 'dl-html');
      const icon = format === 'pptx' ? '📑' : '🌐';
      st.innerHTML = `<div class="results"><a class="dl-btn ${cls}" href="${d.download_url}" target="_blank">${icon} Télécharger ${d.filename}</a></div>`;
    } else {
      st.className = 'status';
      st.innerHTML = `<div class="err">❌ ${d.error}</div>`;
    }
  } catch(e) {
    st.className = 'status';
    st.innerHTML = `<div class="err">❌ Erreur réseau</div>`;
  }
  ['btnPptx','btnHtml','btnComp'].forEach(id => { const b=document.getElementById(id); if(b) b.disabled=false; });
  // re-check if files are loaded
  if(!document.getElementById('fi-nbb').files[0]) { ['btnPptx','btnHtml'].forEach(id=>{ const b=document.getElementById(id); if(b) b.disabled=true; }); }
  if(!document.getElementById('fi-comp').files[0]) { const b=document.getElementById('btnComp'); if(b) b.disabled=true; }
}

document.getElementById('btnPptx').addEventListener('click', () => generate('fi-nbb','pptx','status-nbb',{threshold: document.getElementById('threshold').value}));
document.getElementById('btnHtml').addEventListener('click', () => generate('fi-nbb','html','status-nbb',{threshold: document.getElementById('threshold').value}));
document.getElementById('btnComp').addEventListener('click', () => generate('fi-comp','compitches','status-comp',{}));
</script>
</body>
</html>"""

# ── ROUTES ───────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/generate", methods=["POST"])
def generate():
    try:
        file = request.files.get("file")
        fmt  = request.form.get("format", "pptx")
        if not file:
            return jsonify({"status": "error", "error": "Fichier Excel manquant"}), 400

        df = pd.read_excel(file)
        required = ["Agency", "NewBiz", "Advertiser", "Integrated Spends"]
        missing  = [c for c in required if c not in df.columns]
        if missing:
            return jsonify({"status": "error",
                            "error": f"Colonnes manquantes : {', '.join(missing)}"}), 400

        ts = datetime.now().strftime("%Y%m%d_%H%M")

        if fmt == "compitches":
            data     = build_compitches_html(df)
            filename = f"Compitches_{ts}.html"
            ctype    = "text/html"

        elif fmt == "html":
            threshold = float(request.form.get("threshold", 5.0))
            data      = build_report_html(df, threshold=threshold)
            filename  = f"NBB_Report_{ts}.html"
            ctype     = "text/html"

        else:
            data     = build_agency_pptx(df, TEMPLATE_PATH)
            filename = f"NBB_Report_{ts}.pptx"
            ctype    = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

        token = store_file(data, filename, ctype)
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
        abort(404, "Lien expiré.")
    return send_file(
        io.BytesIO(entry["data"]),
        mimetype=entry["content_type"],
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
