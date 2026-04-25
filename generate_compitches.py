"""
generate_compitches.py
Generates the Compitches HTML report from an Excel DataFrame.
Includes a full inline editor: add/delete rows, edit text, change background color.
"""
import pandas as pd

GROUP_MAP = {
    'INITIATIVE':'IPG Mediabrands','UM':'IPG Mediabrands','HEARTS & SCIENCE':'IPG Mediabrands',
    'HAVAS MEDIA':'Havas Media Network','HAVAS':'Havas Media Network',
    'ESSENCEMEDIACOM':'WPP Media','MINDSHARE':'WPP Media','WAVEMAKER':'WPP Media',
    'PHD':'Omnicom Media','OMD':'Omnicom Media',
    'SPARK FOUNDRY':'Publicis Media','STARCOM':'Publicis Media',
    'ZENITH':'Publicis Media','PUBLICIS MEDIA':'Publicis Media',
    'CARAT':'dentsu','IPROSPECT':'dentsu',
    'ARENA':'Independent',
}
GROUP_BG = {
    'IPG Mediabrands':'#F3E8FF','Havas Media Network':'#EDE9FE',
    'WPP Media':'#FEE2E2','Omnicom Media':'#FEF9C3',
    'Publicis Media':'#FFEDD5','dentsu':'#DCFCE7','Independent':'#F1F5F9',
}
GROUP_BD = {
    'IPG Mediabrands':'#9333EA','Havas Media Network':'#7C3AED',
    'WPP Media':'#DC2626','Omnicom Media':'#CA8A04',
    'Publicis Media':'#EA580C','dentsu':'#16A34A','Independent':'#64748B',
}
GRADE_CLS = {'A+':'grade-Ap','A':'grade-A','B+':'grade-Bp','B':'grade-B','C':'grade-C'}

def pts_win(sp):
    if sp >= 5: return '+3 pts','pos'
    if sp >= 1: return '+1 pt','pos'
    return '0 pt','neu'

def pts_dep(sp):
    sp = abs(sp)
    if sp >= 5: return '−3 pts','neg'
    if sp >= 1: return '−1 pt','neg'
    return '0 pt','neu'

def pts_ret(sp):
    if sp >= 5: return '+3 pts','gold'
    if sp >= 1: return '+1 pt','gold'
    return '0 pt','neu'

def grade_from_pts(pts):
    if pts >= 13: return 'A+'
    if pts >= 9:  return 'A'
    if pts >= 4:  return 'B+'
    if pts >= 1:  return 'B'
    if pts >= 0:  return 'B'
    if pts >= -2: return 'C'
    return 'C'

def fmt_sp(v):
    v = float(v)
    if v > 0:  return f'${v:.1f}m'
    if v < 0:  return f'−${abs(v):.1f}m'
    return '—'

def sign_cls(v):
    if v > 0: return 'pos'
    if v < 0: return 'neg'
    return 'neu'

def pts_sign(p):
    if '+' in str(p): return 'pos'
    if '−' in str(p) or '-' in str(p): return 'neg'
    return 'neu'


def build_compitches_html(df: pd.DataFrame) -> bytes:
    df.columns = [str(c).strip() for c in df.columns]
    df['Integrated Spends'] = pd.to_numeric(df['Integrated Spends'], errors='coerce').fillna(0)
    df['Agency'] = df['Agency'].astype(str).str.strip().str.upper()

    try:
        market = str(df['Country'].dropna().iloc[0]).title()
    except: market = 'Market'
    try:
        year = str(int(df['Years'].dropna().iloc[0]))
    except: year = '2025'

    # Build agency data
    agencies = []
    for ag in sorted(df['Agency'].dropna().unique()):
        sub = df[df['Agency'] == ag]
        wins_df = sub[sub['NewBiz'] == 'WIN'].sort_values('Integrated Spends', ascending=False)
        deps_df = sub[sub['NewBiz'] == 'DEPARTURE'].sort_values('Integrated Spends')
        rets_df = sub[sub['NewBiz'] == 'RETENTION'].sort_values('Integrated Spends', ascending=False)

        # Calculate RECMA points
        pts = 0
        for _, r in wins_df.iterrows():
            sp = float(r['Integrated Spends'])
            move = str(r.get('Move ?', 'Local'))
            pitch = str(r.get('Pitch participation ?', 'Yes'))
            if pitch.lower() == 'no': continue
            if move == 'Global': pts += 1 if sp >= 5 else (1 if sp >= 1 else 0)
            else: pts += 3 if sp >= 5 else (1 if sp >= 1 else 0)
        for _, r in deps_df.iterrows():
            sp = abs(float(r['Integrated Spends']))
            move = str(r.get('Move ?', 'Local'))
            if move == 'Global': pts -= 1 if sp >= 5 else (1 if sp >= 1 else 0)
            else: pts -= 3 if sp >= 5 else (1 if sp >= 1 else 0)
        for _, r in rets_df.iterrows():
            sp = float(r['Integrated Spends'])
            move = str(r.get('Move ?', 'Local'))
            if move != 'Global': pts += 3 if sp >= 5 else (1 if sp >= 1 else 0)

        group = GROUP_MAP.get(ag, 'Independent')
        agencies.append({
            'name': ag, 'group': group,
            'bg': GROUP_BG.get(group, '#F1F5F9'),
            'border': GROUP_BD.get(group, '#64748B'),
            'pts': pts, 'grade': grade_from_pts(pts),
            'wins_n': len(wins_df), 'deps_n': len(deps_df), 'rets_n': len(rets_df),
            'wins': list(wins_df.to_dict('records')),
            'deps': list(deps_df.to_dict('records')),
            'rets': list(rets_df.to_dict('records')),
        })

    agencies.sort(key=lambda x: -x['pts'])
    for i, a in enumerate(agencies): a['rank'] = i + 1

    # ── Ranking table rows ───────────────────────────────────────
    ranking_rows = ''
    for a in agencies:
        bg = a['bg']; bd = a['border']
        pb = a['pts']
        pb_cls = sign_cls(pb)
        pb_str = f'+{pb}' if pb > 0 else str(pb)
        gcls = GRADE_CLS.get(a['grade'], 'grade-B')
        top_win = str(a['wins'][0].get('Advertiser', '—')).title() if a['wins'] else '—'
        top_dep = str(a['deps'][0].get('Advertiser', '—')).title() if a['deps'] else '—'
        tw_sp   = fmt_sp(a['wins'][0]['Integrated Spends']) if a['wins'] else ''
        td_sp   = fmt_sp(a['deps'][0]['Integrated Spends']) if a['deps'] else ''
        ranking_rows += f'''
        <tr style="background:{bg};border-left:3px solid {bd}" data-editable="row" data-id="rank-{a["name"]}">
          <td class="td-rk" data-editable="text">{a["rank"]}</td>
          <td class="td-ag"><strong data-editable="text">{a["name"].title()}</strong><br>
            <span class="td-grp" data-editable="text">{a["group"]}</span></td>
          <td class="td-grade {gcls}" data-editable="text">{a["grade"]}</td>
          <td class="td-pts {pb_cls}" data-editable="text">{pb_str}</td>
          <td class="td-wl" data-editable="text"><span class="w">{a["wins_n"]}</span> – <span class="l">{a["deps_n"]}</span></td>
          <td class="td-mv" data-editable="text">{top_win} {tw_sp}</td>
          <td class="td-mv neg-txt" data-editable="text">{top_dep + (" " + td_sp if top_dep != "—" else "")}</td>
        </tr>'''

    def clean(v):
        s = str(v).strip()
        return '' if s.lower() in ('nan','none','') else s

    # ── Agency detail sections ───────────────────────────────────
    def move_card(r, kind):
        adv       = clean(r.get('Advertiser', '')) or '—'
        adv       = adv.title()
        sp        = float(r.get('Integrated Spends', 0))
        move      = clean(r.get('Move ?', ''))
        inc       = clean(r.get('Incumbent', ''))
        contender = clean(r.get('Contender', ''))
        remark    = clean(r.get('Remarks', ''))
        assign    = clean(r.get('Assignment', ''))

        if kind == 'win':
            p, pc = pts_win(sp)
            cls = 'win-item'; border = 'var(--pos)'
            sp_cls = 'pos'; sp_str = fmt_sp(sp)
            tag = 'KEY PITCH' if sp >= 5 else ('OTHER PITCH' if sp >= 1 else '')
            tag_style = 'background:#1E3A35;color:#fff'
        elif kind == 'dep':
            p, pc = pts_dep(sp)
            cls = 'dep-item'; border = 'var(--neg)'
            sp_cls = 'neg'; sp_str = fmt_sp(sp)
            tag = 'KEY PITCH' if abs(sp) >= 5 else ('OTHER PITCH' if abs(sp) >= 1 else '')
            tag_style = 'background:#7F1D1D;color:#FCA5A5'
        else:
            p, pc = pts_ret(sp)
            cls = 'ret-item'; border = 'var(--gold)'
            sp_cls = 'gold'; sp_str = fmt_sp(sp)
            tag = 'RETENTION'; tag_style = 'background:#92400E;color:#fff'

        tag_html  = f'<span class="mv-tag-badge" style="{tag_style}" data-editable="text">{tag}</span>' if tag else ''
        move_html = f'<span class="mv-scope" data-editable="text">{move}</span>' if move else ''
        assign_html = f'<span class="mv-scope" data-editable="text">{assign}</span>' if assign else ''

        # Info line: From / Contenders — always shown, editable even if empty
        inc_val  = f'From: {inc}' if inc else 'From: —'
        cont_val = contender if contender else '—'
        rem_val  = remark if remark else ''

        info_line = f'''<div class="mv-info">
          <span class="mv-info-item"><span class="mv-info-lbl">From</span>
            <span class="mv-info-val" data-editable="text">{inc if inc else "—"}</span></span>
          <span class="mv-info-item"><span class="mv-info-lbl">Contenders</span>
            <span class="mv-info-val" data-editable="text">{cont_val}</span></span>
          {f'<span class="mv-info-item mv-info-note" data-editable="text">{rem_val}</span>' if rem_val else ''}
        </div>'''

        return f'''<div class="mv-item {cls}" data-editable="card" style="border-left:3px solid {border}">
          <button class="edit-del-btn" title="Delete this row" onclick="deleteCard(this)">✕</button>
          <div class="mv-main">
            <span class="mv-adv" data-editable="text">{adv}</span>
            <span class="mv-sp {sp_cls}" data-editable="text">{sp_str}</span>
          </div>
          <div class="mv-meta">{tag_html}{move_html}{assign_html}<span class="mv-pts {pc}" data-editable="text">{p}</span></div>
          {info_line}
        </div>'''

    detail_sections = ''
    nav_ag_tabs = ''
    for i, a in enumerate(agencies):
        sec_idx = i + 2
        bg = a['bg']; bd = a['border']
        pb = a['pts']
        pb_cls = sign_cls(pb)
        pb_str = f'+{pb}' if pb > 0 else str(pb)
        gcls = GRADE_CLS.get(a['grade'], 'grade-B')

        wins_html = ''.join(move_card(r,'win') for r in a['wins']) or '<div class="mv-empty">No wins recorded</div>'
        deps_html = ''.join(move_card(r,'dep') for r in a['deps']) or '<div class="mv-empty">No departures</div>'
        rets_html = ''.join(move_card(r,'ret') for r in a['rets'])
        ret_col   = f'<div class="ag-col" data-col="ret"><div class="col-hdr ret-hdr">↺ Retentions</div>{rets_html}<button class="add-row-btn" onclick="addCard(this,\'ret\')">+ Add row</button></div>' if a['rets'] else ''
        grid_cls  = 'three' if a['rets'] else ''

        nav_ag_tabs += f'<a href="#section-{sec_idx}" class="nav-tab"><span class="nav-num">{sec_idx+1:02d}</span>{a["name"].title()}</a>'

        detail_sections += f'''
<section id="section-{sec_idx}" class="page ag-page" style="border-top:2px solid var(--border)" data-agency="{a["name"]}">
  <div class="ag-head" style="background:{bg};border-left:4px solid {bd}">
    <div>
      <div class="ag-name" data-editable="text">{a["name"].title()}</div>
      <div class="ag-group" data-editable="text">{a["group"]}</div>
    </div>
    <div class="ag-kpis">
      <div class="kpi-grade {gcls}" data-editable="text">{a["grade"]}</div>
      <div class="kpi-pill {pb_cls}" data-editable="text">{pb_str} pts</div>
      <div class="kpi-pill neu" data-editable="text">{a["wins_n"]}W · {a["deps_n"]}D</div>
      <button class="edit-color-btn" title="Change header color" onclick="pickColor(this)">🎨</button>
    </div>
  </div>
  <div class="ag-cols {grid_cls}">
    <div class="ag-col" data-col="win">
      <div class="col-hdr win-hdr">↑ Wins</div>
      {wins_html}
      <button class="add-row-btn" onclick="addCard(this,'win')">+ Add row</button>
    </div>
    <div class="ag-col" data-col="dep">
      <div class="col-hdr dep-hdr">↓ Departures</div>
      {deps_html}
      <button class="add-row-btn" onclick="addCard(this,'dep')">+ Add row</button>
    </div>
    {ret_col}
  </div>
</section>'''

    # ── CSS + JS editor ──────────────────────────────────────────
    EDITOR_CSS = """
/* ── EDITOR TOOLBAR ── */
#edit-toolbar {
  position: fixed; bottom: 24px; right: 24px; z-index: 9999;
  display: flex; flex-direction: column; gap: 8px; align-items: flex-end;
}
#edit-toggle {
  background: #0F172A; color: #fff; border: 1px solid #334155;
  padding: 10px 18px; border-radius: 24px; cursor: pointer;
  font-family: var(--ff-m); font-size: 12px; font-weight: 600;
  letter-spacing: .06em; box-shadow: 0 4px 20px rgba(0,0,0,.4);
  transition: background .2s;
}
#edit-toggle:hover { background: #1E293B; }
#edit-toggle.active { background: #2D5C54; border-color: #38BDF8; color: #38BDF8; }
#export-btn {
  background: #2D5C54; color: #fff; border: none;
  padding: 10px 18px; border-radius: 24px; cursor: pointer;
  font-family: var(--ff-m); font-size: 12px; font-weight: 600;
  letter-spacing: .06em; box-shadow: 0 4px 20px rgba(0,0,0,.3);
  display: none;
}
#export-btn:hover { background: #38BDF8; color: #0F172A; }

/* ── EDIT MODE ── */
body.edit-mode [data-editable="text"] {
  cursor: text; border-radius: 3px; outline: 1px dashed rgba(56,189,248,.4);
  min-width: 20px; display: inline-block;
}
body.edit-mode [data-editable="text"]:hover { outline-color: #38BDF8; background: rgba(56,189,248,.06); }
body.edit-mode [data-editable="text"]:focus { outline: 2px solid #38BDF8; background: rgba(56,189,248,.1); }
body.edit-mode [data-editable="card"] { position: relative; }
body.edit-mode .edit-del-btn { display: flex; }
body.edit-mode .add-row-btn { display: flex; }
body.edit-mode .edit-color-btn { display: flex; }

/* ── BUTTONS HIDDEN BY DEFAULT ── */
.edit-del-btn {
  display: none; position: absolute; top: 6px; right: 6px;
  width: 20px; height: 20px; border-radius: 50%; border: none;
  background: #DC2626; color: #fff; font-size: 10px; font-weight: 700;
  cursor: pointer; align-items: center; justify-content: center; z-index: 10;
  line-height: 1;
}
.add-row-btn {
  display: none; width: 100%; margin-top: 6px; padding: 6px;
  border: 1px dashed #38BDF8; border-radius: 6px; background: transparent;
  color: #38BDF8; font-size: 11px; font-weight: 600; cursor: pointer;
  letter-spacing: .05em; align-items: center; justify-content: center; gap: 4px;
}
.add-row-btn:hover { background: rgba(56,189,248,.07); }
.edit-color-btn {
  display: none; padding: 4px 8px; border-radius: 6px; border: 1px solid var(--border);
  background: var(--surface); cursor: pointer; font-size: 13px;
}

/* ── COLOR PICKER POPUP ── */
#color-picker-popup {
  position: fixed; z-index: 99999; background: #1C2333;
  border: 1px solid #334155; border-radius: 10px; padding: 12px;
  box-shadow: 0 8px 32px rgba(0,0,0,.5); display: none;
}
#color-picker-popup.show { display: block; }
#color-picker-popup h4 { font-size: 11px; color: #64748B; letter-spacing: .08em; text-transform: uppercase; margin-bottom: 10px; }
.cp-swatches { display: grid; grid-template-columns: repeat(5, 28px); gap: 6px; margin-bottom: 10px; }
.cp-swatch { width: 28px; height: 28px; border-radius: 5px; cursor: pointer; border: 2px solid transparent; transition: border-color .15s; }
.cp-swatch:hover { border-color: #38BDF8; }
.cp-custom { display: flex; align-items: center; gap: 6px; }
.cp-custom label { font-size: 11px; color: #64748B; }
.cp-custom input[type=color] { width: 32px; height: 32px; border: none; border-radius: 4px; cursor: pointer; background: none; padding: 0; }
#cp-close { margin-top: 8px; width: 100%; padding: 6px; border: 1px solid #334155; border-radius: 6px; background: transparent; color: #64748B; font-size: 11px; cursor: pointer; }
#cp-close:hover { color: #fff; border-color: #64748B; }
"""

    EDITOR_JS = """
// ── EDITOR ───────────────────────────────────────────────────────
let editMode = false;
let colorTarget = null;

const SWATCHES = [
  '#F3E8FF','#EDE9FE','#DCFCE7','#FEF9C3','#FFEDD5',
  '#FEE2E2','#F1F5F9','#D1FAE5','#FEF3C7','#FFE4E6',
  '#E0F2FE','#F0FDF4','#FFF7ED','#FEFCE8','#F8FAFC',
  '#1E3A35','#0F172A','#1C2333','#2D5C54','#334155',
];

function buildColorPicker() {
  const popup = document.getElementById('color-picker-popup');
  const sw = popup.querySelector('.cp-swatches');
  sw.innerHTML = '';
  SWATCHES.forEach(c => {
    const d = document.createElement('div');
    d.className = 'cp-swatch';
    d.style.background = c;
    d.title = c;
    d.onclick = () => applyColor(c);
    sw.appendChild(d);
  });
}

function pickColor(btn) {
  colorTarget = btn.closest('.ag-head');
  const popup = document.getElementById('color-picker-popup');
  const rect = btn.getBoundingClientRect();
  popup.style.top  = (rect.bottom + 8 + window.scrollY) + 'px';
  popup.style.left = Math.max(8, rect.right - 280 + window.scrollX) + 'px';
  popup.classList.add('show');
}

function applyColor(color) {
  if (!colorTarget) return;
  colorTarget.style.background = color;
  document.getElementById('color-picker-popup').classList.remove('show');
  colorTarget = null;
}

function deleteCard(btn) {
  if (confirm('Delete this row?')) btn.closest('.mv-item').remove();
}

function addCard(btn, kind) {
  const col = btn.closest('.ag-col');
  const bgMap = { win: '#F0FDF4', dep: '#FEF2F2', ret: '#FFFBEB' };
  const borderMap = { win: 'var(--pos)', dep: 'var(--neg)', ret: 'var(--gold)' };
  const clsMap  = { win: 'win-item', dep: 'dep-item', ret: 'ret-item' };
  const spCls   = { win: 'pos', dep: 'neg', ret: 'gold' };
  const tagMap  = { win: 'KEY PITCH', dep: 'KEY PITCH', ret: 'RETENTION' };
  const tagSty  = { win: 'background:#1E3A35;color:#fff', dep: 'background:#1E3A35;color:#fff', ret: 'background:#92400E;color:#fff' };
  const ptsDef  = { win: '+3 pts', dep: '−3 pts', ret: '+3 pts' };
  const ptsC    = { win: 'pos', dep: 'neg', ret: 'gold' };

  const card = document.createElement('div');
  card.className = `mv-item ${clsMap[kind]}`;
  card.setAttribute('data-editable','card');
  card.style.borderLeft = `3px solid ${borderMap[kind]}`;
  card.innerHTML = `
    <button class="edit-del-btn" title="Delete" onclick="deleteCard(this)">✕</button>
    <div class="mv-main">
      <span class="mv-adv" data-editable="text" contenteditable="true">New Advertiser</span>
      <span class="mv-sp ${spCls[kind]}" data-editable="text" contenteditable="true">$0.0m</span>
    </div>
    <div class="mv-meta">
      <span class="mv-tag-badge" style="${tagSty[kind]}" data-editable="text" contenteditable="true">${tagMap[kind]}</span>
      <span class="mv-scope" data-editable="text" contenteditable="true">Local</span>
      <span class="mv-pts ${ptsC[kind]}" data-editable="text" contenteditable="true">${ptsDef[kind]}</span>
    </div>`;
  col.insertBefore(card, btn);
}

function toggleEditMode() {
  editMode = !editMode;
  document.body.classList.toggle('edit-mode', editMode);
  const btn = document.getElementById('edit-toggle');
  btn.textContent = editMode ? '✏️ EDITING ON' : '✏️ EDIT';
  btn.classList.toggle('active', editMode);
  document.getElementById('export-btn').style.display = editMode ? 'block' : 'none';

  // Make all [data-editable="text"] contenteditable
  document.querySelectorAll('[data-editable="text"]').forEach(el => {
    el.contentEditable = editMode ? 'true' : 'false';
  });
}

function exportHTML() {
  // Temporarily disable edit mode visuals for clean export
  document.body.classList.remove('edit-mode');
  document.querySelectorAll('[data-editable="text"]').forEach(el => el.contentEditable = 'false');
  document.querySelectorAll('.edit-del-btn,.add-row-btn,.edit-color-btn').forEach(el => el.style.display = 'none');
  document.getElementById('edit-toolbar').style.display = 'none';
  document.getElementById('color-picker-popup').classList.remove('show');

  const html = '<!DOCTYPE html>\\n' + document.documentElement.outerHTML;
  const blob = new Blob([html], {type: 'text/html'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'compitches-edited.html';
  a.click();

  // Restore
  setTimeout(() => {
    document.getElementById('edit-toolbar').style.display = 'flex';
    if (editMode) {
      document.body.classList.add('edit-mode');
      document.querySelectorAll('[data-editable="text"]').forEach(el => el.contentEditable = 'true');
      document.querySelectorAll('.edit-del-btn,.add-row-btn,.edit-color-btn').forEach(el => el.style.display = '');
    }
  }, 500);
}

// Init
document.addEventListener('DOMContentLoaded', () => {
  buildColorPicker();
  document.getElementById('color-picker-popup').querySelector('#cp-close').onclick = () =>
    document.getElementById('color-picker-popup').classList.remove('show');
  document.getElementById('color-picker-popup').querySelector('.cp-custom input').oninput = function() {
    applyColor(this.value);
  };
  // Close popup on outside click
  document.addEventListener('click', e => {
    const popup = document.getElementById('color-picker-popup');
    if (!popup.contains(e.target) && !e.target.classList.contains('edit-color-btn')) {
      popup.classList.remove('show');
    }
  });
});

// Nav scroll spy
const tabs=[...document.querySelectorAll('.nav-tab')];
const secs=[...document.querySelectorAll('section[id]')];
const io=new IntersectionObserver(e=>{e.forEach(x=>{if(x.isIntersecting){const i=secs.indexOf(x.target);tabs.forEach(t=>t.classList.remove('active'));if(tabs[i])tabs[i].classList.add('active');}});},{threshold:0.2});
secs.forEach(s=>io.observe(s));
tabs.forEach(t=>t.addEventListener('click',e=>{e.preventDefault();document.querySelector(t.getAttribute('href'))?.scrollIntoView({behavior:'smooth'})}));
"""

    # ── Full HTML ─────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{market} · Compitches {year} · RECMA</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600&family=DM+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
:root {{
  --bg:#F7F8FC;--surface:#FFFFFF;--border:#E2E8F0;--nav-bg:#0F172A;
  --accent:#2D5C54;--accent2:#38BDF8;
  --pos:#059669;--neg:#DC2626;--gold:#D97706;
  --text:#1E293B;--muted:#64748B;--max:1100px;
  --ff-h:'Syne',sans-serif;--ff-b:'DM Sans',sans-serif;--ff-m:'DM Mono',monospace;
}}
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
html{{scroll-behavior:smooth}}
body{{background:var(--bg);color:var(--text);font-family:var(--ff-b);font-size:14px;line-height:1.5}}
.top-nav{{position:sticky;top:0;z-index:100;background:var(--nav-bg);box-shadow:0 2px 16px rgba(0,0,0,.35)}}
.nav-inner{{max-width:var(--max);margin:0 auto;display:flex;overflow-x:auto;scrollbar-width:none;align-items:stretch}}
.nav-inner::-webkit-scrollbar{{display:none}}
.nav-brand{{font-family:var(--ff-h);font-size:12px;font-weight:800;color:#fff;letter-spacing:.06em;padding:0 16px;border-right:1px solid #1E293B;display:flex;align-items:center;white-space:nowrap;flex-shrink:0}}
.nav-brand span{{color:var(--accent2)}}
.nav-tab{{display:flex;align-items:center;gap:5px;padding:13px 12px;color:#64748B;text-decoration:none;font-size:11.5px;font-weight:500;white-space:nowrap;border-bottom:2px solid transparent;transition:color .15s}}
.nav-tab:hover{{color:#fff}} .nav-tab.active{{color:#fff;border-bottom-color:var(--accent2)}}
.nav-num{{font-family:var(--ff-h);font-size:12px;font-weight:800;opacity:.4}}
.page{{max-width:var(--max);margin:0 auto;padding:2.5rem 1.5rem}}
.sec-label{{font-family:var(--ff-m);font-size:10px;letter-spacing:.14em;text-transform:uppercase;color:var(--muted);margin-bottom:.4rem}}
.sec-title{{font-family:var(--ff-h);font-size:clamp(1.4rem,3vw,2rem);font-weight:700;color:var(--accent);line-height:1.15;margin-bottom:1.5rem}}
.sec-title .sub{{color:var(--accent2);font-weight:400;font-size:.7em}}
.cover-hero{{background:var(--accent);padding:3.5rem 3rem 3rem}}
.cover-eyebrow{{font-family:var(--ff-m);font-size:10px;letter-spacing:.2em;text-transform:uppercase;color:rgba(255,255,255,.4);margin-bottom:1rem}}
.cover-title{{font-family:var(--ff-h);font-size:clamp(2.2rem,6vw,3.8rem);font-weight:800;color:#fff;line-height:1;margin-bottom:.6rem}}
.cover-year{{font-family:var(--ff-h);font-size:1.2rem;font-weight:400;color:var(--accent2)}}
.method-wrap{{background:var(--surface);padding:2rem 3rem 2.5rem;border-top:1px solid var(--border)}}
.method-label{{font-family:var(--ff-m);font-size:10px;letter-spacing:.15em;text-transform:uppercase;color:var(--muted);margin-bottom:1.25rem}}
.method-grid{{display:grid;grid-template-columns:repeat(4,1fr);border:1px solid var(--border);border-radius:10px;overflow:hidden;margin-bottom:2rem}}
.method-cell{{padding:1.25rem 1rem;border-right:1px solid var(--border);background:var(--surface)}}
.method-cell:last-child{{border-right:none}}
.mc-type{{font-family:var(--ff-h);font-size:.62rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;margin-bottom:.4rem}}
.mc-cond{{font-size:11.5px;color:var(--muted);line-height:1.5;margin-bottom:.85rem}}
.mc-pts{{display:flex;flex-direction:column;gap:5px}}
.pt-row{{display:flex;align-items:baseline;gap:.4rem;font-family:var(--ff-m);font-size:14px;font-weight:700}}
.pt-row .lbl{{font-family:var(--ff-b);font-size:10.5px;font-weight:400;color:var(--muted)}}
.grade-Ap{{color:#059669}} .grade-A{{color:#16A34A}} .grade-Bp{{color:#2563EB}} .grade-B{{color:#3B82F6}} .grade-C{{color:#DC2626}}
.rank-wrap{{overflow-x:auto;margin-top:1.25rem}}
table.rtable{{width:100%;border-collapse:collapse;font-size:13px;min-width:680px}}
.rtable thead th{{font-family:var(--ff-m);font-size:10px;letter-spacing:.07em;text-transform:uppercase;color:var(--muted);padding:8px 10px;border-bottom:2px solid var(--border);background:var(--bg);white-space:nowrap}}
.rtable tbody td{{padding:9px 10px;border-bottom:1px solid var(--border);vertical-align:middle}}
.rtable tbody tr:hover{{filter:brightness(.97)}}
.td-rk{{font-family:var(--ff-m);color:var(--muted);font-size:11px;width:28px}}
.td-ag{{min-width:130px}} .td-grp{{font-size:10.5px;color:var(--muted);font-weight:400}}
.td-grade{{font-family:var(--ff-h);font-weight:800;font-size:1.2rem;text-align:center;width:44px}}
.td-pts{{font-family:var(--ff-m);font-weight:700;font-size:14px;text-align:center;width:50px}}
.td-wl{{text-align:center;font-size:12px;white-space:nowrap}}
.td-wl .w{{color:var(--pos);font-weight:700}} .td-wl .l{{color:var(--neg);font-weight:700}}
.td-mv{{font-size:11.5px;color:var(--muted);max-width:180px}} .td-mv.neg-txt{{color:var(--neg)}}
.group-legend{{display:flex;flex-wrap:wrap;gap:.5rem;margin-bottom:1.5rem}}
.gleg{{display:flex;align-items:center;gap:.35rem;font-size:11px;padding:3px 10px;border-radius:20px;border:1px solid;font-weight:500}}
.ag-page{{padding-top:2rem;padding-bottom:2rem}}
.ag-head{{display:flex;align-items:center;justify-content:space-between;gap:1rem;flex-wrap:wrap;padding:1rem 1.5rem;border-radius:10px 10px 0 0}}
.ag-name{{font-family:var(--ff-h);font-size:1.2rem;font-weight:800;color:var(--text)}}
.ag-group{{font-size:11px;color:var(--muted);margin-top:2px}}
.ag-kpis{{display:flex;gap:.5rem;flex-wrap:wrap;align-items:center}}
.kpi-grade{{font-family:var(--ff-h);font-weight:800;font-size:1.6rem;line-height:1}}
.kpi-pill{{font-family:var(--ff-m);font-size:12px;font-weight:600;padding:4px 12px;border-radius:20px;background:var(--surface);border:1px solid var(--border)}}
.ag-cols{{display:grid;grid-template-columns:1fr 1fr;border:1px solid var(--border);border-top:none;border-radius:0 0 10px 10px;overflow:hidden}}
.ag-cols.three{{grid-template-columns:1fr 1fr 1fr}}
.ag-col{{padding:1.25rem}} .ag-col+.ag-col{{border-left:1px solid var(--border)}}
.col-hdr{{font-family:var(--ff-m);font-size:10px;letter-spacing:.1em;text-transform:uppercase;font-weight:600;padding-bottom:.6rem;margin-bottom:.75rem;border-bottom:2px solid}}
.win-hdr{{color:var(--pos);border-color:var(--pos)}} .dep-hdr{{color:var(--neg);border-color:var(--neg)}} .ret-hdr{{color:var(--gold);border-color:var(--gold)}}
.mv-item{{border-radius:7px;padding:.65rem .8rem;margin-bottom:.5rem;font-size:12.5px;position:relative}}
.win-item{{background:#F0FDF4}} .dep-item{{background:#FEF2F2}} .ret-item{{background:#FFFBEB}}
.mv-main{{display:flex;align-items:center;justify-content:space-between;gap:.5rem;margin-bottom:.4rem}}
.mv-adv{{font-weight:700;flex:1}} .mv-sp{{font-family:var(--ff-m);font-size:12px;font-weight:700;flex-shrink:0}}
.mv-meta{{display:flex;align-items:center;gap:.35rem;flex-wrap:wrap}}
.mv-tag-badge{{font-size:9.5px;font-weight:700;letter-spacing:.04em;padding:2px 6px;border-radius:3px}}
.mv-scope{{font-size:10px;color:var(--muted);background:var(--surface);padding:1px 6px;border-radius:10px;border:1px solid var(--border)}}
.mv-pts{{font-family:var(--ff-m);font-size:11.5px;font-weight:700;padding:2px 8px;border-radius:10px}}
.mv-pts.pos{{background:#D1FAE5;color:var(--pos)}} .mv-pts.neg{{background:#FEE2E2;color:var(--neg)}}
.mv-pts.gold{{background:#FEF3C7;color:var(--gold)}} .mv-pts.neu{{background:var(--border);color:var(--muted)}}
.mv-note{{font-size:11px;color:var(--muted);font-style:italic;margin-top:.3rem}}
.mv-empty{{font-size:12px;color:var(--muted);font-style:italic;padding:.5rem .25rem}}
.mv-info{{display:flex;flex-wrap:wrap;gap:.4rem .85rem;margin-top:.45rem;padding-top:.45rem;border-top:1px solid rgba(0,0,0,.06)}}
.mv-info-item{{display:flex;align-items:baseline;gap:.3rem;font-size:11px}}
.mv-info-lbl{{color:var(--muted);font-weight:600;text-transform:uppercase;font-family:var(--ff-m);font-size:9.5px;letter-spacing:.06em;flex-shrink:0}}
.mv-info-val{{color:var(--text);font-weight:500}}
.mv-info-note{{color:var(--muted);font-style:italic}}
.pos{{color:var(--pos)}} .neg{{color:var(--neg)}} .gold{{color:var(--gold)}} .neu{{color:var(--muted)}} .neg-txt{{color:var(--neg)}}
@media(max-width:680px){{
  .cover-hero{{padding:2rem 1.25rem 1.75rem}} .method-wrap{{padding:1.5rem 1.25rem}}
  .method-grid{{grid-template-columns:1fr 1fr}}
  .method-cell:nth-child(2){{border-right:none}} .method-cell:nth-child(3){{border-right:1px solid var(--border)}}
  .method-cell:nth-child(1),.method-cell:nth-child(2){{border-bottom:1px solid var(--border)}}
  .ag-cols,.ag-cols.three{{grid-template-columns:1fr}}
  .ag-col+.ag-col{{border-left:none;border-top:1px solid var(--border)}}
}}
{EDITOR_CSS}
</style>
</head>
<body>

<nav class="top-nav">
  <div class="nav-inner">
    <div class="nav-brand">RECMA <span>·</span> {market.upper()}</div>
    <a href="#section-0" class="nav-tab active"><span class="nav-num">01</span>Overview</a>
    <a href="#section-1" class="nav-tab"><span class="nav-num">02</span>Rankings</a>
    {nav_ag_tabs}
  </div>
</nav>

<section id="section-0">
  <div class="cover-hero">
    <div class="cover-eyebrow" data-editable="text">RECMA · {market} · Competitiveness in Pitches · {year}</div>
    <div class="cover-title" data-editable="text">Compitches<br>{year}</div>
    <div class="cover-year" data-editable="text">Preliminary · {market}</div>
  </div>
  <div class="method-wrap">
    <div class="method-label">Scoring Methodology</div>
    <div class="method-grid">
      <div class="method-cell">
        <div class="mc-type" style="color:var(--accent)">Key Pitch</div>
        <div class="mc-cond">Ad spend <strong>&gt; $5m</strong><br>Incumbent &amp; contenders known</div>
        <div class="mc-pts"><div class="pt-row pos">+3 pts <span class="lbl">win</span></div><div class="pt-row neg">−3 pts <span class="lbl">departure</span></div></div>
      </div>
      <div class="method-cell">
        <div class="mc-type" style="color:var(--accent)">Other Pitch</div>
        <div class="mc-cond">Ad spend <strong>$1m – $5m</strong><br>Incumbent &amp; contenders known</div>
        <div class="mc-pts"><div class="pt-row pos">+1 pt <span class="lbl">win</span></div><div class="pt-row neg">−1 pt <span class="lbl">departure</span></div></div>
      </div>
      <div class="method-cell">
        <div class="mc-type" style="color:var(--gold)">Retention</div>
        <div class="mc-cond">Client kept after pitch with identified contenders</div>
        <div class="mc-pts"><div class="pt-row gold">+3 pts <span class="lbl">&gt; $5m</span></div><div class="pt-row gold">+1 pt <span class="lbl">$1m–$5m</span></div></div>
      </div>
      <div class="method-cell">
        <div class="mc-type" style="color:var(--accent2)">Hub Bonus</div>
        <div class="mc-cond">Local agency = regional hub for reference client</div>
        <div class="mc-pts"><div class="pt-row" style="color:var(--accent2)">+1 pt</div></div>
        <div style="margin-top:.75rem;padding-top:.75rem;border-top:1px solid var(--border);font-size:11px;color:var(--muted);line-height:1.5">Not counted: &lt;$1m · unknown contenders</div>
      </div>
    </div>
  </div>
</section>

<section id="section-1" class="page" style="border-top:2px solid var(--border)">
  <div class="sec-label">02 — Competitiveness in Pitches</div>
  <div class="sec-title">Rankings {year} <span class="sub">· {market} · By Points Balance</span></div>
  <div class="group-legend">
    <div class="gleg" style="background:#F3E8FF;border-color:#9333EA;color:#9333EA">IPG Mediabrands</div>
    <div class="gleg" style="background:#EDE9FE;border-color:#7C3AED;color:#7C3AED">Havas Media Network</div>
    <div class="gleg" style="background:#FEE2E2;border-color:#DC2626;color:#DC2626">WPP Media</div>
    <div class="gleg" style="background:#FEF9C3;border-color:#CA8A04;color:#CA8A04">Omnicom Media</div>
    <div class="gleg" style="background:#FFEDD5;border-color:#EA580C;color:#EA580C">Publicis Media</div>
    <div class="gleg" style="background:#DCFCE7;border-color:#16A34A;color:#16A34A">dentsu</div>
    <div class="gleg" style="background:#F1F5F9;border-color:#64748B;color:#64748B">Independent</div>
  </div>
  <div class="rank-wrap">
    <table class="rtable">
      <thead>
        <tr>
          <th>#</th><th>Agency</th>
          <th style="text-align:center">Grade</th>
          <th style="text-align:center">Pts</th>
          <th style="text-align:center">W – D</th>
          <th>Key Win</th><th>Key Departure</th>
        </tr>
      </thead>
      <tbody>{ranking_rows}</tbody>
    </table>
  </div>
</section>

{detail_sections}

<!-- ── EDITOR UI ── -->
<div id="edit-toolbar">
  <button id="export-btn" onclick="exportHTML()">💾 Export HTML</button>
  <button id="edit-toggle" onclick="toggleEditMode()">✏️ EDIT</button>
</div>

<div id="color-picker-popup">
  <h4>Background Color</h4>
  <div class="cp-swatches"></div>
  <div class="cp-custom">
    <label>Custom</label>
    <input type="color" value="#F3E8FF">
  </div>
  <button id="cp-close">Close</button>
</div>

<script>{EDITOR_JS}</script>
</body>
</html>"""

    return html.encode('utf-8')
