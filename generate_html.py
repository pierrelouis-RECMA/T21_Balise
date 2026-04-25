#!/usr/bin/env python3
"""
generate_html.py — NBB Report HTML Generator
═════════════════════════════════════════════
Génère un rapport HTML complet depuis un DataFrame Excel NBB.
Responsive : mobile, desktop, impression A4.
Fichier unique autonome, sans dépendances externes.

Usage :
    python generate_html.py <excel> <output.html>
    python generate_html.py data.xlsx NBB_Report.html

Appelé depuis app.py :
    from generate_html import build_report_html
    html_bytes = build_report_html(df)
"""

import sys, os, io
import pandas as pd

# ─────────────────────────────────────────────────────────────
# DATA (réutilise fill_template)
# ─────────────────────────────────────────────────────────────

def get_data(df):
    sys.path.insert(0, os.path.dirname(__file__))
    from fill_template import load_data_from_df
    data = load_data_from_df(df)
    # Extraire market et period depuis l'Excel
    df2 = df.copy()
    df2.columns = [str(c).strip() for c in df2.columns]
    market = str(df2['Country'].dropna().iloc[0]).title() if 'Country' in df2.columns and len(df2['Country'].dropna()) > 0 else 'Market'
    year   = str(int(df2['Years'].dropna().iloc[0]))      if 'Years'   in df2.columns and len(df2['Years'].dropna())   > 0 else '2025'
    data['market'] = market
    data['period'] = f'NBB {year}'
    return data


# ─────────────────────────────────────────────────────────────
# CONSTANTES GLOBALES
# ─────────────────────────────────────────────────────────────

# Couleurs officielles groupes
GROUP_COLORS = {
    'Publicis Media':      '#FFEADD',
    'Omnicom Media':       '#FFE699',
    'Dentsu':              '#E2F0D9',
    'Havas Media Network': '#DCB9FF',
    'WPP Media':           '#FFE4FF',
    'Independant':         '#F3F4F6',
}

# Couleur de bordure / barre (version foncée assortie)
GROUP_BORDER_COLORS = {
    'Publicis Media':      '#C4570A',
    'Omnicom Media':       '#B8960A',
    'Dentsu':              '#4A8A3C',
    'Havas Media Network': '#7B3FA0',
    'WPP Media':           '#A040A0',
    'Independant':         '#6B7280',
}

def group_bg(group_name):
    return GROUP_COLORS.get(group_name, '#F8F8F8')

def group_border(group_name):
    return GROUP_BORDER_COLORS.get(group_name, '#CCCCCC')

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────

def fmt(v):
    if v is None: return ''
    if isinstance(v, float) and v == 0: return '$0m'
    if isinstance(v, (int, float)):
        return f'+{v:.1f}m$' if v > 0 else f'{v:.1f}m$'
    return str(v)

def fmtv(v):
    """Valeur sans $"""
    if not v and v != 0: return ''
    if isinstance(v, (int, float)):
        return f'+{v:.1f}m' if v > 0 else (f'{v:.1f}m' if v != 0 else '')
    return str(v)

def nbb_class(v):
    if isinstance(v, (int, float)):
        return 'pos' if v > 0 else ('neg' if v < 0 else 'neu')
    s = str(v)
    return 'pos' if s.startswith('+') else ('neg' if s.startswith('-') else 'neu')

def trunc(s, n=30):
    s = str(s).strip()
    return s[:n-1]+'…' if len(s) > n else s


# ─────────────────────────────────────────────────────────────
# SECTION BUILDERS
# ─────────────────────────────────────────────────────────────

def build_nav(active):
    tabs = [
        ('01', 'Key Findings'),
        ('02', 'TOP moves'),
        ('03', 'NBB · Agencies'),
        ('04', 'NBB · Groups'),
        ('05', 'Retentions'),
        ('06', 'Details'),
    ]
    items = ''
    for i, (num, label) in enumerate(tabs):
        cls = ' active' if i == active else ''
        items += f'<a href="#section-{i}" class="nav-tab{cls}"><span class="nav-num">{num}</span>{label}</a>'
    return f'<nav class="top-nav"><div class="nav-inner">{items}</div></nav>'


def build_cover(data):
    """
    Slide 01 — Key Findings
    Reproduit exactement la slide PPT :
      - Titre "New Business Balance · Key Findings"
      - Bloc Perimeter & Methodology (texte statique)
      - Tableau TOP 4 agencies + Group ranking bar chart côte à côte
      - Bloc 5 Key Takeaways
    """
    agencies = data['agencies']
    top4     = agencies[:4]

    # ── Tableau TOP 4 agencies ───────────────────────────────
    ag_rows = ''
    for a in top4:
        ag_rows += f"""<tr>
            <td class="kf-rank">{a["rank"]}</td>
            <td class="kf-agency">{a["agency"].title()}</td>
            <td class="kf-nbb {nbb_class(a["nbb"])}">{fmt(a["nbb"])}</td>
            <td class="kf-num pos">{int(round(a["wins"]))}</td>
            <td class="kf-num neg">{int(round(a["deps"]))}</td>
          </tr>"""

    # ── Group ranking bar chart ──────────────────────────────
    groups_sorted = sorted(
        [gs for gs in data['group_stats'].values() if gs['agencies']],
        key=lambda g: -g['nbb']
    )
    max_val  = max(abs(gs['nbb']) for gs in groups_sorted) or 1
    # Axe X : va de -max_val à +max_val, zero au centre
    chart_bars = ''
    for gs in groups_sorted:
        nbb     = gs['nbb']
        is_pos  = nbb >= 0
        abs_pct = abs(nbb) / max_val * 45   # max 45% de chaque côté
        short_name = gs['name'].replace(' Media Network','').replace(' Media','')
        bar_style = (
            f"margin-left:50%;width:{abs_pct:.1f}%"
            if is_pos else
            f"margin-left:{50 - abs_pct:.1f}%;width:{abs_pct:.1f}%"
        )
        val_str = f"{int(round(abs(nbb)))}"
        chart_bars += f"""<div class="grk-row">
          <div class="grk-label">{short_name}</div>
          <div class="grk-track">
            <div class="grk-bar {'grk-pos' if is_pos else 'grk-neg'}" style="{bar_style}"></div>
            <span class="grk-val {'grk-val-pos' if is_pos else 'grk-val-neg'}"
                  style="left:{( (50 + abs_pct + 1) if is_pos else max(0, 50 - abs_pct - 8) ):.1f}%">
              {val_str}
            </span>
          </div>
        </div>"""

    # ── Key Takeaways dynamiques ─────────────────────────────
    takeaways_html = _takeaways(data)

    return f'''<section id="section-0" class="page kf-page">
  <div class="page-header">
    <div class="section-header-inner"><span class="section-num">01</span><h2>New Business Balance <span class="kf-dot">·</span> <span class="kf-sub-title">Key Findings</span></h2></div>
  </div>

  <!-- PERIMETER & METHODOLOGY -->
  <div class="kf-perimeter">
    <div class="kf-perim-header">
      <strong>Perimeter</strong> &amp; <strong>Methodology</strong>
    </div>
    <div class="kf-perim-body">
      <p>The perimeter studied by RECMA includes the 5 international media groups.</p>
      <p>Moves are considered based on their <strong>date of announcement</strong>. We registered all <strong>classical media assignments</strong> (incl. planning, buying, strategic planning for all or only selected media) as well as <strong>digital or other specialized services assignments</strong>.</p>
      <p>Spends are in <strong>Integrated Spendings</strong> incl. non-traditional activity (digital, data, content)</p>
      <p>The newbiz balance ranking is focusing on <strong>Net New Biz</strong> and therefore retentions, contract renewal and transfers are not included in the calculation. However, these key informations can be found in the detailed tables for each network. Moreover, we do value retentions/contract renewals in two different criteria of the qualitative evaluation : the <em>Competitiveness in pitches</em> criteria (Table 18) as well as <em>Client relationship stability</em> (Table 25).</p>
    </div>
  </div>

  <!-- TOP 3 + GROUP RANKING -->
  <div class="kf-two-col">

    <!-- LEFT : TOP 4 agencies table -->
    <div class="kf-block">
      <div class="kf-block-header">
        <span class="kf-bh-bold">TOP 3</span> agencies <span class="kf-bh-bold">2025</span>
        <span class="kf-bh-sub">· Perimeter &amp; Market Growth</span>
      </div>
      <div class="kf-table-wrap">
        <table class="kf-table">
          <thead>
            <tr>
              <th>Rk</th>
              <th class="kf-th-market">Market<br><small>Top agencies</small></th>
              <th>New Biz<br>Balance*</th>
              <th class="kf-th-green">Wins<br>$m</th>
              <th class="kf-th-red">Dep.<br>$m</th>
            </tr>
          </thead>
          <tbody>{ag_rows}</tbody>
        </table>
        <div class="kf-footnote">* Not incl. Retentions, Client renewal and Transfers.</div>
      </div>
    </div>

    <!-- RIGHT : Group ranking chart -->
    <div class="kf-block">
      <div class="kf-block-header">
        <span class="kf-bh-bold">Group</span> ranking <span class="kf-bh-bold">2025</span>
        <span class="kf-bh-sub">· NBB Balance</span>
      </div>
      <div class="kf-chart">
        <div class="grk-zero-line"></div>
        {chart_bars}
      </div>
    </div>

  </div>

  <!-- 5 KEY TAKEAWAYS -->
  <div class="kf-takeaways-block">
    <div class="kf-takeaways-header">5 Key Takeaways</div>
    <div class="kf-takeaways-body">
      {takeaways_html}
    </div>
  </div>

</section>'''

def _takeaways(data):
    """Génère les 5 Key Takeaways dynamiques basés sur les groupes."""
    groups_sorted = sorted(
        [gs for gs in data['group_stats'].values() if gs['agencies']],
        key=lambda g: -g['nbb']
    )[:5]
    lines = []
    for gs in groups_sorted:
        top_ag   = gs['agencies'][0] if gs['agencies'] else None
        nbb_str  = fmt(gs['nbb'])
        cls      = nbb_class(gs['nbb'])
        top_name = top_ag['agency'].title() if top_ag else ''
        top_nbb  = fmt(top_ag['nbb']) if top_ag else ''
        wins_str = f"{gs['wc']} win{'s' if gs['wc']>1 else ''}"
        deps_str = f"{gs['dc']} departure{'s' if gs['dc']>1 else ''}"
        detail   = f" — {top_name} leads with <strong>{top_nbb}</strong>" if top_ag else ''
        lines.append(
            f'''<div class="takeaway">
              <span class="take-bullet">•</span>
              <span><strong>{gs["name"]}</strong>
                : NBB <strong class="{cls}">{nbb_str}</strong>
                ({wins_str} / {deps_str}){detail}
              </span>
            </div>'''
        )
    return '\n'.join(lines)


def build_top_moves(data):
    wins = data['top_wins'][:5]
    deps = data['top_deps'][:5]
    rets = data['top_rets'][:4]

    def win_items():
        out = ''
        for r in wins:
            adv = trunc(str(r.get('Advertiser','')), 26)
            ag  = trunc(str(r.get('Agency','')), 20)
            val = fmtv(float(r.get('Integrated Spends', 0)))
            out += f'''<div class="move-item win-item">
              <div class="move-main">
                <span class="move-adv">{adv}</span>
                <span class="move-val pos">{val}</span>
              </div>
              <div class="move-ag">→ {ag}</div>
            </div>'''
        return out

    def dep_items():
        out = ''
        for r in deps:
            adv = trunc(str(r.get('Advertiser','')), 26)
            ag  = trunc(str(r.get('Agency','')), 20)
            val = fmtv(float(r.get('Integrated Spends', 0)))
            out += f'''<div class="move-item dep-item">
              <div class="move-main">
                <span class="move-adv">{adv}</span>
                <span class="move-val neg">{val}</span>
              </div>
              <div class="move-ag">← {ag}</div>
            </div>'''
        return out

    def ret_items():
        out = ''
        for r in rets:
            adv = trunc(str(r.get('Advertiser','')), 26)
            ag  = trunc(str(r.get('Agency','')), 20)
            val = fmtv(float(r.get('Integrated Spends', 0)))
            out += f'''<div class="move-item ret-item">
              <div class="move-main">
                <span class="move-adv">{adv}</span>
                <span class="move-val" style="color:var(--gold)">{val}</span>
              </div>
              <div class="move-ag">↺ {ag}</div>
            </div>'''
        return out

    return f'''<section id="section-1" class="page">
  <div class="page-header">
    <div class="section-header-inner"><span class="section-num">02</span><h2>TOP moves / retentions <span class="year">· 2025</span></h2></div>
    <p class="page-sub">Top wins, departures &amp; retentions by Integrated Spends</p>
  </div>

  <div class="moves-grid">
    <div class="moves-col">
      <div class="moves-col-header win-hdr">
        <span class="col-icon">↑</span> WINS MAJEURS
      </div>
      {win_items()}
    </div>
    <div class="moves-col">
      <div class="moves-col-header dep-hdr">
        <span class="col-icon">↓</span> DÉPARTS MAJEURS
      </div>
      {dep_items()}
    </div>
    <div class="moves-col">
      <div class="moves-col-header ret-hdr">
        <span class="col-icon">↺</span> RÉTENTIONS NOTABLES
      </div>
      {ret_items()}
    </div>
  </div>

  <div class="page-note">Retentions / renewals / transfers not included in NBB calculation</div>
</section>'''


def build_agencies_overview(data, threshold=5.0):
    """
    threshold : seuil en $m pour filtrer les wins/deps dans le tableau
                (configurable depuis l'interface Render)
    """
    agencies = data['agencies']

    # ── Bar chart ────────────────────────────────────────────
    max_abs = max(abs(a['nbb']) for a in agencies) or 1
    chart_bars = ''
    for a in agencies:
        pct  = abs(a['nbb']) / max_abs * 100
        side = 'pos' if a['nbb'] >= 0 else 'neg'
        chart_bars += f'''<div class="bar-row">
          <div class="bar-label">{trunc(a["agency"], 18)}</div>
          <div class="bar-track">
            <div class="bar-fill {side}" style="width:{pct:.1f}%;{"margin-left:auto" if side=="neg" else ""}"></div>
          </div>
          <div class="bar-val {side}">{fmt(a["nbb"])}</div>
        </div>'''

    # ── Tableau avec seuil ────────────────────────────────────
    thr = float(threshold)
    table_rows = ''
    for a in agencies:
        bg     = group_bg(a['group'])
        border = group_border(a['group'])

        wins_str = '  ·  '.join([
            f"{trunc(r['Advertiser'],18)} {fmtv(float(r['Integrated Spends']))}"
            for r in a['wins_rows']
            if float(r.get('Integrated Spends', 0)) >= thr
        ])
        deps_str = '  ·  '.join([
            f"{trunc(r['Advertiser'],18)} {fmtv(float(r['Integrated Spends']))}"
            for r in a['dep_rows']
            if float(r.get('Integrated Spends', 0)) <= -thr
        ])
        top_ret  = trunc(a['ret_rows'][0]['Advertiser'], 20) if a['ret_rows'] else '—'

        table_rows += f'''<tr style="background:{bg};border-left:3px solid {border}">
          <td class="td-rank">#{a["rank"]}</td>
          <td class="td-agency" style="font-weight:700">{trunc(a["agency"], 18)}</td>
          <td class="td-nbb {nbb_class(a["nbb"])}">{fmt(a["nbb"])}</td>
          <td class="td-wins pos">{fmtv(a["wins"]) or "0"}</td>
          <td class="td-deps neg">{fmtv(a["deps"]) or "0"}</td>
          <td class="td-topwins">{wins_str or "—"}</td>
          <td class="td-topdeps">{deps_str or "—"}</td>
          <td class="td-topret">{top_ret}</td>
        </tr>'''

    thr_display = f"≥ ${thr:.0f}m / ≤ -${thr:.0f}m"

    return f'''<section id="section-2" class="page">
  <div class="page-header">
    <div class="section-header-inner"><span class="section-num">03</span><h2>NBB 2025 agencies overview <span class="year">· Sep.24 – Sep.25</span></h2></div>
    <p class="page-sub">Retentions, contract renewals &amp; transfers not included · By decreasing NBB balance</p>
  </div>

  <div class="chart-section">
    <div class="chart-legend">
      <span class="leg neg">◀ Departures ($m)</span>
      <span class="leg pos">NBB balance ($m) ▶</span>
    </div>
    <div class="bar-chart">{chart_bars}</div>
  </div>

  <div class="table-section-header">
    <span>Agency details</span>
    <span class="threshold-badge">Threshold: {thr_display}</span>
  </div>
  <div class="table-scroll">
    <table class="data-table agencies-table">
      <thead>
        <tr>
          <th>#</th><th>Agency</th><th>NBB</th>
          <th>Wins $m</th><th>Dep. $m</th>
          <th>Top wins ({thr_display.split("/")[0].strip()})</th>
          <th>Top dep. ({thr_display.split("/")[1].strip()})</th>
          <th>Top Ret.</th>
        </tr>
      </thead>
      <tbody>{table_rows}</tbody>
    </table>
  </div>
  <div class="page-note">Retentions / renewals / transfers not included in NBB · By decreasing NBB balance</div>
</section>'''

def build_groups_overview(data):
    sorted_groups = sorted(
        [gs for gs in data['group_stats'].values() if gs['agencies']],
        key=lambda g: -g['nbb']
    )
    rows = ''
    for gs in sorted_groups:
        bg     = group_bg(gs['name'])
        border = group_border(gs['name'])
        rows += f'''<tr class="group-row" style="background:{bg}">
          <td class="td-group" style="border-left:4px solid {border}"><strong>#{gs["rank"]} {gs["name"]}</strong></td>
          <td class="td-num">{gs["wc"]}</td>
          <td class="td-num">{gs["dc"]}</td>
          <td class="td-nbb {nbb_class(gs["nbb"])}"><strong>{fmt(gs["nbb"])}</strong></td>
          <td class="td-wins pos"><strong>{fmtv(gs["wins"])}</strong></td>
          <td class="td-deps neg"><strong>{fmtv(gs["deps"])}</strong></td>
        </tr>'''
        for a in gs['agencies']:
            rows += f'''<tr class="agency-sub-row" style="background:{bg}44">
              <td class="td-agency-sub" style="border-left:3px solid {border}">
                &nbsp;&nbsp;&nbsp;{trunc(a["agency"].title(), 24)}
              </td>
              <td class="td-num">{a["wc"]}</td>
              <td class="td-num">{a["dc"]}</td>
              <td class="td-nbb {nbb_class(a["nbb"])}">{fmt(a["nbb"])}</td>
              <td class="td-wins pos">{fmtv(a["wins"]) or "0"}</td>
              <td class="td-deps neg">{fmtv(a["deps"]) or "0"}</td>
            </tr>'''

    total_nbb  = sum(gs['nbb']  for gs in data['group_stats'].values())
    total_wins = sum(gs['wins'] for gs in data['group_stats'].values())
    total_deps = sum(gs['deps'] for gs in data['group_stats'].values())
    total_wc   = sum(gs['wc']   for gs in data['group_stats'].values())
    total_dc   = sum(gs['dc']   for gs in data['group_stats'].values())
    rows += f'''<tr class="total-row">
      <td><strong>TOTAL 5 groups</strong></td>
      <td class="td-num"><strong>{total_wc}</strong></td>
      <td class="td-num"><strong>{total_dc}</strong></td>
      <td class="td-nbb {nbb_class(total_nbb)}"><strong>{fmt(total_nbb)}</strong></td>
      <td class="td-wins pos"><strong>{fmtv(total_wins)}</strong></td>
      <td class="td-deps neg"><strong>{fmtv(total_deps)}</strong></td>
    </tr>'''

    return f'''<section id="section-3" class="page">
  <div class="page-header">
    <div class="section-header-inner"><span class="section-num">04</span><h2>NBB 2025 groups overview <span class="year">· Sep.24 – Sep.25</span></h2></div>
    <p class="page-sub">Retentions, contract renewals &amp; transfers not included · By decreasing NBB balance</p>
  </div>
  <div class="table-scroll">
    <table class="data-table groups-table">
      <thead>
        <tr>
          <th>Media Group / Agency</th>
          <th>Wins (nb)</th><th>Dep. (nb)</th>
          <th>Net Balance ($m)</th>
          <th>Wins ($m)</th><th>Dep. ($m)</th>
        </tr>
      </thead>
      <tbody>{rows}</tbody>
    </table>
  </div>
  <div class="page-note">Retentions / renewals / transfers not included in NBB · By decreasing NBB balance</div>
</section>'''

def build_retentions(data):
    ret_data = data['ret_by_agency']
    ag_group = {a['agency']: a['group'] for a in data['agencies']}
    max_val  = max((r['balance'] for r in ret_data), default=1)

    bars = ''
    for r in ret_data[:8]:
        pct    = r['balance'] / max_val * 100 if max_val else 0
        grp    = ag_group.get(r['agency'], 'Independant')
        color  = group_border(grp)   # couleur SOLIDE du groupe
        bg     = group_bg(grp)       # fond clair du groupe
        # Dégradé foncé→clair pour que même les petites barres aient la bonne couleur
        gradient = f"linear-gradient(90deg, {color} 0%, {color}99 100%)"
        ag_name  = trunc(r["agency"].title(), 20)

        bars += f'''<div class="ret-bar-row">
          <div class="ret-bar-label" style="background:{bg};border-left:4px solid {color}">
            {ag_name}
          </div>
          <div class="ret-bar-track">
            <div class="ret-bar-fill" style="width:{pct:.1f}%;background:{gradient}"></div>
          </div>
          <div class="ret-val" style="color:{color}">{fmt(r["balance"])}</div>
        </div>'''

    # Table
    table_rows = ''
    for r in ret_data[:8]:
        grp    = ag_group.get(r['agency'], 'Independant')
        bg     = group_bg(grp)
        border = group_border(grp)
        table_rows += f'''<tr style="background:{bg}">
          <td class="td-agency" style="border-left:4px solid {border};font-weight:700">{trunc(r["agency"].title(), 22)}</td>
          <td class="td-nbb pos" style="color:{border}">{fmt(r["balance"])}</td>
          <td class="td-topclient">{r["top_client"]}</td>
        </tr>'''

    return f'''<section id="section-4" class="page">
  <div class="page-header">
    <div class="section-header-inner"><span class="section-num">05</span><h2>NBB 2025 retentions ranking <span class="year">· Sep.24 – Sep.25</span></h2></div>
    <p class="page-sub">Retentions &amp; contract renewals not included in the NBB calculation · By decreasing retention balance</p>
  </div>
  <div class="two-col">
    <div>
      <div class="ret-chart-title">Retention balance by agency ($m)</div>
      <div class="ret-chart">{bars}</div>
    </div>
    <div class="table-scroll">
      <table class="data-table ret-table">
        <thead>
          <tr><th>Agency</th><th>Balance ($m)</th><th>Principal account retained</th></tr>
        </thead>
        <tbody>{table_rows}</tbody>
      </table>
    </div>
  </div>
</section>'''

def build_agency_details(data):
    agencies = data['agencies']
    CHUNK = 4

    def agency_card(a):
        def col_items(rows_key, cls):
            items = a[rows_key]
            if not items:
                return '<div class="det-empty">—</div>'
            out = ''
            for r in items:
                adv = trunc(str(r.get('Advertiser','')), 24)
                val = fmtv(float(r.get('Integrated Spends', 0)))
                val_cls = 'pos' if (val and val.startswith('+')) else ('neg' if val else '')
                out += f'<div class="det-item"><span class="det-adv">{adv}</span>'
                if val:
                    out += f'<span class="det-val {val_cls}">{val}</span>'
                out += '</div>'
            return out

        nbb_v  = a['nbb']
        bg     = group_bg(a['group'])
        border = group_border(a['group'])
        return f'''<div class="agency-card" style="border-top:3px solid {border}">
          <div class="card-header" style="background:{bg};border-bottom:2px solid {border}">
            <div class="card-agency-name" style="color:#1E293B">{a["agency"]}</div>
            <div class="card-group" style="color:#64748B">({a["group"]})</div>
            <div class="card-badge {nbb_class(nbb_v)}">{fmt(nbb_v)}</div>
          </div>
          <div class="card-body">
            <div class="det-col">
              <div class="det-col-hdr win-lbl">WIN</div>
              <div class="det-col-ul win-ul"></div>
              {col_items("wins_rows", "win")}
            </div>
            <div class="det-col">
              <div class="det-col-hdr dep-lbl">DEPARTURE</div>
              <div class="det-col-ul dep-ul"></div>
              {col_items("dep_rows", "dep")}
            </div>
            <div class="det-col">
              <div class="det-col-hdr ret-lbl">RETENTION</div>
              <div class="det-col-ul ret-ul"></div>
              {col_items("ret_rows", "ret")}
            </div>
          </div>
        </div>'''

    pages_html = ''
    for chunk_i, start in enumerate(range(0, len(agencies), CHUNK)):
        chunk = agencies[start:start+CHUNK]
        cards = ''.join(agency_card(a) for a in chunk)
        pages_html += f'<div class="detail-page print-page">{cards}</div>'

    return f'''<section id="section-5" class="page details-section">
  <div class="page-header">
    <div class="section-header-inner"><span class="section-num">06</span><h2>Details by agency <span class="year">· 2025</span></h2></div>
    <p class="page-sub">Retentions &amp; renewals not included in NBB calculation</p>
  </div>
  {pages_html}
</section>'''


# ─────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────

CSS = """
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

:root {
  --bg:       #F7F8FC;
  --surface:  #FFFFFF;
  --border:   #E2E8F0;
  --nav-bg:   #0F172A;
  --nav-txt:  #64748B;
  --nav-act:  #FFFFFF;
  --accent:   #2D5C54;
  --accent2:  #38BDF8;
  --pos:      #059669;
  --neg:      #E11D48;
  --ret:      #B45309;
  --text:     #1E293B;
  --muted:    #64748B;
  --heading:  'Syne', sans-serif;
  --body:     'DM Sans', sans-serif;
  --mono:     'DM Mono', monospace;

  --page-max: 960px;
  --gap:      1.5rem;
}

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

html { scroll-behavior: smooth; }

body {
  background: var(--bg);
  color: var(--text);
  font-family: var(--body);
  font-size: 14px;
  line-height: 1.5;
}

/* ── NAV ── */
.top-nav {
  position: sticky;
  top: 0;
  z-index: 100;
  background: var(--nav-bg);
  box-shadow: 0 2px 12px rgba(0,0,0,.3);
}
.nav-inner {
  max-width: var(--page-max);
  margin: 0 auto;
  display: flex;
  overflow-x: auto;
  scrollbar-width: none;
}
.nav-inner::-webkit-scrollbar { display: none; }
.nav-tab {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 12px 16px;
  color: var(--nav-txt);
  text-decoration: none;
  font-size: 12px;
  font-weight: 500;
  letter-spacing: .04em;
  white-space: nowrap;
  transition: color .2s, background .2s;
  border-bottom: 2px solid transparent;
}
.nav-tab:hover { color: #fff; }
.nav-tab.active { color: var(--nav-act); border-bottom-color: var(--accent2); }
.nav-num {
  font-family: var(--heading);
  font-size: 13px;
  font-weight: 800;
  opacity: .6;
  letter-spacing: -.02em;
  min-width: 20px;
}

/* ── PAGE ── */
.page {
  max-width: var(--page-max);
  margin: 0 auto;
  padding: 2.5rem 1.5rem;
}
.page + .page { border-top: 2px solid var(--border); }

.page-header { margin-bottom: 1.75rem; }
.page-header h2 {
  font-family: var(--heading);
  font-size: clamp(1.3rem, 3vw, 1.8rem);
  font-weight: 700;
  color: var(--accent);
  line-height: 1.2;
}
.page-header h2 .year { color: var(--accent2); font-weight: 400; }
.page-sub { color: var(--muted); font-size: 12px; margin-top: 4px; }
.page-note { color: var(--muted); font-size: 11px; margin-top: 1rem; font-style: italic; }

/* ── COVER ── */
/* ── COVER PPT-STYLE ── */
.cover-page { padding: 0 !important; }
.cover-layout {
  display: grid;
  grid-template-columns: 240px 1fr;
  min-height: 420px;
}
.cover-left {
  background: var(--accent);
  padding: 2.5rem 1.75rem;
  display: flex;
  flex-direction: column;
  gap: 1rem;
}
.cover-label-top {
  font-family: var(--mono);
  font-size: 10px;
  letter-spacing: .18em;
  text-transform: uppercase;
  color: rgba(255,255,255,.5);
}
.cover-title {
  font-family: var(--heading);
  font-size: clamp(1.8rem, 4vw, 2.8rem);
  font-weight: 800;
  color: #fff;
  line-height: 1.05;
}
.cover-market {
  font-family: var(--mono);
  font-size: 13px;
  font-weight: 600;
  color: var(--accent2);
  margin-top: auto;
}
.cover-period {
  font-size: 12px;
  color: rgba(255,255,255,.5);
  font-family: var(--mono);
}
.cover-right {
  background: var(--bg);
  padding: 2.5rem 2rem;
  display: flex;
  align-items: center;
}
/* TOC list */
.toc-list { display: flex; flex-direction: column; gap: .85rem; width: 100%; }
.toc-item {
  display: grid;
  grid-template-columns: 60px 16px 1fr;
  align-items: center;
  gap: 0;
  text-decoration: none;
  color: var(--text);
  padding-bottom: .85rem;
  border-bottom: 1px solid var(--border);
  transition: color .15s;
}
.toc-item:last-child { border-bottom: none; padding-bottom: 0; }
.toc-item:hover { color: var(--accent); }
.toc-num {
  font-family: var(--heading);
  font-size: 1.6rem;
  font-weight: 800;
  color: var(--accent);
  line-height: 1;
  letter-spacing: -.03em;
}
.toc-sep {
  color: var(--accent);
  font-size: 1.1rem;
  font-weight: 300;
  margin: 0 8px;
}
.toc-label {
  font-size: 13px;
  font-weight: 500;
}
/* Cover bottom cards */
.cover-bottom {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: var(--gap);
  padding: 1.5rem;
  background: var(--surface);
  border-top: 1px solid var(--border);
}
.cover-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 10px;
  padding: 1.25rem;
}
.card-title {
  font-family: var(--heading);
  font-size: 13px;
  font-weight: 700;
  letter-spacing: .06em;
  text-transform: uppercase;
  color: var(--accent);
  margin-bottom: 1rem;
  padding-bottom: .5rem;
  border-bottom: 2px solid var(--accent);
}
.cover-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.cover-table thead th {
  text-align: left;
  font-size: 10px;
  text-transform: uppercase;
  letter-spacing: .06em;
  color: var(--muted);
  padding: 4px 6px;
  border-bottom: 1px solid var(--border);
}
.cover-table td { padding: 6px; border-bottom: 1px solid var(--border); }
.cover-table .rank { font-family: var(--mono); color: var(--muted); font-size: 12px; }
.cover-table .agency-name { font-weight: 600; }
.cover-table .num { text-align: right; font-family: var(--mono); font-size: 12px; }
.footnote { font-size: 10px; color: var(--muted); margin-top: .5rem; font-style: italic; }

.takeaways { display: flex; flex-direction: column; gap: .75rem; }
.takeaway {
  display: flex;
  gap: .5rem;
  font-size: 12.5px;
  line-height: 1.5;
  padding: .5rem;
  background: var(--bg);
  border-radius: 5px;
  border-left: 3px solid var(--accent);
}
.take-bullet { color: var(--accent); font-weight: 700; flex-shrink: 0; }

/* ── KEY FINDINGS (section-0) ── */
.kf-page { background: var(--surface); }
.kf-page .page-header h2 {
  font-size: clamp(1.4rem, 3vw, 2rem);
  color: var(--accent);
}
.kf-dot { color: var(--muted); font-weight: 300; margin: 0 6px; }
.kf-sub-title { color: var(--accent2); font-weight: 400; }

/* Perimeter block */
.kf-perimeter {
  border: 1px solid var(--border);
  border-radius: 8px;
  overflow: hidden;
  margin-bottom: 1.5rem;
}
.kf-perim-header {
  background: var(--accent);
  color: #fff;
  padding: 10px 16px;
  font-size: 14px;
  text-align: center;
  letter-spacing: .03em;
}
.kf-perim-body {
  padding: 1rem 1.25rem;
  background: #F8FAF9;
  display: flex;
  flex-direction: column;
  gap: .6rem;
  font-size: 12.5px;
  line-height: 1.6;
  color: var(--text);
}

/* Two column layout */
.kf-two-col {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 1.25rem;
  margin-bottom: 1.5rem;
}
.kf-block {
  border: 1px solid var(--border);
  border-radius: 8px;
  overflow: hidden;
}
.kf-block-header {
  background: var(--accent);
  color: #fff;
  padding: 9px 14px;
  font-size: 13px;
  font-weight: 400;
}
.kf-bh-bold { font-weight: 700; }
.kf-bh-sub  { font-size: 11px; opacity: .7; margin-left: 4px; }

/* TOP agencies table */
.kf-table-wrap { padding: 0; }
.kf-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12px;
  background: var(--surface);
}
.kf-table thead th {
  background: #2D5C5422;
  color: var(--accent);
  padding: 7px 8px;
  font-size: 10px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: .04em;
  border-bottom: 2px solid var(--accent);
  text-align: center;
  vertical-align: middle;
}
.kf-th-market { text-align: left; }
.kf-th-green  { color: var(--pos) !important; }
.kf-th-red    { color: var(--neg) !important; }
.kf-table tbody tr { border-bottom: 1px solid var(--border); }
.kf-table tbody tr:last-child { border-bottom: none; }
.kf-table td { padding: 8px; vertical-align: middle; }
.kf-rank   { text-align: center; font-family: var(--mono); color: var(--muted); font-size: 11px; }
.kf-agency { font-weight: 600; font-size: 13px; }
.kf-nbb    { text-align: center; font-family: var(--mono); font-weight: 700; font-size: 13px; }
.kf-num    { text-align: center; font-family: var(--mono); font-weight: 600; font-size: 12px; }
.kf-footnote {
  font-size: 10px;
  color: var(--muted);
  font-style: italic;
  padding: 6px 10px;
  border-top: 1px solid var(--border);
}

/* Group ranking chart */
.kf-chart {
  padding: .75rem 1rem;
  background: var(--surface);
  position: relative;
  min-height: 180px;
}
.grk-zero-line {
  position: absolute;
  left: calc(140px + 50%);
  top: 0; bottom: 0;
  width: 1px;
  background: var(--border);
  transform: translateX(-50%);
}
.grk-row {
  display: grid;
  grid-template-columns: 80px 1fr;
  align-items: center;
  gap: .5rem;
  margin-bottom: .55rem;
}
.grk-label {
  font-size: 11px;
  font-weight: 500;
  text-align: right;
  color: var(--text);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.grk-track {
  height: 20px;
  position: relative;
  background: #F1F5F9;
  border-radius: 3px;
  overflow: visible;
}
.grk-bar {
  position: absolute;
  top: 0; bottom: 0;
  border-radius: 2px;
  min-width: 3px;
}
.grk-pos { background: linear-gradient(90deg, #BBF7D0, var(--pos)); }
.grk-neg { background: linear-gradient(90deg, var(--neg), #FECDD3); }
.grk-val {
  position: absolute;
  top: 50%;
  transform: translateY(-50%);
  font-family: var(--mono);
  font-size: 10px;
  font-weight: 700;
  white-space: nowrap;
}
.grk-val-pos { color: var(--pos); }
.grk-val-neg { color: var(--neg); }

/* Key Takeaways block */
.kf-takeaways-block {
  border: 1px solid var(--border);
  border-radius: 8px;
  overflow: hidden;
}
.kf-takeaways-header {
  background: var(--accent);
  color: #fff;
  padding: 10px 16px;
  font-size: 14px;
  font-weight: 700;
  text-align: center;
  letter-spacing: .02em;
}
.kf-takeaways-body {
  background: #F8FAF9;
  padding: 1rem 1.25rem;
  display: flex;
  flex-direction: column;
  gap: .6rem;
  min-height: 120px;
}
.kf-takeaways-body .takeaway {
  background: transparent;
  border: none;
  border-left: 3px solid var(--accent);
  border-radius: 0;
  padding: .35rem .75rem;
  font-size: 12.5px;
}

/* Responsive */
@media (max-width: 640px) {
  .kf-two-col { grid-template-columns: 1fr; }
  .grk-zero-line { display: none; }
  .grk-row { grid-template-columns: 70px 1fr; }
}

@media print {
  .kf-two-col { grid-template-columns: 1fr 1fr; }
  .kf-perimeter, .kf-block, .kf-takeaways-block { page-break-inside: avoid; }
}

/* ── TOP MOVES ── */
.moves-grid {
  display: grid;
  grid-template-columns: 1fr 1fr 1fr;
  gap: var(--gap);
}
.moves-col-header {
  font-size: 11px;
  font-weight: 700;
  letter-spacing: .1em;
  text-transform: uppercase;
  padding: .6rem .75rem;
  border-radius: 5px 5px 0 0;
  margin-bottom: .5rem;
}
.win-hdr  { background: #D1FAE5; color: var(--pos); }
.dep-hdr  { background: #FFE4E6; color: var(--neg); }
.ret-hdr  { background: #FEF3C7; color: var(--ret); }
.col-icon { margin-right: 4px; }

.move-item {
  padding: .75rem;
  border-radius: 6px;
  margin-bottom: .5rem;
  border: 1px solid var(--border);
  background: var(--surface);
}
.win-item { border-left: 3px solid var(--pos); }
.dep-item { border-left: 3px solid var(--neg); }
.ret-item { border-left: 3px solid var(--ret); }
.move-main { display: flex; justify-content: space-between; align-items: baseline; gap: .5rem; }
.move-adv { font-weight: 600; font-size: 13px; }
.move-val { font-family: var(--mono); font-size: 13px; font-weight: 700; white-space: nowrap; }
.move-ag { font-size: 11px; color: var(--muted); margin-top: 2px; }

/* ── BAR CHART ── */
.chart-section { margin-bottom: 1.5rem; }
.chart-legend {
  display: flex;
  gap: 1.5rem;
  margin-bottom: .75rem;
  font-size: 11px;
}
.leg { font-family: var(--mono); }
.leg.pos { color: var(--pos); }
.leg.neg { color: var(--neg); }

.bar-chart { display: flex; flex-direction: column; gap: .35rem; }
.bar-row {
  display: grid;
  grid-template-columns: 140px 1fr 80px;
  align-items: center;
  gap: .5rem;
  font-size: 12px;
}
.bar-label { text-align: right; font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.bar-track {
  height: 18px;
  background: var(--border);
  border-radius: 3px;
  overflow: hidden;
  display: flex;
}
.bar-fill {
  height: 100%;
  border-radius: 3px;
  transition: width .4s ease;
  min-width: 2px;
}
.bar-fill.pos { background: var(--pos); }
.bar-fill.neg { background: var(--neg); }
.bar-val { font-family: var(--mono); font-size: 11px; font-weight: 600; }
.bar-val.pos { color: var(--pos); }
.bar-val.neg { color: var(--neg); }

/* ── TABLES ── */
.table-scroll { overflow-x: auto; border-radius: 8px; border: 1px solid var(--border); }
.data-table { width: 100%; border-collapse: collapse; font-size: 12.5px; background: var(--surface); }
.data-table thead th {
  background: var(--accent);
  color: #fff;
  padding: 8px 10px;
  text-align: left;
  font-size: 11px;
  font-weight: 600;
  letter-spacing: .04em;
  white-space: nowrap;
}
.data-table tbody tr:nth-child(even) { background: #F8FAFC; }
.data-table tbody tr:hover { background: #EEF2FF; }
.data-table td { padding: 7px 10px; border-bottom: 1px solid var(--border); vertical-align: middle; }
.td-rank { font-family: var(--mono); color: var(--muted); font-size: 11px; }
.td-agency { font-weight: 600; }
.td-agency-sub { font-size: 12px; }
.td-nbb { font-family: var(--mono); font-weight: 700; white-space: nowrap; }
.td-wins { font-family: var(--mono); font-weight: 600; white-space: nowrap; }
.td-deps { font-family: var(--mono); font-weight: 600; white-space: nowrap; }
.td-num { text-align: center; font-family: var(--mono); }
.td-topwins, .td-topdeps { font-size: 11px; color: var(--muted); max-width: 200px; }
.td-topret { font-size: 11px; color: var(--ret); }
.td-topclient { font-size: 12px; }

.group-row td { font-size: 13px; }
.total-row td {
  background: #E8F5E9 !important;
  border-top: 2px solid var(--pos);
  font-size: 13px;
}

/* ── RETENTIONS ── */
.ret-chart {
  margin-bottom: 1.5rem;
  display: flex;
  flex-direction: column;
  gap: .4rem;
}
.ret-bar-row {
  display: grid;
  grid-template-columns: 140px 1fr;
  gap: .5rem;
  align-items: center;
}
.ret-bar-label { font-size: 12px; font-weight: 500; text-align: right; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.ret-bar-track {
  height: 20px;
  background: var(--border);
  border-radius: 4px;
  overflow: visible;
  display: flex;
  align-items: center;
  position: relative;
}
.ret-bar-fill {
  height: 100%;
  background: linear-gradient(90deg, #D1FAE5, var(--pos));
  border-radius: 4px;
  min-width: 4px;
}
.ret-bar-val {
  font-family: var(--mono);
  font-size: 11px;
  font-weight: 700;
  color: var(--pos);
  margin-left: .5rem;
  white-space: nowrap;
}

/* ── AGENCY DETAIL CARDS ── */
.detail-page {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 1rem;
  padding-bottom: 1.5rem;
  margin-bottom: 1.5rem;
  border-bottom: 1px solid var(--border);
}
.detail-page:last-child { border-bottom: none; margin-bottom: 0; }

.agency-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 10px;
  overflow: hidden;
  box-shadow: 0 1px 6px rgba(0,0,0,.05);
}
.card-header {
  padding: .65rem 1rem;
  display: flex;
  align-items: center;
  gap: .5rem;
  flex-wrap: wrap;
  border-radius: 0;
}
.card-agency-name {
  font-family: var(--heading);
  font-size: 14px;
  font-weight: 800;
  letter-spacing: .04em;
  flex: 1;
  color: #1E293B;
}
.card-group { font-size: 10px; color: #64748B; }
.card-badge {
  font-family: var(--mono);
  font-size: 12px;
  font-weight: 700;
  padding: 2px 8px;
  border-radius: 4px;
}
.card-badge.pos { background: #059669; color: #fff; }
.card-badge.neg { background: #E11D48; color: #fff; }
.card-badge.neu { background: #64748B; color: #fff; }

.card-body {
  display: grid;
  grid-template-columns: 1fr 1fr 1fr;
  gap: 0;
}
.det-col {
  padding: .6rem .5rem;
  border-right: 1px solid var(--border);
}
.det-col:last-child { border-right: none; }
.det-col-hdr {
  font-size: 9px;
  font-weight: 700;
  letter-spacing: .12em;
  text-transform: uppercase;
  margin-bottom: 4px;
}
.win-lbl  { color: var(--pos); }
.dep-lbl  { color: var(--neg); }
.ret-lbl  { color: var(--ret); }
.det-col-ul { height: 2px; margin-bottom: 6px; border-radius: 1px; }
.win-ul { background: var(--pos); }
.dep-ul { background: var(--neg); }
.ret-ul { background: var(--ret); }
.det-item {
  display: flex;
  justify-content: space-between;
  align-items: baseline;
  gap: 4px;
  padding: 2px 0;
  border-bottom: 1px solid #F1F5F9;
  font-size: 11px;
}
.det-item:last-child { border-bottom: none; }
.det-adv { color: var(--text); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex: 1; }
.det-val { font-family: var(--mono); font-size: 10px; font-weight: 600; white-space: nowrap; }
.det-empty { color: var(--muted); font-size: 11px; font-style: italic; }

/* ── COLORS ── */
.pos { color: var(--pos); }
.neg { color: var(--neg); }
.neu { color: var(--muted); }

/* ── RESPONSIVE MOBILE ── */
@media (max-width: 640px) {
  :root { --page-max: 100%; }
  .page { padding: 1.5rem 1rem; }
  .cover-layout { grid-template-columns: 1fr; }
  .cover-left { min-height: 200px; }
  .cover-bottom { grid-template-columns: 1fr; }
  .moves-grid { grid-template-columns: 1fr; }
  .bar-row { grid-template-columns: 100px 1fr 60px; }
  .ret-bar-row { grid-template-columns: 90px 1fr; }
  .detail-page { grid-template-columns: 1fr; }
  .card-body { grid-template-columns: 1fr; }
  .det-col { border-right: none; border-bottom: 1px solid var(--border); }
  .agencies-table .td-topwins,
  .agencies-table .td-topdeps,
  .agencies-table .td-topret { display: none; }
}


/* ══ RICH TEXT EDITOR TOOLBAR ══ */
#editor-toolbar {
  position: fixed;
  bottom: 0; left: 0; right: 0;
  z-index: 9999;
  background: #0F172A;
  border-top: 2px solid #38BDF8;
  padding: 6px 12px;
  display: flex;
  align-items: center;
  gap: 6px;
  flex-wrap: nowrap;
  overflow-x: auto;
  box-shadow: 0 -4px 24px rgba(0,0,0,.5);
  scrollbar-width: none;
}
#editor-toolbar::-webkit-scrollbar { display:none; }

.tb-group {
  display: flex;
  align-items: center;
  gap: 4px;
  flex-shrink: 0;
}
.tb-group-label {
  font-size: 10px;
  font-family: 'DM Mono', monospace;
  color: #475569;
  letter-spacing: .08em;
  text-transform: uppercase;
  margin-right: 2px;
  white-space: nowrap;
}
.tb-hint {
  font-size: 10px;
  color: #475569;
  font-style: italic;
  white-space: nowrap;
}
.tb-sep-v {
  width: 1px; height: 22px;
  background: #1E293B;
  flex-shrink: 0;
  margin: 0 4px;
}

/* Boutons */
.tb-btn {
  background: #1E293B;
  color: #CBD5E1;
  border: 1px solid #334155;
  padding: 4px 10px;
  border-radius: 5px;
  font-size: 12px;
  font-family: 'DM Sans', sans-serif;
  cursor: pointer;
  white-space: nowrap;
  transition: background .12s, border-color .12s;
  flex-shrink: 0;
}
.tb-btn:hover { background: #334155; border-color: #475569; }
.tb-primary    { background: #0C4A6E; border-color: #38BDF8; color: #fff; font-weight: 700; }
.tb-primary:hover { background: #075985; }
.tb-export-html { background: #064E3B; border-color: #10B981; color: #fff; font-weight: 600; }
.tb-export-html:hover { background: #065F46; }
.tb-export-pdf  { background: #4C1D95; border-color: #8B5CF6; color: #fff; font-weight: 600; }
.tb-export-pdf:hover  { background: #5B21B6; }

.tb-select {
  background: #1E293B;
  color: #CBD5E1;
  border: 1px solid #334155;
  border-radius: 4px;
  padding: 3px 5px;
  font-size: 11px;
  cursor: pointer;
}

/* Color buttons */
.tb-color-btn {
  width: 22px; height: 22px;
  border: 2px solid #334155;
  border-radius: 4px;
  cursor: pointer;
  padding: 0;
  flex-shrink: 0;
}
.tb-color-btn:hover { border-color: #38BDF8; }

/* Swatches fond cellule */
.tb-swatch {
  width: 20px; height: 20px;
  border-radius: 4px;
  border: 1px solid #334155;
  cursor: pointer;
  flex-shrink: 0;
  font-size: 10px;
  display: flex; align-items: center; justify-content: center;
  color: #475569;
  transition: transform .1s, border-color .1s;
}
.tb-swatch:hover { transform: scale(1.2); border-color: #38BDF8; }

/* Padding bas body */
body { padding-bottom: 56px; }

/* Cellule sélectionnée */
.cell-selected {
  outline: 2px solid #38BDF8 !important;
  outline-offset: -2px;
}

/* Zones éditables en mode édition */
.edit-mode .editable-zone {
  cursor: text;
}
.edit-mode .editable-zone:hover {
  outline: 1px dashed rgba(56,189,248,.5);
  outline-offset: 1px;
}
.editing-active {
  outline: 2px solid #38BDF8 !important;
  background: rgba(56,189,248,.06) !important;
}

/* Menu contextuel */
#nbb-ctx-menu {
  position: fixed;
  z-index: 99999;
  background: #1E293B;
  border: 1px solid #38BDF8;
  border-radius: 8px;
  min-width: 200px;
  box-shadow: 0 8px 24px rgba(0,0,0,.5);
  overflow: hidden;
  animation: ctxFadeIn .1s ease;
}
@keyframes ctxFadeIn {
  from { opacity:0; transform: scale(.96); }
  to   { opacity:1; transform: scale(1); }
}
.ctx-item {
  padding: 9px 14px;
  font-size: 13px;
  color: #E2E8F0;
  cursor: pointer;
  transition: background .1s;
  font-family: 'DM Sans', sans-serif;
}
.ctx-item:hover { background: #334155; }
.ctx-danger     { color: #F87171; }
.ctx-danger:hover { background: #450A0A; }
.ctx-sep {
  height: 1px;
  background: #334155;
  margin: 2px 0;
}

/* ── PRINT / A4 ── */
@media print {
  * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
  
  body { background: white; font-size: 11px; }
  
  .top-nav { display: none; }
  #editor-toolbar { display: none !important; }
  body { padding-bottom: 0; }
  
  .page {
    max-width: 100%;
    padding: 1cm 1.2cm;
    page-break-after: always;
    break-after: page;
  }
  .page:last-child { page-break-after: auto; }
  
  .detail-page {
    page-break-inside: avoid;
    break-inside: avoid;
  }
  .agency-card {
    page-break-inside: avoid;
    break-inside: avoid;
  }
  
  .page-header h2 { font-size: 16pt; }
  .data-table { font-size: 9pt; }
  .data-table thead th { font-size: 8pt; padding: 4px 6px; }
  .data-table td { padding: 4px 6px; }
  
  .table-scroll { overflow: visible; border: 1px solid #ccc; }
  
  .cover-title { font-size: 28pt; }
  .bar-fill { print-color-adjust: exact; }
  
  @page {
    size: A4 portrait;
    margin: 1.5cm 1.5cm 2cm;
  }
}
"""

# ─────────────────────────────────────────────────────────────
# MAIN BUILDER
# ─────────────────────────────────────────────────────────────


PREMIUM_CSS = """
/* ════════════════════════════════════════════
   NBB Report Premium — Design System v4
   ════════════════════════════════════════════ */
:root {{
  --ink:      #0D1117;
  --gold:     #C9A84C;
  --gold-lt:  #FEF3C7;
  --green:    #1A6B4A;
  --green-lt: #D4EDDA;
  --red:      #C0392B;
  --red-lt:   #FAD7D3;
  --slate:    #4A5568;
  --border:   #E2E8F0;
  --bg:       #FFFFFF;
  --bg2:      #FAFBFC;
  --shadow:   0 4px 6px -1px rgba(0,0,0,.07), 0 2px 4px -1px rgba(0,0,0,.04);
  --shadow-lg:0 10px 20px -4px rgba(0,0,0,.1), 0 4px 6px -2px rgba(0,0,0,.05);
  --radius:   12px;
  --mast-h:   56px;
  --rail-w:   220px;
  --publicis: #C4570A;
  --omnicom:  #B8960A;
  --dentsu:   #4A8A3C;
  --havas:    #7B3FA0;
  --wpp:      #A040A0;
  --heading:  'Playfair Display', Georgia, serif;
  --body:     'IBM Plex Sans', sans-serif;
  --mono:     'IBM Plex Mono', monospace;
}}

*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
html {{ scroll-behavior: smooth; }}
body {{
  background: #fff;
  color: var(--ink);
  font-family: var(--body);
  font-size: 14px;
  line-height: 1.6;
}}

/* ══ MASTHEAD ════════════════════════════════ */
.masthead {{
  position: sticky;
  top: 0;
  z-index: 1000;
  background: #fff;
  border-bottom: 3px solid var(--gold);
  box-shadow: 0 2px 8px -2px rgba(0,0,0,.1);
  height: var(--mast-h);
  width: 100%;
}}
.masthead-inner {{
  display: flex;
  align-items: stretch;
  height: 100%;
  width: 100%;
}}
.mast-brand {{
  padding: 0 20px;
  display: flex;
  flex-direction: column;
  justify-content: center;
  border-right: 1px solid var(--border);
  flex-shrink: 0;
  width: var(--rail-w);
}}
.mast-recma {{ font-family:var(--mono); font-size:9px; letter-spacing:.2em; text-transform:uppercase; color:var(--gold); }}
.mast-title {{ font-family:var(--heading); font-size:14px; font-weight:700; color:var(--ink); line-height:1.2; white-space:nowrap; }}
.mast-sub   {{ font-size:10px; color:var(--slate); font-family:var(--mono); }}
.mast-nav {{
  display: flex;
  align-items: stretch;
  flex: 1;
  overflow-x: auto;
  scrollbar-width: none;
}}
.mast-nav::-webkit-scrollbar {{ display: none; }}
.mast-nav a {{
  display: flex;
  align-items: center;
  padding: 0 16px;
  color: #94A3B8;
  text-decoration: none;
  font-size: 11px;
  font-weight: 500;
  letter-spacing: .05em;
  text-transform: uppercase;
  white-space: nowrap;
  border-right: 1px solid var(--border);
  border-bottom: 3px solid transparent;
  margin-bottom: -3px;
  transition: color .15s, border-color .15s, background .15s;
}}
.mast-nav a:hover  {{ color: var(--ink); background: var(--bg2); }}
.mast-nav a.active {{ color: var(--green); border-bottom-color: var(--green); font-weight: 700; background: #F0FFF4; }}
.mast-badge {{
  padding: 0 16px;
  display: flex;
  align-items: center;
  gap: 8px;
  flex-shrink: 0;
  border-left: 1px solid var(--border);
}}
.mast-market {{ font-family:var(--mono); font-size:11px; font-weight:700; color:var(--gold); background:var(--gold-lt); padding:3px 8px; border-radius:4px; }}
.mast-date   {{ font-size:10px; color:var(--slate); font-family:var(--mono); }}

/* ══ LAYOUT ══════════════════════════════════ */
.layout {{
  display: grid;
  grid-template-columns: var(--rail-w) 1fr;
  min-height: calc(100vh - var(--mast-h));
}}

/* ══ LEFT RAIL ═══════════════════════════════ */
.left-rail {{
  background: var(--bg2);
  border-right: 1px solid var(--border);
  padding: 1.5rem 1rem;
  position: sticky;
  top: var(--mast-h);
  height: calc(100vh - var(--mast-h));
  overflow-y: auto;
  scrollbar-width: thin;
  display: flex;
  flex-direction: column;
  gap: 1.25rem;
}}
.rail-section-title {{
  font-family: var(--mono); font-size: 9px; letter-spacing: .18em;
  text-transform: uppercase; color: var(--gold);
  padding-bottom: .4rem; border-bottom: 1px solid var(--border); margin-bottom: .5rem;
}}
.kpi-card {{
  background: #fff; border: 1px solid var(--border); border-radius: 10px;
  padding: .7rem .85rem; margin-bottom: .4rem;
  box-shadow: var(--shadow); transition: box-shadow .2s, transform .15s;
}}
.kpi-card:hover {{ box-shadow: var(--shadow-lg); transform: translateY(-1px); }}
.kpi-label {{ font-size: 9px; color: #94A3B8; font-family: var(--mono); letter-spacing: .06em; text-transform: uppercase; margin-bottom: 2px; }}
.kpi-value {{ font-family: var(--heading); font-size: 20px; font-weight: 700; line-height: 1.1; }}
.kpi-value.pos {{ color: var(--green); }}
.kpi-value.neg {{ color: var(--red); }}
.kpi-sub   {{ font-size: 10px; color: #94A3B8; font-family: var(--mono); margin-top: 2px; }}
.mini-bar  {{ height: 3px; background: var(--border); border-radius: 2px; margin-top: 5px; overflow: hidden; }}
.mini-bar-fill     {{ height: 100%; border-radius: 2px; }}
.mini-bar-fill.pos {{ background: var(--green); }}
.mini-bar-fill.neg {{ background: var(--red); }}
.rail-agency {{
  display: flex; align-items: center; gap: 6px; padding: 5px 0;
  border-bottom: 1px solid var(--border); font-size: 11px; cursor: pointer;
  transition: color .15s;
}}
.rail-agency:last-child {{ border-bottom: none; }}
.rail-agency:hover {{ color: var(--green); }}
.rail-rank   {{ font-family: var(--mono); font-size: 9px; color: #94A3B8; min-width: 14px; }}
.group-dot   {{ width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }}
.rail-ag-name {{ flex: 1; font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
.rail-ag-nbb  {{ font-family: var(--mono); font-size: 10px; font-weight: 700; white-space: nowrap; }}
.rail-ag-nbb.pos {{ color: var(--green); }}
.rail-ag-nbb.neg {{ color: var(--red); }}
.rail-more {{
  display: block; margin-top: .6rem; font-size: 10px; color: var(--gold);
  font-family: var(--mono); text-decoration: none; padding: 5px 8px;
  border: 1px dashed var(--gold); border-radius: 6px; text-align: center;
  transition: background .15s;
}}
.rail-more:hover {{ background: var(--gold-lt); }}

/* ══ MAIN ════════════════════════════════════ */
.main {{ background: #fff; min-width: 0; overflow: hidden; }}

/* ══ SECTIONS ════════════════════════════════ */
.section {{ padding: 2rem 2rem; border-bottom: 1px solid var(--border); }}
.section:nth-child(even) {{ background: var(--bg2); }}
.section-header {{
  display: flex; align-items: center; gap: .75rem;
  margin-bottom: 1.5rem; padding-bottom: .75rem;
  border-bottom: 2px solid var(--ink);
  flex-wrap: wrap;
}}
.section-num {{
  font-family: var(--mono); font-size: 10px; color: var(--gold);
  background: var(--gold-lt); padding: 3px 8px; border-radius: 4px;
  letter-spacing: .08em; flex-shrink: 0;
}}
.section-title {{ font-family: var(--heading); font-size: clamp(1.1rem, 2vw, 1.5rem); font-weight: 700; flex: 1; min-width: 0; }}
.section-sub   {{ font-size: 10px; color: var(--slate); font-family: var(--mono); }}

/* ══ PERIMETER ═══════════════════════════════ */
.perim-box {{
  background: #fff; border: 1px solid var(--border);
  border-left: 4px solid var(--ink); border-radius: 0 var(--radius) var(--radius) 0;
  padding: 1.1rem 1.3rem; margin-bottom: 1.25rem;
  box-shadow: 0 2px 6px -2px rgba(0,0,0,.06);
}}
.perim-box p {{ font-size: 12.5px; line-height: 1.65; color: var(--slate); margin-bottom: .45rem; }}
.perim-box p:last-child {{ margin-bottom: 0; }}
.perim-box strong {{ color: var(--ink); font-weight: 600; }}
.perim-box em {{ font-style: italic; }}

/* ══ TWO COL ══════════════════════════════════ */
.two-col {{ display: grid; grid-template-columns: 1fr 1fr; gap: 1.25rem; margin-top: 1.25rem; }}

/* ══ TABLEAUX ════════════════════════════════ */
.tbl-wrap, .chart-box {{
  background: #fff; border: 1px solid var(--border);
  border-radius: var(--radius); overflow: hidden;
  box-shadow: var(--shadow); transition: box-shadow .2s;
}}
.tbl-wrap:hover, .chart-box:hover {{ box-shadow: var(--shadow-lg); }}
.tbl-head-bar {{
  background: var(--ink); color: #fff; padding: 9px 14px;
  font-size: 10px; font-weight: 700; letter-spacing: .06em;
  text-transform: uppercase; font-family: var(--mono);
  display: flex; align-items: center; justify-content: space-between;
}}
.tbl-head-bar span {{ color: var(--gold); }}
table.ptbl {{ width: 100%; border-collapse: collapse; font-size: 12.5px; }}
table.ptbl thead th {{
  background: var(--bg2); padding: 7px 10px; text-align: left;
  font-size: 10px; font-weight: 600; font-family: var(--mono);
  letter-spacing: .05em; text-transform: uppercase; color: var(--slate);
  border-bottom: 1px solid var(--border); white-space: nowrap;
}}
table.ptbl tbody tr {{ border-bottom: 1px solid #F0F4F8; transition: background .1s; }}
table.ptbl tbody tr:last-child {{ border-bottom: none; }}
table.ptbl tbody tr:hover {{ background: var(--bg2); }}
table.ptbl td {{ padding: 7px 10px; vertical-align: middle; }}
.td-rank {{ font-family: var(--mono); font-size: 11px; color: #94A3B8; white-space: nowrap; }}
.td-ag   {{ font-weight: 600; }}
.td-mono {{ font-family: var(--mono); font-size: 12px; font-weight: 600; white-space: nowrap; }}
.pos {{ color: var(--green); }}
.neg {{ color: var(--red); }}
.neu {{ color: var(--slate); }}
.table-section-header {{
  display: flex; align-items: center; justify-content: space-between;
  padding: .6rem 0; margin-bottom: .5rem;
  font-family: var(--mono); font-size: 10px; color: var(--slate);
}}
.threshold-badge {{
  background: var(--gold-lt); color: var(--gold);
  padding: 2px 8px; border-radius: 4px; font-weight: 700;
}}

/* ══ MOVES ════════════════════════════════════ */
.moves-cols {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 1rem; }}
.moves-col-head {{
  font-family: var(--mono); font-size: 10px; font-weight: 700;
  letter-spacing: .1em; text-transform: uppercase;
  padding: 8px 10px; border-radius: var(--radius) var(--radius) 0 0;
}}
.win-head {{ background: var(--green); color: #fff; }}
.dep-head {{ background: var(--red);   color: #fff; }}
.ret-head {{ background: var(--gold);  color: var(--ink); }}
.move-card {{
  background: #fff; border: 1px solid var(--border); border-top: none;
  padding: .65rem .85rem; display: flex; flex-direction: column; gap: 2px;
  transition: box-shadow .15s;
}}
.move-card:hover {{ box-shadow: 0 2px 8px rgba(0,0,0,.08); }}
.move-card:last-of-type {{ border-radius: 0 0 var(--radius) var(--radius); }}
.mc-adv {{ font-weight: 700; font-size: 12.5px; }}
.mc-ag  {{ font-size: 11px; color: var(--slate); }}
.mc-val {{ font-family: var(--mono); font-weight: 700; font-size: 12px; align-self: flex-end; }}

/* ══ GROUPS ═══════════════════════════════════ */
.groups-list {{ display: flex; flex-direction: column; gap: .7rem; }}
.group-block {{
  background: #fff; border: 1px solid var(--border); border-radius: var(--radius);
  overflow: hidden; box-shadow: var(--shadow); transition: box-shadow .2s;
}}
.group-block:hover {{ box-shadow: var(--shadow-lg); }}
.group-block-header {{
  display: grid;
  grid-template-columns: 8px 1fr 100px 70px 70px 50px 50px;
  align-items: center; gap: 10px; padding: 10px 14px;
  cursor: pointer; transition: background .1s;
}}
.group-block-header:hover {{ background: var(--bg2); }}
.group-color-bar {{ width: 4px; height: 28px; border-radius: 2px; flex-shrink: 0; }}
.group-name  {{ font-weight: 700; font-size: 13px; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
.group-nbb   {{ font-family: var(--mono); font-weight: 700; font-size: 13px; white-space: nowrap; text-align: right; }}
.group-stat  {{ font-family: var(--mono); font-size: 11px; color: var(--slate); text-align: center; }}
.group-stat small {{ display: block; font-size: 9px; color: #94A3B8; letter-spacing: .04em; text-transform: uppercase; }}
.group-agencies     {{ display: none; padding: 0 14px 10px; border-top: 1px solid var(--border); }}
.group-agencies.open {{ display: block; }}
.agency-sub-row {{
  border-bottom: 1px solid #F0F4F8;
}}
.agency-sub-row td {{
  font-size: 11px; padding: 5px 8px; vertical-align: middle;
}}
.agency-sub-row:last-child {{ border-bottom: none; }}
.agency-sub-row:last-child {{ border-bottom: none; }}

/* ══ RETENTIONS ═══════════════════════════════ */
.ret-bars {{
  background: #fff; border: 1px solid var(--border); border-radius: var(--radius);
  padding: 1rem; box-shadow: var(--shadow);
  display: flex; flex-direction: column; gap: .45rem;
}}
.ret-row {{ display: grid; grid-template-columns: 120px 1fr 80px; align-items: center; gap: .6rem; }}
.ret-name {{ font-size: 12px; font-weight: 600; text-align: right; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
.ret-bar-track {{ height: 20px; background: var(--bg2); border-radius: 4px; overflow: hidden; }}
.ret-bar-fill {{
  height: 100%; border-radius: 4px;
  display: flex; align-items: center; justify-content: flex-end; padding-right: 6px;
  transition: width .8s ease;
}}
.ret-bar-fill span {{
  font-family: var(--mono); font-size: 9px; font-weight: 700;
  color: #fff; white-space: nowrap;
}}
.ret-val {{ font-family: var(--mono); font-size: 12px; font-weight: 700; color: var(--green); }}

/* ══ DETAIL CARDS ════════════════════════════ */
.detail-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; }}
.ag-card {{
  background: #fff; border: 1px solid var(--border); border-radius: var(--radius);
  overflow: hidden; box-shadow: var(--shadow);
  transition: box-shadow .2s, transform .2s;
}}
.ag-card:hover {{ box-shadow: var(--shadow-lg); transform: translateY(-2px); }}
.ag-card-header {{
  display: flex; align-items: center; gap: 8px;
  padding: 9px 11px; border-bottom: 1px solid var(--border);
}}
.ag-card-stripe {{ width: 4px; align-self: stretch; border-radius: 2px; flex-shrink: 0; }}
.ag-card-name   {{ font-family: var(--heading); font-size: 13px; font-weight: 700; flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
.ag-card-group  {{ font-size: 9px; color: var(--slate); font-family: var(--mono); white-space: nowrap; }}
.ag-card-nbb    {{ font-family: var(--mono); font-size: 11px; font-weight: 700; padding: 2px 7px; border-radius: 4px; white-space: nowrap; flex-shrink: 0; }}
.ag-card-nbb.pos {{ background: var(--green-lt); color: var(--green); }}
.ag-card-nbb.neg {{ background: var(--red-lt);   color: var(--red); }}
.ag-card-nbb.neu {{ background: var(--bg2);       color: var(--slate); }}
.ag-card-body   {{ display: grid; grid-template-columns: 1fr 1fr 1fr; }}
.ag-col         {{ padding: .55rem .6rem; border-right: 1px solid #F0F4F8; min-height: 60px; }}
.ag-col:last-child {{ border-right: none; }}
.ag-col-label   {{ font-family: var(--mono); font-size: 9px; font-weight: 700; letter-spacing: .1em; text-transform: uppercase; margin-bottom: 3px; padding-bottom: 3px; border-bottom: 2px solid; }}
.lbl-win {{ color: var(--green); border-color: var(--green); }}
.lbl-dep {{ color: var(--red);   border-color: var(--red); }}
.lbl-ret {{ color: var(--gold);  border-color: var(--gold); }}
.ag-item {{ display: flex; justify-content: space-between; align-items: baseline; padding: 2px 0; border-bottom: 1px solid #F8F9FA; font-size: 11px; gap: 3px; }}
.ag-item:last-child {{ border-bottom: none; }}
.ag-item-name {{ flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; min-width: 0; }}
.ag-item-val  {{ font-family: var(--mono); font-size: 10px; font-weight: 700; white-space: nowrap; flex-shrink: 0; }}
.ag-empty     {{ font-size: 11px; color: #CBD5E1; font-style: italic; }}

/* ══ TAKEAWAYS ════════════════════════════════ */
.takeaways-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: .7rem; margin-top: 1rem; }}
.takeaway-card {{
  background: #fff; border: 1px solid var(--border); border-radius: 10px;
  padding: .8rem 1rem; display: flex; gap: .7rem; align-items: flex-start;
  box-shadow: var(--shadow); transition: border-color .15s, box-shadow .15s;
}}
.takeaway-card:hover {{ border-color: var(--gold); box-shadow: var(--shadow-lg); }}
.tc-bullet {{
  width: 26px; height: 26px; border-radius: 50%; background: var(--ink);
  color: var(--gold); font-family: var(--heading); font-weight: 700; font-size: 12px;
  display: flex; align-items: center; justify-content: center; flex-shrink: 0;
}}
.tc-body {{ font-size: 12px; line-height: 1.5; }}

/* ══ ANIMATIONS ═══════════════════════════════ */
@keyframes fadeUp {{ from {{ opacity:0; transform:translateY(14px); }} to {{ opacity:1; transform:translateY(0); }} }}
.card-hidden  {{ opacity: 0; transform: translateY(14px); transition: opacity .45s ease, transform .45s ease; }}
.card-visible {{ opacity: 1; transform: translateY(0); }}

/* ══ RESPONSIVE ═══════════════════════════════ */
@media (max-width: 900px) {{
  :root {{ --rail-w: 0px; }}
  .layout {{ grid-template-columns: 1fr; }}
  .left-rail {{ display: none; }}
  .mast-brand {{ width: auto; min-width: 160px; }}
  .two-col {{ grid-template-columns: 1fr; }}
  .moves-cols {{ grid-template-columns: 1fr; }}
  .detail-grid {{ grid-template-columns: 1fr; }}
  .takeaways-grid {{ grid-template-columns: 1fr; }}
}}
@media (max-width: 600px) {{
  .section {{ padding: 1.5rem 1rem; }}
  .mast-title {{ font-size: 12px; }}
  .mast-nav a {{ padding: 0 10px; font-size: 10px; }}
  .mast-badge {{ display: none; }}
  .group-block-header {{ grid-template-columns: 8px 1fr auto; }}
  .group-stat:not(:first-of-type) {{ display: none; }}
  .ag-card-body {{ grid-template-columns: 1fr; }}
  .ag-col {{ border-right: none; border-bottom: 1px solid #F0F4F8; }}
}}

/* ══ PRINT A4 ══════════════════════════════════ */
/* ══ PITCH + SECTEUR COMPONENTS ══════════════════════════ */
.pitch-row {{
  display:flex; align-items:center; gap:8px;
  padding:5px 0; border-bottom:1px solid var(--border);
}}
.pitch-type {{
  font-family:var(--mono); font-size:9px; font-weight:700;
  letter-spacing:.08em; padding:2px 6px; border-radius:3px;
  min-width:66px; text-align:center;
}}
.pitch-type.local    {{ background:#EFF6FF; color:#2563EB; }}
.pitch-type.regional {{ background:#F5F3FF; color:#7C3AED; }}
.pitch-type.global   {{ background:#FFF7ED; color:#EA580C; }}
.pitch-count  {{ font-family:var(--heading); font-size:18px; font-weight:700; color:var(--ink); margin-left:auto; }}
.pitch-total  {{ font-size:10px; color:var(--slate); font-family:var(--mono); margin-top:6px; text-align:right; }}
.sector-row {{
  display:flex; align-items:center; gap:6px;
  padding:5px 0; border-bottom:1px solid var(--border); font-size:11px;
}}
.sector-row:last-child {{ border-bottom:none; }}
.sector-rank {{ font-family:var(--mono); font-size:9px; color:#94A3B8; min-width:12px; }}
.sector-name {{ flex:1; font-weight:500; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
.sector-val  {{ font-family:var(--mono); font-size:10px; font-weight:700; white-space:nowrap; }}

/* ══ PRINT A4 ══════════════════════════════════════════════ */
@media print {{
  * {{ -webkit-print-color-adjust:exact !important; print-color-adjust:exact !important; }}
  .masthead {{ position:fixed; top:0; left:0; right:0; z-index:9999; }}
  body {{ padding-top:56px; }}
  .left-rail {{ display:none; }}
  .layout    {{ grid-template-columns:1fr; }}
  .section   {{ page-break-after:always; break-after:page; }}
  .section:last-child {{ page-break-after:auto; }}
  .ag-card   {{ page-break-inside:avoid; break-inside:avoid; }}
  .detail-grid {{ grid-template-columns:1fr 1fr; }}
  .group-agencies {{ display:block !important; }}
  @page {{ size:A4 portrait; margin:2cm 1.5cm 1.5cm; }}
}}


/* ═══════════════════════════════════════════════
   DESIGN OVERRIDE v2 — Glass & Soft UI
   ═══════════════════════════════════════════════ */

/* Variables enrichies */
:root {{
  --border:   #E5EAF0;
  --bg2:      #F7F9FC;
  --shadow:   0 6px 20px rgba(0,0,0,.06);
  --shadow-lg:0 12px 30px rgba(0,0,0,.10);
  --radius:   16px;
}}

body {{
  background: linear-gradient(180deg, #ffffff 0%, #f7f9fc 100%);
  font-size: 14px;
  line-height: 1.6;
}}

/* ── HEADER glass ─────────────────────────────── */
.masthead {{
  background: rgba(255,255,255,0.88);
  backdrop-filter: blur(12px);
  -webkit-backdrop-filter: blur(12px);
  border-bottom: 1px solid rgba(229,234,240,0.8);
  box-shadow: 0 2px 12px rgba(0,0,0,.04);
}}
.mast-nav a {{ color: #94A3B8; }}
.mast-nav a.active {{
  background: rgba(26,107,74,.06);
  border-bottom: 2px solid var(--green);
  color: var(--ink);
  font-weight: 700;
}}

/* ── SECTIONS ─────────────────────────────────── */
.section {{ padding: 2.5rem 2.5rem; }}
.section:nth-child(even) {{ background: #FAFBFD; }}
.section-header {{ border-bottom: none; margin-bottom: 2rem; }}
.section-title  {{ font-size: clamp(1.2rem, 2.5vw, 1.75rem); letter-spacing: -.01em; }}
.section-num    {{ font-size: 10px; }}

/* ── CARDS verre dépoli ───────────────────────── */
.tbl-wrap, .chart-box {{
  border: none;
  border-radius: var(--radius);
  background: rgba(255,255,255,0.88);
  backdrop-filter: blur(6px);
  -webkit-backdrop-filter: blur(6px);
  box-shadow: var(--shadow);
}}
.tbl-wrap:hover, .chart-box:hover {{
  transform: translateY(-3px);
  box-shadow: var(--shadow-lg);
}}

/* ── KPI cards ────────────────────────────────── */
.kpi-card {{
  border: none;
  border-radius: 14px;
  background: linear-gradient(135deg, #ffffff 0%, #f9fafb 100%);
  box-shadow: var(--shadow);
}}
.kpi-card:hover {{ transform: translateY(-2px); box-shadow: var(--shadow-lg); }}
.kpi-value {{ font-size: 22px; }}

/* ── Left rail ────────────────────────────────── */
.left-rail {{
  background: linear-gradient(180deg, #F7F9FC 0%, #f0f4f8 100%);
  padding: 1.75rem 1.1rem;
  gap: 1.5rem;
}}

/* ── Tables ───────────────────────────────────── */
.tbl-head-bar {{ border-radius: var(--radius) var(--radius) 0 0; }}
table.ptbl thead th {{ background: rgba(247,249,252,0.9); }}
table.ptbl tbody tr:hover {{ background: #f0f5fb; }}

/* ── Takeaway cards ───────────────────────────── */
.takeaway-card {{
  border: none;
  background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
  box-shadow: var(--shadow);
}}
.takeaway-card:hover {{ box-shadow: var(--shadow-lg); border-color: transparent; }}
.tc-bullet {{ width: 28px; height: 28px; font-size: 13px; }}

/* ── Group blocks ─────────────────────────────── */
.group-block {{ border: none; box-shadow: var(--shadow); }}
.group-block:hover {{ box-shadow: var(--shadow-lg); }}
.group-block-header:hover {{ background: rgba(247,249,252,0.9); }}

/* ── Agency detail cards ──────────────────────── */
.ag-card {{ border: none; box-shadow: var(--shadow); }}
.ag-card:hover {{ transform: translateY(-3px); box-shadow: var(--shadow-lg); }}

/* ── Perim box ────────────────────────────────── */
.perim-box {{
  background: rgba(255,255,255,0.88);
  border: none;
  border-left: 4px solid var(--ink);
  box-shadow: var(--shadow);
  border-radius: 0 var(--radius) var(--radius) 0;
}}

/* ── Move cards ───────────────────────────────── */
.move-card {{ background: rgba(255,255,255,0.9); border: none; }}
.move-card:hover {{ box-shadow: var(--shadow); background: #fff; }}
.moves-col-head {{ border-radius: var(--radius) var(--radius) 0 0; }}

/* ── Ret bars container ───────────────────────── */
#retBars {{
  border: none !important;
  box-shadow: var(--shadow) !important;
  border-radius: var(--radius) !important;
  background: rgba(255,255,255,0.88) !important;
}}

/* ── Transitions CIBLÉES (pas de transition: all) ─ */
.tbl-wrap, .chart-box, .kpi-card, .ag-card,
.takeaway-card, .group-block, .move-card {{
  transition: transform 0.22s ease, box-shadow 0.22s ease;
}}
.mast-nav a {{ transition: color 0.15s, background 0.15s, border-color 0.15s; }}
.ag-card-header, table.ptbl tbody tr {{ transition: background 0.12s; }}
.card-hidden {{ transition: opacity 0.45s ease, transform 0.45s ease; }}


/* ══ RICH TEXT EDITOR — Panel latéral ════════════════════════ */
#editBtn {
  position:fixed; bottom:24px; right:24px; z-index:9998;
  background:var(--ink); color:#fff; border:none;
  padding:10px 18px; border-radius:10px; font-size:13px;
  font-family:'IBM Plex Sans',sans-serif; font-weight:600;
  cursor:pointer; box-shadow:0 4px 16px rgba(0,0,0,.25);
  display:flex; align-items:center; gap:6px;
  transition:background .15s, transform .1s;
}
#editBtn:hover { background:#1E293B; transform:translateY(-1px); }
#editBtn.active { background:#1A6B4A; }
body { padding-bottom:0; }

#editorPanel {
  position:fixed; right:20px; top:70px; width:260px;
  background:#fff; border-radius:14px;
  box-shadow:0 10px 40px rgba(0,0,0,.15);
  padding:16px; display:none; z-index:9999;
  border:1px solid #E2E8F0; font-family:'IBM Plex Sans',sans-serif; font-size:13px;
}
#editorPanel.visible { display:block; }

.ep-title { font-weight:700; font-size:13px; margin-bottom:12px; color:#0D1117; display:flex; align-items:center; justify-content:space-between; }
.ep-close { cursor:pointer; color:#94A3B8; font-size:16px; background:none; border:none; padding:0; }
.ep-close:hover { color:#0D1117; }
.ep-section { margin-bottom:14px; padding-bottom:14px; border-bottom:1px solid #F1F5F9; }
.ep-section:last-child { border-bottom:none; margin-bottom:0; padding-bottom:0; }
.ep-label { font-size:10px; font-weight:600; letter-spacing:.08em; text-transform:uppercase; color:#94A3B8; margin-bottom:6px; font-family:'IBM Plex Mono',monospace; }

#editInput { width:100%; min-height:70px; border:1px solid #E2E8F0; border-radius:7px; padding:8px 10px; font-size:12px; font-family:'IBM Plex Sans',sans-serif; resize:vertical; outline:none; transition:border-color .15s; }
#editInput:focus { border-color:#1A6B4A; }
#saveEdit { width:100%; padding:7px; background:#1A6B4A; color:#fff; border:none; border-radius:6px; font-size:12px; font-weight:600; cursor:pointer; margin-top:6px; font-family:'IBM Plex Sans',sans-serif; }
#saveEdit:hover { opacity:.85; }

.fmt-btns { display:flex; gap:4px; flex-wrap:wrap; margin-bottom:8px; }
.fmt-btn { background:#F8F9FA; border:1px solid #E2E8F0; border-radius:5px; padding:4px 8px; font-size:12px; cursor:pointer; font-family:'IBM Plex Sans',sans-serif; }
.fmt-btn:hover { background:#E2E8F0; }

.cell-swatches { display:flex; gap:5px; flex-wrap:wrap; margin-bottom:8px; }
.cs-dot { width:22px; height:22px; border-radius:5px; border:2px solid transparent; cursor:pointer; transition:transform .1s, border-color .1s; }
.cs-dot:hover { transform:scale(1.15); border-color:#0D1117; }

.color-row { display:flex; align-items:center; gap:8px; margin-bottom:6px; }
.color-row label { font-size:11px; color:#64748B; flex:1; }
.color-row input[type=color] { width:28px; height:24px; border:1px solid #E2E8F0; border-radius:4px; padding:1px; cursor:pointer; }

.export-btns { display:flex; gap:6px; }
.exp-btn { flex:1; padding:7px 6px; border:none; border-radius:6px; font-size:11px; font-weight:600; cursor:pointer; font-family:'IBM Plex Sans',sans-serif; }
.exp-btn:hover { opacity:.85; }
.exp-html { background:#064E3B; color:#fff; }
.exp-pdf  { background:#4C1D95; color:#fff; }

.tbl-actions { display:flex; gap:4px; flex-wrap:wrap; }
.tbl-btn { background:#F8F9FA; border:1px solid #E2E8F0; border-radius:5px; padding:4px 8px; font-size:11px; cursor:pointer; font-family:'IBM Plex Sans',sans-serif; }
.tbl-btn:hover { background:#E2E8F0; }
.tbl-btn.danger { color:#C0392B; }
.tbl-btn.danger:hover { background:#FFE4E6; }

.edit-mode [data-edit]:hover { outline:2px dashed rgba(26,107,74,.5); outline-offset:3px; border-radius:3px; cursor:pointer; }
.edit-mode [data-edit].selected { outline:2px solid #1A6B4A; outline-offset:3px; border-radius:3px; background:rgba(26,107,74,.04); }
.cell-selected { outline:2px solid #38BDF8 !important; outline-offset:-1px; }
.row-editing td { background:rgba(26,107,74,.04); }

#ctxMenu { position:fixed; z-index:99999; background:#fff; border:1px solid #E2E8F0; border-radius:10px; min-width:190px; box-shadow:0 8px 24px rgba(0,0,0,.12); overflow:hidden; font-family:'IBM Plex Sans',sans-serif; animation:ctxIn .1s ease; }
@keyframes ctxIn { from{opacity:0;transform:scale(.95)} to{opacity:1;transform:scale(1)} }
.ctx-item { padding:9px 14px; font-size:13px; cursor:pointer; transition:background .1s; display:flex; align-items:center; gap:8px; }
.ctx-item:hover { background:#F8F9FA; }
.ctx-sep { height:1px; background:#F1F5F9; }
.ctx-danger { color:#C0392B; }
.ctx-danger:hover { background:#FFF5F5; }

.save-flash { position:fixed; bottom:80px; right:24px; background:#1A6B4A; color:#fff; padding:7px 14px; border-radius:8px; font-size:12px; font-weight:600; font-family:'IBM Plex Sans',sans-serif; opacity:0; transition:opacity .3s; pointer-events:none; z-index:9997; }
.save-flash.show { opacity:1; }

@media print { #editBtn,#editorPanel,#ctxMenu,.save-flash{display:none!important} }

/* ══ PRINT A4 ══════════════════════════════════ */
/* ══ PITCH + SECTEUR COMPONENTS ══════════════════════════ */
.pitch-row {{
  display:flex; align-items:center; gap:8px;
  padding:5px 0; border-bottom:1px solid var(--border);
}}
.pitch-type {{
  font-family:var(--mono); font-size:9px; font-weight:700;
  letter-spacing:.08em; padding:2px 6px; border-radius:3px;
  min-width:66px; text-align:center;
}}
.pitch-type.local    {{ background:#EFF6FF; color:#2563EB; }}
.pitch-type.regional {{ background:#F5F3FF; color:#7C3AED; }}
.pitch-type.global   {{ background:#FFF7ED; color:#EA580C; }}
.pitch-count  {{ font-family:var(--heading); font-size:18px; font-weight:700; color:var(--ink); margin-left:auto; }}
.pitch-total  {{ font-size:10px; color:var(--slate); font-family:var(--mono); margin-top:6px; text-align:right; }}
.sector-row {{
  display:flex; align-items:center; gap:6px;
  padding:5px 0; border-bottom:1px solid var(--border); font-size:11px;
}}
.sector-row:last-child {{ border-bottom:none; }}
.sector-rank {{ font-family:var(--mono); font-size:9px; color:#94A3B8; min-width:12px; }}
.sector-name {{ flex:1; font-weight:500; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
.sector-val  {{ font-family:var(--mono); font-size:10px; font-weight:700; white-space:nowrap; }}


/* ══ ALIASES : classes "page" → classes Premium ═════════════ */
/* Les sections HTML utilisent .page, .kf-page etc. */
/* On les mappe aux styles Premium */
.page, .kf-page, .details-section {
  padding: 2rem 2rem;
  border-bottom: 1px solid var(--border);
  background: #fff;
}
.page:nth-child(even), .kf-page:nth-child(even) {
  background: #FAFBFD;
}
.page-header {
  display: flex;
  align-items: center;
  gap: .75rem;
  margin-bottom: 1.5rem;
  padding-bottom: .75rem;
  border-bottom: 2px solid var(--ink);
  flex-wrap: wrap;
}
.page-header h2 {
  font-family: var(--heading);
  font-size: clamp(1.1rem, 2vw, 1.5rem);
  font-weight: 700;
  color: var(--ink);
  flex: 1;
}
.page-note {
  font-size: 11px;
  color: var(--slate);
  margin-top: .75rem;
  font-style: italic;
}
.page-sub {
  font-size: 11px;
  color: var(--slate);
  font-family: var(--mono);
}
/* Tableaux */
.table-scroll {
  overflow-x: auto;
  border-radius: var(--radius);
  border: 1px solid var(--border);
  box-shadow: var(--shadow);
}
.data-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12.5px;
  background: #fff;
}
.data-table thead th {
  background: var(--ink);
  color: #fff;
  padding: 8px 10px;
  text-align: left;
  font-size: 10px;
  font-weight: 600;
  font-family: var(--mono);
  letter-spacing: .04em;
  text-transform: uppercase;
  white-space: nowrap;
}
.data-table tbody tr { border-bottom: 1px solid #F0F4F8; transition: background .1s; }
.data-table tbody tr:hover { background: #F8FAFB; }
.data-table td { padding: 7px 10px; vertical-align: middle; }
.td-rank  { font-family: var(--mono); font-size: 11px; color: #94A3B8; }
.td-agency, .td-ag { font-weight: 600; }
.td-nbb, .td-mono  { font-family: var(--mono); font-size: 12px; font-weight: 600; }
/* Move cards */
.moves-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 1rem; }
.moves-col-header {
  font-family: var(--mono); font-size: 10px; font-weight: 700;
  letter-spacing: .1em; text-transform: uppercase;
  padding: 8px 10px; border-radius: var(--radius) var(--radius) 0 0;
}
.win-hdr { background: var(--green); color: #fff; }
.dep-hdr { background: var(--red); color: #fff; }
.ret-hdr { background: var(--gold); color: var(--ink); }
.move-item {
  background: #fff; border: 1px solid var(--border); border-top: none;
  padding: .65rem .85rem; display: flex; flex-direction: column; gap: 2px;
  transition: box-shadow .15s;
}
.move-item:hover { box-shadow: 0 2px 8px rgba(0,0,0,.08); }
.move-item:last-of-type { border-radius: 0 0 var(--radius) var(--radius); }
/* Key Findings */
.kf-perimeter {
  border: 1px solid var(--border); border-radius: var(--radius);
  overflow: hidden; margin-bottom: 1.5rem;
  box-shadow: var(--shadow);
}
.kf-perim-header {
  background: var(--ink); color: #fff; padding: 10px 16px;
  font-size: 13px; font-weight: 700; text-align: center;
}
.kf-perim-body {
  padding: 1rem 1.25rem; background: #fff;
  display: flex; flex-direction: column; gap: .5rem;
}
.kf-perim-body p { font-size: 12.5px; line-height: 1.65; color: var(--slate); }
.kf-perim-body strong { color: var(--ink); font-weight: 600; }
.kf-two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 1.25rem; margin-top: 1.25rem; }
.kf-block {
  background: #fff; border: 1px solid var(--border);
  border-radius: var(--radius); overflow: hidden;
  box-shadow: var(--shadow); transition: box-shadow .2s;
}
.kf-block:hover { box-shadow: var(--shadow-lg); }
.kf-block-header {
  background: var(--ink); color: #fff; padding: 9px 14px;
  font-size: 11px; font-weight: 700; font-family: var(--mono);
  letter-spacing: .05em; text-transform: uppercase;
  display: flex; align-items: center; justify-content: space-between;
}
.kf-block-header .kf-bh-bold { font-weight: 900; color: var(--gold); }
.kf-takeaways-block {
  margin-top: 1.5rem; border: 1px solid var(--border);
  border-radius: var(--radius); overflow: hidden;
  box-shadow: var(--shadow);
}
.kf-takeaways-header {
  background: var(--ink); color: #fff; padding: 10px 16px;
  font-size: 13px; font-weight: 700; text-align: center;
}
.kf-takeaways-body {
  background: #fff; padding: 1rem 1.25rem;
  display: flex; flex-direction: column; gap: .55rem; min-height: 100px;
}
.takeaway {
  display: flex; gap: .5rem; padding: .45rem .75rem;
  background: var(--bg2); border-radius: 5px;
  border-left: 3px solid var(--green); font-size: 12.5px;
}
/* Agency detail cards */
.detail-page { display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-bottom: 1.5rem; }
.agency-card {
  background: #fff; border: 1px solid var(--border);
  border-radius: var(--radius); overflow: hidden;
  box-shadow: var(--shadow); transition: box-shadow .2s, transform .2s;
}
.agency-card:hover { box-shadow: var(--shadow-lg); transform: translateY(-2px); }
.card-body { display: grid; grid-template-columns: 1fr 1fr 1fr; }
.det-col { padding: .55rem .6rem; border-right: 1px solid #F0F4F8; }
.det-col:last-child { border-right: none; }
.det-col-hdr { font-family: var(--mono); font-size: 9px; font-weight: 700; letter-spacing: .1em; text-transform: uppercase; margin-bottom: 3px; padding-bottom: 3px; border-bottom: 2px solid; }
.win-lbl { color: var(--green); border-color: var(--green); }
.dep-lbl { color: var(--red); border-color: var(--red); }
.ret-lbl { color: var(--gold); border-color: var(--gold); }
.det-item { display: flex; justify-content: space-between; align-items: baseline; padding: 2px 0; border-bottom: 1px solid #F8F9FA; font-size: 11px; gap: 3px; }
.det-item:last-child { border-bottom: none; }
.det-adv { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.det-val { font-family: var(--mono); font-size: 10px; font-weight: 700; white-space: nowrap; }
.det-empty { font-size: 11px; color: #CBD5E1; font-style: italic; }
/* Ret bars CSS */
.ret-chart { display: flex; flex-direction: column; gap: .4rem; margin-bottom: 1.5rem; }
.ret-bar-row { display: grid; grid-template-columns: 130px 1fr 80px; align-items: center; gap: .6rem; }
.ret-bar-label { font-size: 12px; font-weight: 600; text-align: right; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.ret-bar-track { height: 20px; background: var(--bg2); border-radius: 4px; overflow: hidden; }
.ret-bar-fill { height: 100%; border-radius: 4px; display: flex; align-items: center; justify-content: flex-end; padding-right: 5px; transition: width 1.2s cubic-bezier(.16,1,.3,1); }
.ret-bar-fill span { font-family: var(--mono); font-size: 9px; font-weight: 700; color: #fff; }
.ret-val { font-family: var(--mono); font-size: 12px; font-weight: 700; color: var(--green); }
/* Two col */
.two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 1.25rem; margin-top: 1.25rem; }
/* Responsive */
@media (max-width: 768px) {
  .kf-two-col, .two-col, .moves-grid, .detail-page { grid-template-columns: 1fr; }
  .card-body { grid-template-columns: 1fr; }
  .det-col { border-right: none; border-bottom: 1px solid #F0F4F8; }
}


/* ══ GRAPHIQUES CSS (bar-chart) ══════════════════════════════ */
.chart-section { margin-bottom: 1.5rem; }
.chart-legend {
  display: flex; justify-content: space-between;
  font-family: var(--mono); font-size: 10px; color: var(--slate);
  margin-bottom: .5rem; padding: 0 4px;
}
.leg.pos { color: var(--green); font-weight: 600; }
.leg.neg { color: var(--red);   font-weight: 600; }

.bar-chart {
  background: #fff;
  border: 1px solid var(--border);
  border-radius: var(--radius);
  overflow: hidden;
  box-shadow: var(--shadow);
}
.bar-row {
  display: grid;
  grid-template-columns: 140px 1fr 90px;
  align-items: center;
  gap: .5rem;
  padding: 5px 10px;
  border-bottom: 1px solid #F0F4F8;
  transition: background .1s;
}
.bar-row:last-child { border-bottom: none; }
.bar-row:hover { background: var(--bg2); }
.bar-label {
  font-size: 11px; font-weight: 600;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  color: var(--ink);
}
.bar-track {
  height: 16px;
  background: #F1F5F9;
  border-radius: 4px;
  overflow: hidden;
  position: relative;
}
.bar-fill {
  height: 100%;
  border-radius: 4px;
  min-width: 2px;
  position: relative;
}
.bar-fill.pos {
  background: linear-gradient(90deg, #86EFAC 0%, #1A6B4A 100%);
}
.bar-fill.neg {
  background: linear-gradient(90deg, #FCA5A5 0%, #C0392B 100%);
}
.bar-val {
  font-family: var(--mono); font-size: 11px; font-weight: 700;
  white-space: nowrap; text-align: right;
}
.bar-val.pos { color: var(--green); }
.bar-val.neg { color: var(--red); }

/* Cellule nom agence avec couleur groupe */
.bar-label-cell {
  display: flex; align-items: center; gap: 5px;
}
.bar-group-dot {
  width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0;
}

/* ══ TABLEAUX DÉTAILLÉS ═══════════════════════════════════════ */
.groups-table td, .groups-table th { padding: 7px 10px; vertical-align: middle; }
.group-row td { font-size: 13px; }
.total-row td {
  background: #F8F9FA !important;
  border-top: 2px solid var(--border);
  font-size: 12px;
}
.td-group   { font-weight: 700; }
.td-num     { font-family: var(--mono); font-size: 12px; text-align: right; white-space: nowrap; }
.td-wins    { font-family: var(--mono); font-size: 12px; color: var(--green); font-weight: 700; text-align: right; white-space: nowrap; }
.td-deps    { font-family: var(--mono); font-size: 12px; color: var(--red);   font-weight: 700; text-align: right; white-space: nowrap; }
.td-nbb     { font-family: var(--mono); font-size: 12px; font-weight: 700; text-align: right; white-space: nowrap; }
.td-nbb.pos { color: var(--green); }
.td-nbb.neg { color: var(--red); }
.td-topwins, .td-topdeps, .td-topclient, .td-agency, .td-agency-sub {
  font-size: 11px; vertical-align: middle;
}
.td-agency     { font-weight: 600; }
.td-agency-sub { color: var(--slate); padding-left: 20px !important; }

/* Fond couleur groupe sur les lignes */
.group-row td:first-child  { font-weight: 700; }
.agency-sub-row { font-size: 12px; }
.agency-sub-row td { padding: 5px 10px; }

/* Ret bars alignement */
.ret-bar-row {
  grid-template-columns: 130px 1fr 80px;
}
.ret-bar-label {
  font-size: 11px; font-weight: 600;
  text-align: right; white-space: nowrap;
  overflow: hidden; text-overflow: ellipsis;
}
.ret-bar-track { background: #F1F5F9; }
.ret-bar-fill.pos {
  background: linear-gradient(90deg, #86EFAC 0%, #1A6B4A 100%);
}

/* Année */
.year { font-size: .75em; color: var(--slate); font-weight: 400; margin-left: .25rem; }

/* Pos/Neg couleurs */
.pos { color: var(--green); }
.neg { color: var(--red); }
.neu { color: var(--slate); }


/* ══ FIX : Groups table ══════════════════════════════════════ */
.group-row {
  border-bottom: 1px solid rgba(0,0,0,.06);
}
.group-row td:first-child {
  border-left: 4px solid transparent; /* remplacé par style inline */
  font-weight: 700;
}
.agency-sub-row td { font-size: 11px; color: var(--slate); }
.total-row { border-top: 2px solid var(--border) !important; }
.total-row td { font-weight: 700; background: var(--bg2) !important; }
/* Largeur colonnes groupe */
.groups-table .td-group  { min-width: 160px; }
.groups-table .td-num    { width: 60px; text-align: right; }
.groups-table .td-nbb    { width: 110px; text-align: right; }
.groups-table .td-wins   { width: 90px; text-align: right; }
.groups-table .td-deps   { width: 90px; text-align: right; }

/* ══ FIX : Retention bars ════════════════════════════════════ */
.ret-chart {
  display: flex; flex-direction: column; gap: .4rem;
  background: #fff; border: 1px solid var(--border);
  border-radius: var(--radius); padding: 1rem;
  box-shadow: var(--shadow); margin-bottom: 1.5rem;
}
.ret-bar-row {
  display: grid;
  grid-template-columns: 140px 1fr 80px;
  align-items: center; gap: .6rem;
}
.ret-bar-label {
  font-size: 11px; font-weight: 600;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  text-align: left; /* NOM À GAUCHE */
  padding: 3px 6px;
  border-radius: 4px;
  /* Le border-left + background couleur groupe viennent du style inline */
}
.ret-bar-track {
  height: 22px; background: #F1F5F9;
  border-radius: 6px; overflow: hidden; position: relative;
}
.ret-bar-fill {
  height: 100%; border-radius: 6px;
  /* Le background couleur groupe vient du style inline — ne pas écraser */
   transition: width .8s ease;
}
/* Dégradé sur la barre : on ajoute un pseudo-element transparent par-dessus */
.ret-bar-fill::after {
  content: '';
  position: absolute; inset: 0;
  background: linear-gradient(90deg, rgba(255,255,255,.3) 0%, transparent 60%);
  border-radius: 6px;
}
.ret-bar-fill span {
  position: relative; z-index: 1;
  font-family: var(--mono); font-size: 9px; font-weight: 700;
  color: #fff; float: right; padding-right: 6px; line-height: 22px;
}
.ret-val {
  font-family: var(--mono); font-size: 12px; font-weight: 700;
  color: var(--green); text-align: right;
}

/* ══ FIX : Agency detail cards ═══════════════════════════════ */
.agency-card {
  background: #fff; border: 1px solid var(--border);
  border-radius: var(--radius); overflow: hidden;
  box-shadow: var(--shadow); transition: box-shadow .2s, transform .2s;
  break-inside: avoid;
}
.agency-card:hover { box-shadow: var(--shadow-lg); transform: translateY(-2px); }

.card-header {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 9px 12px;
  border-bottom: 1px solid var(--border);
  /* background et border-bottom-color viennent du style inline */
}
.card-agency-name {
  font-family: var(--heading);
  font-size: 13px; font-weight: 800;
  flex: 1; min-width: 0;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.card-group {
  font-size: 9px; font-family: var(--mono);
  white-space: nowrap; flex-shrink: 0;
}
.card-badge {
  font-family: var(--mono); font-size: 11px; font-weight: 700;
  padding: 3px 8px; border-radius: 5px;
  white-space: nowrap; flex-shrink: 0;
}
.card-badge.pos { background: var(--green-lt); color: var(--green); }
.card-badge.neg { background: var(--red-lt);   color: var(--red); }
.card-badge.neu { background: var(--bg2);       color: var(--slate); }

.card-body {
  display: grid;
  grid-template-columns: 1fr 1fr 1fr;
}
.det-col {
  padding: .55rem .65rem;
  border-right: 1px solid #F0F4F8;
  min-height: 60px;
}
.det-col:last-child { border-right: none; }
.det-col-hdr {
  font-family: var(--mono); font-size: 9px; font-weight: 700;
  letter-spacing: .08em; text-transform: uppercase;
  margin-bottom: 4px; padding-bottom: 3px; border-bottom: 2px solid;
}
.win-lbl { color: var(--green); border-color: var(--green); }
.dep-lbl { color: var(--red);   border-color: var(--red); }
.ret-lbl { color: var(--gold);  border-color: var(--gold); }

.det-col-ul { display: flex; flex-direction: column; gap: 1px; }
.det-item {
  display: flex; justify-content: space-between; align-items: baseline;
  padding: 2px 0; border-bottom: 1px solid #F8F9FA; font-size: 11px; gap: 4px;
}
.det-item:last-child { border-bottom: none; }
.det-adv {
  flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  min-width: 0;
}
.det-val {
  font-family: var(--mono); font-size: 10px; font-weight: 700;
  white-space: nowrap; flex-shrink: 0;
}
.det-empty { font-size: 10px; color: #CBD5E1; font-style: italic; }

/* Detail grid 2 colonnes */
.detail-page {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 1rem; margin-bottom: 1.5rem;
}
@media (max-width: 768px) {
  .detail-page { grid-template-columns: 1fr; }
  .card-body    { grid-template-columns: 1fr; }
  .det-col      { border-right: none; border-bottom: 1px solid #F0F4F8; }
}


/* ══ COVER : Group ranking chart (grk-*) ════════════════════ */
.kf-chart {
  padding: .75rem;
  position: relative;
}
.grk-zero-line {
  position: absolute;
  left: 50%; top: .5rem; bottom: .5rem;
  width: 1px; background: #E2E8F0;
  pointer-events: none;
}
.grk-row {
  display: grid;
  grid-template-columns: 90px 1fr 55px;
  align-items: center;
  gap: .4rem;
  padding: 4px 0;
  border-bottom: 1px solid #F0F4F8;
}
.grk-row:last-child { border-bottom: none; }
.grk-label {
  font-size: 11px; font-weight: 700;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  text-align: right; padding-right: 6px;
  color: var(--ink);
}
.grk-track {
  height: 18px;
  background: #F1F5F9;
  border-radius: 4px;
  overflow: hidden;
  position: relative;
  display: flex;
  align-items: center;
}
.grk-bar {
  height: 100%;
  border-radius: 4px;
  position: absolute;
  transition: width .6s ease;
}
.grk-bar.grk-pos {
  background: linear-gradient(90deg, #86EFAC 0%, #1A6B4A 100%);
  left: 50%;
}
.grk-bar.grk-neg {
  background: linear-gradient(90deg, #C0392B 0%, #FCA5A5 100%);
  right: 50%;
}
.grk-val {
  font-family: var(--mono); font-size: 10px; font-weight: 700;
  white-space: nowrap; text-align: right;
}
.grk-val.pos { color: var(--green); }
.grk-val.neg { color: var(--red); }

/* ══ COVER : kf-table ════════════════════════════════════════ */
.kf-table-wrap { overflow-x: auto; }
.kf-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12px;
}
.kf-table th {
  background: var(--bg2);
  padding: 6px 8px;
  text-align: left;
  font-size: 10px; font-weight: 700;
  font-family: var(--mono);
  letter-spacing: .04em; text-transform: uppercase;
  color: var(--slate);
  border-bottom: 1px solid var(--border);
  white-space: nowrap;
}
.kf-table td {
  padding: 6px 8px;
  border-bottom: 1px solid #F0F4F8;
  vertical-align: middle;
  font-size: 12px;
}
.kf-table tbody tr:last-child td { border-bottom: none; }
.kf-table tbody tr:hover { background: var(--bg2); }
.kf-th-green { color: var(--green) !important; }
.kf-th-red   { color: var(--red) !important; }
.kf-th-market { min-width: 80px; }
.kf-bh-sub { font-size: 10px; color: #94A3B8; font-weight: 400; }
.kf-dot    { color: var(--gold); margin: 0 .25rem; }

/* ══ GROUPS TABLE : alignement colonnes ══════════════════════ */
/* Forcer display:table pour border-left sur les lignes */
.groups-table { border-collapse: separate; border-spacing: 0 2px; }
.groups-table .group-row td:first-child {
  position: relative;
}
/* Le border-left est sur le TR via style inline — on le reporte sur la 1ère td */
.groups-table tr[style*="border-left"] { border-left: none !important; }
.groups-table tr[style*="border-left"] > td:first-child {
  box-shadow: inset 4px 0 0 currentColor;
}
/* Alignement colonnes */
.groups-table th, .groups-table td { 
  white-space: nowrap;
  vertical-align: middle;
}
.groups-table th:first-child,
.groups-table td:first-child { white-space: normal; min-width: 140px; }
.td-agency-sub {
  padding-left: 24px !important;
  font-size: 11px; color: var(--slate);
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  max-width: 160px;
}


/* ══ Section headers avec numéro ════════════════════════════ */
.section-header-inner {
  display: flex;
  align-items: center;
  gap: .75rem;
  flex-wrap: wrap;
}
.section-num {
  font-family: var(--mono);
  font-size: 10px;
  font-weight: 700;
  color: var(--gold);
  background: var(--gold-lt);
  padding: 3px 9px;
  border-radius: 5px;
  letter-spacing: .08em;
  flex-shrink: 0;
}
.section-header-inner h2 {
  font-family: var(--heading);
  font-size: clamp(1.1rem, 2.5vw, 1.5rem);
  font-weight: 700;
  margin: 0;
}
.kf-sub-title {
  color: var(--slate);
  font-weight: 400;
}


/* ══ Retention chart title ══════════════════════════════════ */
.ret-chart-title {{
  font-family: var(--mono);
  font-size: 10px;
  font-weight: 700;
  letter-spacing: .08em;
  text-transform: uppercase;
  color: var(--slate);
  margin-bottom: .6rem;
  padding: 0 2px;
}}

/* ══ Ret bar layout ══════════════════════════════════════════ */
.ret-bar-row {{
  display: grid;
  grid-template-columns: 130px 1fr 75px;
  align-items: center;
  gap: .6rem;
  margin-bottom: 4px;
}}
.ret-bar-label {{
  font-size: 11px; font-weight: 600;
  padding: 3px 6px; border-radius: 4px;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  text-align: left;
}}
.ret-bar-track {{
  height: 18px; background: #F1F5F9;
  border-radius: 5px; overflow: hidden;
}}
.ret-bar-fill {{
  height: 100%; border-radius: 5px;
  /* background et width viennent du style inline */
  transition: width .8s ease;
}}
.ret-val {{
  font-family: var(--mono); font-size: 12px; font-weight: 700;
  text-align: right; white-space: nowrap;
}}

/* ══ PRINT A4 ══════════════════════════════════════════════ */
@media print {{
  * {{ -webkit-print-color-adjust:exact !important; print-color-adjust:exact !important; }}
  .masthead {{ position:fixed; top:0; left:0; right:0; z-index:9999; }}
  body {{ padding-top:56px; }}
  .left-rail {{ display:none; }}
  .layout    {{ grid-template-columns:1fr; }}
  .section   {{ page-break-after:always; break-after:page; }}
  .section:last-child {{ page-break-after:auto; }}
  .ag-card   {{ page-break-inside:avoid; break-inside:avoid; }}
  .detail-grid {{ grid-template-columns:1fr 1fr; }}
  .group-agencies {{ display:block !important; }}
  @page {{ size:A4 portrait; margin:2cm 1.5cm 1.5cm; }}
}}


/* ═══════════════════════════════════════════════
   DESIGN OVERRIDE v2 — Glass & Soft UI
   ═══════════════════════════════════════════════ */

/* Variables enrichies */
:root {{
  --border:   #E5EAF0;
  --bg2:      #F7F9FC;
  --shadow:   0 6px 20px rgba(0,0,0,.06);
  --shadow-lg:0 12px 30px rgba(0,0,0,.10);
  --radius:   16px;
}}

body {{
  background: linear-gradient(180deg, #ffffff 0%, #f7f9fc 100%);
  font-size: 14px;
  line-height: 1.6;
}}

/* ── HEADER glass ─────────────────────────────── */
.masthead {{
  background: rgba(255,255,255,0.88);
  backdrop-filter: blur(12px);
  -webkit-backdrop-filter: blur(12px);
  border-bottom: 1px solid rgba(229,234,240,0.8);
  box-shadow: 0 2px 12px rgba(0,0,0,.04);
}}
.mast-nav a {{ color: #94A3B8; }}
.mast-nav a.active {{
  background: rgba(26,107,74,.06);
  border-bottom: 2px solid var(--green);
  color: var(--ink);
  font-weight: 700;
}}

/* ── SECTIONS ─────────────────────────────────── */
.section {{ padding: 2.5rem 2.5rem; }}
.section:nth-child(even) {{ background: #FAFBFD; }}
.section-header {{ border-bottom: none; margin-bottom: 2rem; }}
.section-title  {{ font-size: clamp(1.2rem, 2.5vw, 1.75rem); letter-spacing: -.01em; }}
.section-num    {{ font-size: 10px; }}

/* ── CARDS verre dépoli ───────────────────────── */
.tbl-wrap, .chart-box {{
  border: none;
  border-radius: var(--radius);
  background: rgba(255,255,255,0.88);
  backdrop-filter: blur(6px);
  -webkit-backdrop-filter: blur(6px);
  box-shadow: var(--shadow);
}}
.tbl-wrap:hover, .chart-box:hover {{
  transform: translateY(-3px);
  box-shadow: var(--shadow-lg);
}}

/* ── KPI cards ────────────────────────────────── */
.kpi-card {{
  border: none;
  border-radius: 14px;
  background: linear-gradient(135deg, #ffffff 0%, #f9fafb 100%);
  box-shadow: var(--shadow);
}}
.kpi-card:hover {{ transform: translateY(-2px); box-shadow: var(--shadow-lg); }}
.kpi-value {{ font-size: 22px; }}

/* ── Left rail ────────────────────────────────── */
.left-rail {{
  background: linear-gradient(180deg, #F7F9FC 0%, #f0f4f8 100%);
  padding: 1.75rem 1.1rem;
  gap: 1.5rem;
}}

/* ── Tables ───────────────────────────────────── */
.tbl-head-bar {{ border-radius: var(--radius) var(--radius) 0 0; }}
table.ptbl thead th {{ background: rgba(247,249,252,0.9); }}
table.ptbl tbody tr:hover {{ background: #f0f5fb; }}

/* ── Takeaway cards ───────────────────────────── */
.takeaway-card {{
  border: none;
  background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
  box-shadow: var(--shadow);
}}
.takeaway-card:hover {{ box-shadow: var(--shadow-lg); border-color: transparent; }}
.tc-bullet {{ width: 28px; height: 28px; font-size: 13px; }}

/* ── Group blocks ─────────────────────────────── */
.group-block {{ border: none; box-shadow: var(--shadow); }}
.group-block:hover {{ box-shadow: var(--shadow-lg); }}
.group-block-header:hover {{ background: rgba(247,249,252,0.9); }}

/* ── Agency detail cards ──────────────────────── */
.ag-card {{ border: none; box-shadow: var(--shadow); }}
.ag-card:hover {{ transform: translateY(-3px); box-shadow: var(--shadow-lg); }}

/* ── Perim box ────────────────────────────────── */
.perim-box {{
  background: rgba(255,255,255,0.88);
  border: none;
  border-left: 4px solid var(--ink);
  box-shadow: var(--shadow);
  border-radius: 0 var(--radius) var(--radius) 0;
}}

/* ── Move cards ───────────────────────────────── */
.move-card {{ background: rgba(255,255,255,0.9); border: none; }}
.move-card:hover {{ box-shadow: var(--shadow); background: #fff; }}
.moves-col-head {{ border-radius: var(--radius) var(--radius) 0 0; }}

/* ── Ret bars container ───────────────────────── */
#retBars {{
  border: none !important;
  box-shadow: var(--shadow) !important;
  border-radius: var(--radius) !important;
  background: rgba(255,255,255,0.88) !important;
}}

/* ── Transitions CIBLÉES (pas de transition: all) ─ */
.tbl-wrap, .chart-box, .kpi-card, .ag-card,
.takeaway-card, .group-block, .move-card {{
  transition: transform 0.22s ease, box-shadow 0.22s ease;
}}
.mast-nav a {{ transition: color 0.15s, background 0.15s, border-color 0.15s; }}
.ag-card-header, table.ptbl tbody tr {{ transition: background 0.12s; }}
.card-hidden {{ transition: opacity 0.45s ease, transform 0.45s ease; }}

/* ══ RICH TEXT EDITOR TOOLBAR ═══════════════════════════════ */
#editor-toolbar {{
  position: fixed;
  bottom: 0; left: 0; right: 0;
  z-index: 9999;
  background: rgba(13,17,23,0.95);
  backdrop-filter: blur(10px);
  -webkit-backdrop-filter: blur(10px);
  border-top: 2px solid var(--gold);
  padding: 7px 14px;
  display: flex;
  align-items: center;
  gap: 6px;
  flex-wrap: nowrap;
  overflow-x: auto;
  scrollbar-width: none;
  box-shadow: 0 -4px 20px rgba(0,0,0,.3);
}}
#editor-toolbar::-webkit-scrollbar {{ display:none; }}
body {{ padding-bottom: 56px; }}

.tb-group {{ display:flex; align-items:center; gap:4px; flex-shrink:0; }}
.tb-sep-v {{ width:1px; height:22px; background:#333; margin:0 4px; flex-shrink:0; }}
.tb-group-label {{ font-size:9px; font-family:'IBM Plex Mono',monospace; color:#475569; letter-spacing:.1em; text-transform:uppercase; white-space:nowrap; }}
.tb-hint {{ font-size:10px; color:#475569; font-style:italic; white-space:nowrap; }}

.tb-btn {{
  background: #1E293B; color: #CBD5E1;
  border: 1px solid #334155; padding: 4px 10px;
  border-radius: 6px; font-size:12px; font-family:'IBM Plex Sans',sans-serif;
  cursor:pointer; white-space:nowrap; flex-shrink:0;
  transition: background .12s, border-color .12s;
}}
.tb-btn:hover {{ background:#334155; border-color:#475569; }}
.tb-primary    {{ background:#0C4A6E; border-color:#38BDF8; color:#fff; font-weight:700; }}
.tb-primary:hover {{ background:#075985; }}
.tb-export-html {{ background:#064E3B; border-color:#10B981; color:#fff; font-weight:600; }}
.tb-export-pdf  {{ background:#4C1D95; border-color:#8B5CF6; color:#fff; font-weight:600; }}

.tb-select {{ background:#1E293B; color:#CBD5E1; border:1px solid #334155; border-radius:4px; padding:3px 5px; font-size:11px; cursor:pointer; }}
.tb-color-btn {{ width:22px; height:22px; border:2px solid #334155; border-radius:4px; cursor:pointer; padding:0; flex-shrink:0; }}
.tb-color-btn:hover {{ border-color:#C9A84C; }}
.tb-swatch {{
  width:20px; height:20px; border-radius:4px; border:1px solid #334155;
  cursor:pointer; flex-shrink:0; font-size:10px;
  display:flex; align-items:center; justify-content:center; color:#475569;
  transition: transform .1s, border-color .1s;
}}
.tb-swatch:hover {{ transform:scale(1.2); border-color:#C9A84C; }}

.cell-selected {{ outline:2px solid #38BDF8 !important; outline-offset:-1px; }}
.edit-mode .editable-zone:hover {{ outline:1px dashed rgba(56,189,248,.4); outline-offset:1px; border-radius:2px; }}
.editing-active {{ outline:2px solid #38BDF8 !important; background:rgba(56,189,248,.05) !important; }}

#nbb-ctx-menu {{
  position:fixed; z-index:99999;
  background:#1E293B; border:1px solid #38BDF8;
  border-radius:10px; min-width:200px;
  box-shadow:0 8px 24px rgba(0,0,0,.5); overflow:hidden;
  animation: ctxIn .1s ease;
}}
@keyframes ctxIn {{ from{{opacity:0;transform:scale(.95)}} to{{opacity:1;transform:scale(1)}} }}
.ctx-item {{ padding:9px 14px; font-size:13px; color:#E2E8F0; cursor:pointer; transition:background .1s; font-family:'IBM Plex Sans',sans-serif; }}
.ctx-item:hover {{ background:#334155; }}
.ctx-danger {{ color:#F87171; }}
.ctx-danger:hover {{ background:#450A0A; }}
.ctx-sep {{ height:1px; background:#334155; margin:2px 0; }}

@media print {{
  #editor-toolbar {{ display:none !important; }}
  body {{ padding-bottom: 0; }}
}.grk-track {{
  height: 18px; background: #F1F5F9; border-radius: 4px;
  overflow: hidden; position: relative;
}}
.grk-track::before {{
  content: ''; position: absolute;
  left: 50%; top: 0; bottom: 0; width: 1px;
  background: rgba(0,0,0,.12); z-index: 1;
}}
.grk-bar {{
  height: 100%; border-radius: 4px;
  position: absolute; top: 0; transition: width .6s ease;
}}
.grk-bar.grk-pos {{
  background: linear-gradient(90deg, #86EFAC 0%, #1A6B4A 100%);
  left: 50%;
}}
.grk-bar.grk-neg {{
  background: linear-gradient(90deg, #C0392B 0%, #FCA5A5 100%);
  right: 50%; left: auto;
}}
}
"""

PREMIUM_JS_STATIC = """
function fmtNBB(v) {
  if (v === 0) return '$0m';
  return (v > 0 ? '+' : '') + v.toFixed(1) + 'm$';
}
function fmtVal(v) {
  if (!v && v !== 0) return '';
  return (v > 0 ? '+' : '') + v.toFixed(1) + 'm';
}
function nbbClass(v) { return v > 0 ? 'pos' : v < 0 ? 'neg' : 'neu'; }
function trunc(s, n=20) { return s && s.length > n ? s.slice(0,n-1)+'…' : (s||''); }

// ══════════════════════════════════════════
// SVG GROUP CHART (horizontal bars centered)
// ══════════════════════════════════════════
function drawGroupChart() {
  const svg = document.getElementById('groupChart');
  if (!svg) return;

  // Layout : nom à GAUCHE | barre vers la droite
  const W      = 500;
  const barH   = 36;
  const gap    = 12;
  const NAME_W = 130;
  const BAR_START = NAME_W + 8;
  const BAR_ZONE  = W - BAR_START - 80;
  const PAD_T  = 14;
  const H      = PAD_T + GROUPS.length * (barH + gap) + 12;

  const maxAbs = Math.max(...GROUPS.map(g => Math.abs(g.nbb)));

  let defs = '<defs>';
  // Dégradés verts + rouges globaux
  defs += `<linearGradient id="gr_pos" x1="0%" y1="0%" x2="100%" y2="0%">
    <stop offset="0%" stop-color="#86EFAC" stop-opacity="0.5"/>
    <stop offset="100%" stop-color="#1A6B4A" stop-opacity="1"/>
  </linearGradient>`;
  defs += `<linearGradient id="gr_neg" x1="0%" y1="0%" x2="100%" y2="0%">
    <stop offset="0%" stop-color="#FCA5A5" stop-opacity="0.5"/>
    <stop offset="100%" stop-color="#C0392B" stop-opacity="1"/>
  </linearGradient>`;
  // Dégradés couleur groupe pour la cellule nom
  GROUPS.forEach(g => {
    defs += `<linearGradient id="gg${g.rank}" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="${g.bg || g.color+'44'}" stop-opacity="1"/>
      <stop offset="100%" stop-color="${g.bg || g.color+'22'}" stop-opacity="1"/>
    </linearGradient>`;
  });
  defs += '</defs>';

  let s = defs;

  // Ligne d'axe
  s += `<line x1="${BAR_START}" y1="${PAD_T - 6}" x2="${W - 10}" y2="${PAD_T - 6}" stroke="#E5EAF0" stroke-width="1"/>`;

  GROUPS.forEach((g, i) => {
    const y     = PAD_T + i * (barH + gap);
    const cy    = y + barH / 2;
    const isPos = g.nbb >= 0;
    const barW  = Math.abs(g.nbb) / maxAbs * BAR_ZONE * 0.92;
    const shortName = g.name.replace(' Media Network','').replace(' Media','');

    // Fond alterné
    s += `<rect x="0" y="${y - 4}" width="${W}" height="${barH + 8}"
      fill="${i % 2 === 0 ? 'rgba(247,249,252,0.8)' : 'transparent'}" rx="4"/>`;

    // Cellule nom — fond couleur groupe
    s += `<rect x="0" y="${y}" width="${NAME_W}" height="${barH}"
      fill="url(#gg${g.rank})" rx="5"/>`;
    s += `<rect x="0" y="${y}" width="4" height="${barH}"
      fill="${g.color}" rx="2"/>`;
    // Rang + Nom
    s += `<text x="10" y="${cy + 4}"
      font-family="IBM Plex Mono" font-size="9" fill="${g.color}" font-weight="700">#${g.rank}</text>`;
    s += `<text x="${NAME_W - 8}" y="${cy + 5}"
      text-anchor="end" font-family="IBM Plex Sans" font-size="13" font-weight="700"
      fill="#1E293B">${shortName}</text>`;

    // Barre — vert ou rouge dégradé
    s += `<rect x="${BAR_START}" y="${y + 4}" width="${Math.max(barW, 3)}" height="${barH - 8}"
      fill="url(#gr_${isPos ? 'pos' : 'neg'})" rx="4"/>`;

    // Valeur
    s += `<text x="${BAR_START + barW + 8}" y="${cy + 5}"
      text-anchor="start" font-family="IBM Plex Mono" font-size="13" font-weight="700"
      fill="${isPos ? '#1A6B4A' : '#C0392B'}">${isPos ? '+' : ''}${g.nbb.toFixed(0)}m</text>`;
  });

  svg.innerHTML = s;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  svg.style.height = H + 'px';
  svg.style.width  = '100%';
}

function drawAgencyChart() {
  const svg = document.getElementById('agencyChart');
  if (!svg) return;

  // Layout : nom à GAUCHE | barre vers la droite (pos=vert, neg=rouge)
  const W      = 700;
  const barH   = 14;
  const gap    = 6;
  const NAME_W = 150;   // zone nom à gauche
  const BAR_START = NAME_W + 8;
  const BAR_ZONE  = W - BAR_START - 80; // zone des barres
  const PAD_T  = 10;
  const H      = PAD_T + AGENCIES.length * (barH + gap) + 10;

  const maxPos = Math.max(...AGENCIES.map(a => a.nbb > 0 ? a.nbb : 0));
  const maxNeg = Math.max(...AGENCIES.map(a => a.nbb < 0 ? Math.abs(a.nbb) : 0));
  const maxAbs = Math.max(maxPos, maxNeg);

  let defs = '<defs>';
  // Dégradé vert (positif) : vert clair → vert foncé
  defs += `<linearGradient id="ag_pos" x1="0%" y1="0%" x2="100%" y2="0%">
    <stop offset="0%" stop-color="#86EFAC" stop-opacity="0.6"/>
    <stop offset="100%" stop-color="#1A6B4A" stop-opacity="1"/>
  </linearGradient>`;
  // Dégradé rouge (négatif) : rouge clair → rouge foncé
  defs += `<linearGradient id="ag_neg" x1="0%" y1="0%" x2="100%" y2="0%">
    <stop offset="0%" stop-color="#FCA5A5" stop-opacity="0.6"/>
    <stop offset="100%" stop-color="#C0392B" stop-opacity="1"/>
  </linearGradient>`;
  defs += '</defs>';

  let s = defs;

  AGENCIES.forEach((a, i) => {
    const y      = PAD_T + i * (barH + gap);
    const cy     = y + barH / 2;
    const isPos  = a.nbb >= 0;
    const barW   = Math.abs(a.nbb) / maxAbs * BAR_ZONE * 0.92;
    const bgCol  = GROUP_BG[a.group]     || '#F8F8F8';
    const brdCol = GROUP_COLORS[a.group] || '#ccc';

    // Cellule nom — fond couleur groupe
    s += `<rect x="0" y="${y}" width="${NAME_W}" height="${barH}"
      fill="${bgCol}" rx="4"/>`;
    s += `<rect x="0" y="${y}" width="3" height="${barH}"
      fill="${brdCol}" rx="1"/>`;
    s += `<text x="${NAME_W - 6}" y="${cy + 4}"
      text-anchor="end" font-family="IBM Plex Sans" font-size="10" font-weight="600"
      fill="#1E293B">${trunc(a.name, 17)}</text>`;

    // Barre dégradé vert ou rouge
    s += `<rect x="${BAR_START}" y="${y}" width="${Math.max(barW, 2)}" height="${barH}"
      fill="url(#ag_${isPos ? 'pos' : 'neg'})" rx="3"/>`;

    // Valeur
    s += `<text x="${BAR_START + barW + 5}" y="${cy + 4}"
      font-family="IBM Plex Mono" font-size="10" font-weight="700"
      fill="${isPos ? '#1A6B4A' : '#C0392B'}">${isPos ? '+' : ''}${a.nbb.toFixed(1)}m</text>`;
  });

  svg.innerHTML = s;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  svg.style.height = H + 'px';
  svg.style.width  = '100%';
}

function buildAgencyTable() {
  const tbody = document.getElementById('agTableBody');
  if (!tbody) return;
  const THR = 5;
  tbody.innerHTML = AGENCIES.map(a => {
    const bg  = GROUP_BG[a.group]  || '#F8F8F8';
    const brd = GROUP_COLORS[a.group] || '#ccc';
    const wins = a.wrows.filter(r => r[1] >= THR).map(r => `${trunc(r[0],14)} ${fmtVal(r[1])}`).join(' · ') || '—';
    const deps = a.drows.filter(r => r[1] <= -THR).map(r => `${trunc(r[0],14)} ${fmtVal(r[1])}`).join(' · ') || '—';
    return `<tr style="background:${bg};border-left:3px solid ${brd}">
      <td class="td-rank">#${a.rank}</td>
      <td class="td-ag">${a.name}</td>
      <td class="td-mono ${nbbClass(a.nbb)}">${fmtNBB(a.nbb)}</td>
      <td class="td-mono pos">${a.wins > 0 ? fmtVal(a.wins) : '0'}</td>
      <td class="td-mono neg">${a.deps < 0 ? fmtVal(a.deps) : '0'}</td>
      <td style="font-size:11px;color:#4A5568">${wins}</td>
      <td style="font-size:11px;color:#4A5568">${deps}</td>
    </tr>`;
  }).join('');
}

// ══════════════════════════════════════════
// GROUPS LIST
// ══════════════════════════════════════════
function buildGroupsList() {
  const container = document.getElementById('groupsList');
  if (!container) return;
  const agsByGroup = {};
  AGENCIES.forEach(a => {
    if (!agsByGroup[a.group]) agsByGroup[a.group] = [];
    agsByGroup[a.group].push(a);
  });

  container.innerHTML = GROUPS.map(g => {
    const ags = (agsByGroup[g.name] || []).sort((a,b) => b.nbb - a.nbb);
    const agRows = ags.map(a => `
      <div class="agency-sub-row">
        <div style="width:6px;height:6px;border-radius:50%;background:${g.color}"></div>
        <div style="font-size:12px;font-weight:500">${a.name.charAt(0)+a.name.slice(1).toLowerCase()}</div>
        <div class="td-mono ${nbbClass(a.nbb)}" style="font-size:12px">${fmtNBB(a.nbb)}</div>
        <div style="font-family:var(--mono);font-size:11px;color:var(--green)">${a.wins > 0 ? fmtVal(a.wins) : '—'}</div>
        <div style="font-family:var(--mono);font-size:11px;color:var(--red)">${a.deps < 0 ? fmtVal(a.deps) : '—'}</div>
      </div>`).join('');

    return `<div class="group-block">
      <div class="group-block-header" onclick="this.nextElementSibling.classList.toggle('open')">
        <div class="group-color-bar" style="background:${g.color}"></div>
        <div class="group-name">#${g.rank} ${g.name}</div>
        <div class="group-nbb ${nbbClass(g.nbb)}">${fmtNBB(g.nbb)}</div>
        <div class="group-stat">${g.wins.toFixed(0)}<small>Wins $m</small></div>
        <div class="group-stat">${g.deps.toFixed(0)}<small>Dep. $m</small></div>
        <div class="group-stat">${g.wc}<small>Wins nb</small></div>
        <div class="group-stat">${g.dc}<small>Dep. nb</small></div>
      </div>
      <div class="group-agencies open">${agRows}</div>
    </div>`;
  }).join('');
}

// ══════════════════════════════════════════
// RETENTIONS
// ══════════════════════════════════════════
function buildRetentions() {
  const maxRet = Math.max(...RET_DATA.map(r => r.balance));

  // Bars
  // Construire le SVG retention chart
  const barsEl = document.getElementById('retBars');
  if (barsEl) {
    const RET   = RET_DATA.slice(0, 8);
    const W     = 500;
    const barH  = 28;
    const gap   = 10;
    const NAME_W  = 140;
    const BAR_START = NAME_W + 8;
    const BAR_ZONE  = W - BAR_START - 80;
    const PAD_T = 10;
    const H     = PAD_T + RET.length * (barH + gap) + 10;

    let defs = '<defs>';
    // Vert dégradé pour les rétentions (toujours positif)
    defs += `<linearGradient id="ret_pos" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="#86EFAC" stop-opacity="0.55"/>
      <stop offset="100%" stop-color="#1A6B4A" stop-opacity="1"/>
    </linearGradient>`;
    // Dégradés fond cellule nom par groupe
    [...new Set(RET.map(r => r.group))].forEach(grp => {
      const bg  = GROUP_BG[grp]     || '#F8F8F8';
      const brd = GROUP_COLORS[grp] || '#94A3B8';
      defs += `<linearGradient id="rn_${grp.replace(/\\W+/g,'_')}" x1="0%" y1="0%" x2="100%" y2="0%">
        <stop offset="0%" stop-color="${bg}" stop-opacity="1"/>
        <stop offset="100%" stop-color="${bg}" stop-opacity="0.6"/>
      </linearGradient>`;
    });
    defs += '</defs>';

    let s = defs;
    const maxVal = Math.max(...RET.map(r => r.balance));

    RET.forEach((r, i) => {
      const y      = PAD_T + i * (barH + gap);
      const cy     = y + barH / 2;
      const barW   = r.balance / maxVal * BAR_ZONE * 0.92;
      const bgCol  = GROUP_BG[r.group]     || '#F8F8F8';
      const brdCol = GROUP_COLORS[r.group] || '#94A3B8';
      const gradId = `rn_${r.group.replace(/\\W+/g,'_')}`;
      const agName = r.agency.charAt(0) + r.agency.slice(1).toLowerCase();

      // Fond alterné
      s += `<rect x="0" y="${y-3}" width="${W}" height="${barH+6}"
        fill="${i%2===0?'rgba(247,249,252,0.8)':'transparent'}" rx="4"/>`;

      // Cellule nom — fond couleur groupe
      s += `<rect x="0" y="${y}" width="${NAME_W}" height="${barH}"
        fill="url(#${gradId})" rx="5"/>`;
      s += `<rect x="0" y="${y}" width="3" height="${barH}"
        fill="${brdCol}" rx="1"/>`;
      s += `<text x="${NAME_W - 6}" y="${cy + 4}"
        text-anchor="end" font-family="IBM Plex Sans" font-size="11" font-weight="600"
        fill="#1E293B">${agName}</text>`;

      // Barre verte dégradée
      s += `<rect x="${BAR_START}" y="${y+3}" width="${Math.max(barW,3)}" height="${barH-6}"
        fill="url(#ret_pos)" rx="4"/>`;

      // Valeur
      s += `<text x="${BAR_START + barW + 8}" y="${cy + 4}"
        font-family="IBM Plex Mono" font-size="12" font-weight="700"
        fill="#1A6B4A">${fmtNBB(r.balance)}</text>`;
    });

    barsEl.innerHTML = `<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:${H}px;display:block;padding:8px 0">${s}</svg>`;
  }
}

// ══════════════════════════════════════════
// NAV SCROLL SPY
// ══════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
  const sections = document.querySelectorAll('[id^="section-"]');
  const navLinks = document.querySelectorAll('.mast-nav a');

  function updateActiveNav() {
    let current = '';
    sections.forEach(section => {
      if (window.pageYOffset >= section.offsetTop - 150) {
        current = section.getAttribute('id');
      }
    });
    navLinks.forEach(link => {
      link.classList.remove('active');
      if (current && link.getAttribute('href') === '#' + current) {
        link.classList.add('active');
      }
    });
  }

  // Smooth scroll sur clic nav
  navLinks.forEach(link => {
    link.addEventListener('click', e => {
      e.preventDefault();
      const id = link.getAttribute('href').replace('#', '');
      const el = document.getElementById(id);
      if (el) window.scrollTo({ top: el.offsetTop - 56, behavior: 'smooth' });
    });
  });

  // Accordeon groupes
  document.querySelectorAll('.group-block-header').forEach(header => {
    header.addEventListener('click', () => {
      header.nextElementSibling?.classList.toggle('open');
    });
  });

  window.addEventListener('scroll', updateActiveNav, { passive: true });
  updateActiveNav();
});

// ══════════════════════════════════════════
// ANIMATIONS au scroll
// ══════════════════════════════════════════

// Ret bar animation
const retObs = new IntersectionObserver(entries => {
  if (entries[0].isIntersecting) {
    document.querySelectorAll('.ret-bar-fill').forEach(el => {
      const w = el.style.width;
      el.style.width = '0';
      requestAnimationFrame(() => requestAnimationFrame(() => { el.style.width = w; }));
    });
    retObs.disconnect();
  }
}, { threshold: 0.3 });
const retSection = document.getElementById('s4');
if (retSection) retObs.observe(retSection);

// Cards fade-in au scroll
const cardObs = new IntersectionObserver((entries) => {
  entries.forEach((e, i) => {
    if (e.isIntersecting) {
      e.target.style.animationDelay = (i * 0.04) + 's';
      e.target.classList.add('card-visible');
      cardObs.unobserve(e.target);
    }
  });
}, { threshold: 0.1, rootMargin: '0px 0px -40px 0px' });

document.querySelectorAll('.ag-card, .group-block, .tbl-wrap, .chart-box, .kpi-card, .takeaway-card')
  .forEach(el => {
    el.classList.add('card-hidden');
    cardObs.observe(el);
  });

// ══════════════════════════════════════════
// RAIL — TOP 5 agencies dynamique
// ══════════════════════════════════════════
function buildRailAgencies() {
  const el = document.getElementById('railAgencies');
  if (!el) return;
  const top5 = AGENCIES.slice(0, 5);
  el.innerHTML = top5.map(a => `
    <div class="rail-agency"
         onclick="window.scrollTo({top:document.getElementById('s5').offsetTop-56,behavior:'smooth'})">
      <div class="rail-rank">${a.rank}</div>
      <div class="group-dot" style="background:${GROUP_COLORS[a.group]||'#94A3B8'}"></div>
      <div class="rail-ag-name">${a.name.charAt(0)+a.name.slice(1).toLowerCase()}</div>
      <div class="rail-ag-nbb ${a.nbb>=0?'pos':'neg'}">${a.nbb>=0?'+':''}${a.nbb.toFixed(1)}m</div>
    </div>`).join('');
}

// ══════════════════════════════════════════
// INIT
// ══════════════════════════════════════════
drawGroupChart();
drawAgencyChart();
buildAgencyTable();
buildGroupsList();
buildRetentions();
buildRailAgencies();




// ══════════════════════════════════════════════════════════════
// NBB RICH EDITOR v2 — data-edit + tableau + localStorage
// ══════════════════════════════════════════════════════════════
(function() {
  let editMode=false,currentEl=null,activeCell=null,ctxMenu=null,ctxRow=null,resizing=null;

  document.body.insertAdjacentHTML('beforeend', `
    <button id="editBtn">✏️ Éditer</button>
    <div class="save-flash" id="saveFlash">✓ Sauvegardé</div>
    <div id="editorPanel">
      <div class="ep-title">Éditeur <button class="ep-close" id="epClose">✕</button></div>
      <div class="ep-section" id="epText" style="display:none">
        <div class="ep-label">Contenu</div>
        <div class="fmt-btns">
          <button class="fmt-btn" onclick="document.execCommand('bold')"><b>B</b></button>
          <button class="fmt-btn" onclick="document.execCommand('italic')"><i>I</i></button>
          <button class="fmt-btn" onclick="document.execCommand('underline')"><u>U</u></button>
        </div>
        <textarea id="editInput" placeholder="Modifier..."></textarea>
        <button id="saveEdit">Appliquer</button>
      </div>
      <div class="ep-section" id="epCell" style="display:none">
        <div class="ep-label">Fond de cellule</div>
        <div class="cell-swatches">
          <div class="cs-dot" style="background:#fff;border:1px solid #ccc" onclick="applyCellBg('#fff')" title="Blanc"></div>
          <div class="cs-dot" style="background:#E8DAEF" onclick="applyCellBg('#E8DAEF')" title="Publicis"></div>
          <div class="cs-dot" style="background:#D5E8D4" onclick="applyCellBg('#D5E8D4')" title="Omnicom"></div>
          <div class="cs-dot" style="background:#D0E8F2" onclick="applyCellBg('#D0E8F2')" title="Dentsu"></div>
          <div class="cs-dot" style="background:#F9E4C8" onclick="applyCellBg('#F9E4C8')" title="Havas"></div>
          <div class="cs-dot" style="background:#F5CBA7" onclick="applyCellBg('#F5CBA7')" title="WPP"></div>
          <div class="cs-dot" style="background:#D1FAE5" onclick="applyCellBg('#D1FAE5')" title="Vert"></div>
          <div class="cs-dot" style="background:#FFE4E6" onclick="applyCellBg('#FFE4E6')" title="Rouge"></div>
          <div class="cs-dot" style="background:#FEF9C3" onclick="applyCellBg('#FEF9C3')" title="Jaune"></div>
        </div>
        <input type="color" id="cellBgPicker" style="width:100%;height:26px;border:1px solid #E2E8F0;border-radius:6px;cursor:pointer" oninput="applyCellBg(this.value)">
      </div>
      <div class="ep-section" id="epTable" style="display:none">
        <div class="ep-label">Tableau</div>
        <div class="tbl-actions">
          <button class="tbl-btn" onclick="nbEd.addAfter()">+ Après</button>
          <button class="tbl-btn" onclick="nbEd.addBefore()">↑ Avant</button>
          <button class="tbl-btn" onclick="nbEd.dup()">Dupliquer</button>
          <button class="tbl-btn danger" onclick="nbEd.del()">🗑 Suppr.</button>
        </div>
        <div style="font-size:10px;color:#94A3B8;margin-top:5px">Drag bord. colonne → redim.</div>
      </div>
      <div class="ep-section">
        <div class="ep-label">Couleurs globales</div>
        <div class="color-row"><label>Vert</label><input type="color" id="cpGreen" value="#1A6B4A" oninput="setVar('--green',this.value)"></div>
        <div class="color-row"><label>Or</label><input type="color" id="cpGold" value="#C9A84C" oninput="setVar('--gold',this.value)"></div>
      </div>
      <div class="ep-section">
        <div class="ep-label">Exporter</div>
        <div class="export-btns">
          <button class="exp-btn exp-html" onclick="doExportHTML()">⬇ HTML</button>
          <button class="exp-btn exp-pdf"  onclick="doExportPDF()">🖨 PDF/A4</button>
        </div>
      </div>
    </div>`);

  const btn=document.getElementById('editBtn'),panel=document.getElementById('editorPanel');

  btn.addEventListener('click',()=>{
    editMode=!editMode;
    btn.textContent=editMode?'✅ Terminer':'✏️ Éditer';
    btn.classList.toggle('active',editMode);
    document.body.classList.toggle('edit-mode',editMode);
    panel.classList.toggle('visible',editMode);
    if(!editMode){clearSel();clearCell();hideSections();}
  });
  document.getElementById('epClose').addEventListener('click',()=>panel.classList.remove('visible'));

  document.addEventListener('click',e=>{
    if(!editMode)return;
    removeCtx();
    const el=e.target.closest('[data-edit]'),cell=e.target.closest('td,th');
    if(el){
      e.stopPropagation();clearSel();clearCell();
      currentEl=el;el.classList.add('selected');
      document.getElementById('epText').style.display='block';
      document.getElementById('editInput').value=el.innerText.trim();
      panel.classList.add('visible');
    } else if(cell){
      clearSel();clearCell();activeCell=cell;cell.classList.add('cell-selected');
      document.getElementById('epCell').style.display='block';
      document.getElementById('epTable').style.display='block';
      document.getElementById('cellBgPicker').value=rgb2hex(getComputedStyle(cell).backgroundColor);
      panel.classList.add('visible');
    } else {clearSel();clearCell();}
  });

  document.addEventListener('dblclick',e=>{
    if(!editMode)return;
    const t=e.target.closest('td,th')||e.target.closest('[data-edit]');
    if(!t)return;
    t.contentEditable='true';t.focus();
    t.addEventListener('blur',()=>{t.contentEditable='false';save(t);},{once:true});
    t.addEventListener('keydown',ev=>{
      if(ev.key==='Enter'&&!ev.shiftKey&&t.tagName!=='P'){ev.preventDefault();t.blur();}
      if(ev.key==='Escape')t.blur();
    });
    e.stopPropagation();
  });

  document.getElementById('saveEdit').addEventListener('click',()=>{
    if(!currentEl)return;
    currentEl.innerText=document.getElementById('editInput').value;
    save(currentEl);flash();
  });

  function save(el){const k=el.dataset?.edit||el.dataset?.field;if(k){localStorage.setItem('nbb_'+k,el.innerHTML);flash();}}
  function loadAll(){
    document.querySelectorAll('[data-edit],[data-field]').forEach(el=>{
      const k=el.dataset?.edit||el.dataset?.field;
      if(k){const v=localStorage.getItem('nbb_'+k);if(v)el.innerHTML=v;}
    });
    const g=localStorage.getItem('nbb_css_green'),gl=localStorage.getItem('nbb_css_gold');
    if(g){document.documentElement.style.setProperty('--green',g);document.getElementById('cpGreen').value=g;}
    if(gl){document.documentElement.style.setProperty('--gold',gl);document.getElementById('cpGold').value=gl;}
  }

  window.applyCellBg=c=>{if(activeCell){activeCell.style.backgroundColor=c;flash();}};
  window.setVar=(v,c)=>{document.documentElement.style.setProperty(v,c);localStorage.setItem('nbb_css_'+v.replace('--',''),c);};

  document.addEventListener('contextmenu',e=>{
    if(!editMode)return;const row=e.target.closest('tr');
    if(!row)return;e.preventDefault();ctxRow=row;showCtx(e.clientX,e.clientY);
  });
  function showCtx(x,y){
    removeCtx();ctxMenu=document.createElement('div');ctxMenu.id='ctxMenu';
    ctxMenu.innerHTML=`
      <div class="ctx-item" onclick="nbEd.addAfter()">+ Après</div>
      <div class="ctx-item" onclick="nbEd.addBefore()">↑ Avant</div>
      <div class="ctx-item" onclick="nbEd.dup()">Dupliquer</div>
      <div class="ctx-sep"></div>
      <div class="ctx-item ctx-danger" onclick="nbEd.del()">🗑 Supprimer</div>`;
    ctxMenu.style.cssText=`left:${x}px;top:${y}px`;
    document.body.appendChild(ctxMenu);
    setTimeout(()=>document.addEventListener('click',removeCtx,{once:true}),10);
  }
  function removeCtx(){if(ctxMenu){ctxMenu.remove();ctxMenu=null;}}
  function mkRow(keep=false){
    const r=ctxRow.cloneNode(true);
    if(!keep)r.querySelectorAll('td,th').forEach(c=>{c.innerHTML='—';c.style.background='';});
    return r;
  }
  window.nbEd={
    addAfter(){if(ctxRow)ctxRow.parentNode.insertBefore(mkRow(),ctxRow.nextSibling);removeCtx();},
    addBefore(){if(ctxRow)ctxRow.parentNode.insertBefore(mkRow(),ctxRow);removeCtx();},
    dup(){if(ctxRow)ctxRow.parentNode.insertBefore(mkRow(true),ctxRow.nextSibling);removeCtx();},
    del(){
      if(!ctxRow)return;
      const tb=ctxRow.closest('tbody');
      if(tb&&tb.rows.length<=1){alert('Dernière ligne.');return;}
      ctxRow.remove();removeCtx();
    }
  };

  document.addEventListener('mousemove',e=>{
    if(!editMode||resizing)return;
    const th=e.target.closest('th');
    document.body.style.cursor=(th&&e.clientX>th.getBoundingClientRect().right-6)?'col-resize':'';
  });
  document.addEventListener('mousedown',e=>{
    if(!editMode)return;const th=e.target.closest('th');
    if(!th||e.clientX<=th.getBoundingClientRect().right-6)return;
    e.preventDefault();resizing={th,startX:e.clientX,startW:th.offsetWidth};
    document.body.style.userSelect='none';
  });
  document.addEventListener('mousemove',e=>{
    if(!resizing)return;
    const w=Math.max(40,resizing.startW+(e.clientX-resizing.startX));
    resizing.th.style.width=resizing.th.style.minWidth=w+'px';
  });
  document.addEventListener('mouseup',()=>{
    if(!resizing)return;resizing=null;
    document.body.style.cursor=document.body.style.userSelect='';
  });

  window.doExportHTML=()=>{
    const cl=document.documentElement.cloneNode(true);
    ['#editBtn','#editorPanel','#ctxMenu','.save-flash'].forEach(s=>cl.querySelector(s)?.remove());
    cl.querySelectorAll('[contenteditable]').forEach(el=>el.removeAttribute('contenteditable'));
    cl.querySelectorAll('.selected,.cell-selected').forEach(el=>el.classList.remove('selected','cell-selected'));
    cl.body.classList.remove('edit-mode');
    const b=new Blob(['<!DOCTYPE html>'+cl.outerHTML],{type:'text/html'});
    Object.assign(document.createElement('a'),{href:URL.createObjectURL(b),download:'NBB_Report_edited.html'}).click();
  };
  window.doExportPDF=()=>{
    btn.style.display=panel.style.display='none';
    document.body.classList.remove('edit-mode');
    window.print();
    setTimeout(()=>{btn.style.display='';panel.style.display=editMode?'block':'none';document.body.classList.toggle('edit-mode',editMode);},500);
  };

  function clearSel(){if(currentEl){currentEl.classList.remove('selected');currentEl=null;}}
  function clearCell(){if(activeCell){activeCell.classList.remove('cell-selected');activeCell=null;}}
  function hideSections(){['epText','epCell','epTable'].forEach(id=>document.getElementById(id).style.display='none');}
  function flash(){const f=document.getElementById('saveFlash');f.classList.add('show');setTimeout(()=>f.classList.remove('show'),1400);}
  function rgb2hex(rgb){if(!rgb||rgb==='transparent')return'#ffffff';const m=rgb.match(/\\d+/g);if(!m)return'#ffffff';return'#'+m.slice(0,3).map(x=>parseInt(x).toString(16).padStart(2,'0')).join('');}

  loadAll();
})();

"""

PREMIUM_TOOLBAR = ""



def _build_data_js(data) -> str:
    import json as _json
    GROUP_COLORS_MAP = {
        'Publicis Media':      '#9B59B6', 'Omnicom Media':       '#27AE60',
        'Dentsu':              '#2980B9', 'Havas Media Network': '#E67E22',
        'WPP Media':           '#E74C3C', 'Independant':         '#95A5A6',
    }
    GROUP_BG_MAP = {
        'Publicis Media':      '#E8DAEF', 'Omnicom Media':       '#D5E8D4',
        'Dentsu':              '#D0E8F2', 'Havas Media Network': '#F9E4C8',
        'WPP Media':           '#F5CBA7', 'Independant':         '#F0F0F0',
    }
    agencies_js = []
    for a in data['agencies']:
        agencies_js.append({
            'rank': a['rank'], 'name': a['agency'], 'group': a['group'],
            'nbb':  round(a['nbb'],1),  'wins': round(a['wins'],1),
            'deps': round(a['deps'],1), 'rets': round(a['rets'],1),
            'wrows': [[str(r.get('Advertiser','')), float(r.get('Integrated Spends',0))] for r in a['wins_rows'][:6]],
            'drows': [[str(r.get('Advertiser','')), float(r.get('Integrated Spends',0))] for r in a['dep_rows'][:6]],
            'rrows': [[str(r.get('Advertiser','')), float(r.get('Integrated Spends',0))] for r in a['ret_rows'][:6]],
        })
    groups_sorted = sorted([gs for gs in data['group_stats'].values() if gs['agencies']], key=lambda g: -g['nbb'])
    groups_js = []
    for gs in groups_sorted:
        groups_js.append({
            'rank': gs['rank'], 'name': gs['name'],
            'color': GROUP_COLORS_MAP.get(gs['name'], '#94A3B8'),
            'bg':    GROUP_BG_MAP.get(gs['name'], '#F0F0F0'),
            'nbb':   round(gs['nbb'],1), 'wins': round(gs['wins'],1),
            'deps':  round(gs['deps'],1), 'wc': gs['wc'], 'dc': gs['dc'],
        })
    market = data.get('market', 'Market')
    period = data.get('period', '2025')
    return (
        f"const AGENCIES = {_json.dumps(agencies_js, ensure_ascii=False)};\n"
        f"const GROUPS = {_json.dumps(groups_js, ensure_ascii=False)};\n"
        "const RET_DATA = AGENCIES.filter(a=>a.rets>0)"
        ".map(a=>({agency:a.name,group:a.group,balance:a.rets,topClient:a.rrows[0]?.[0]||'—'}))"
        ".sort((a,b)=>b.balance-a.balance);\n"
        f"const GROUP_COLORS = {_json.dumps(GROUP_COLORS_MAP, ensure_ascii=False)};\n"
        f"const GROUP_BG     = {_json.dumps(GROUP_BG_MAP, ensure_ascii=False)};\n"
        f"const MARKET = {_json.dumps(market)};\n"
        f"const PERIOD = {_json.dumps(period)};\n"
    )


def build_report_html(df: pd.DataFrame, threshold: float = 5.0) -> bytes:
    """Génère le rapport HTML Premium dynamique depuis un DataFrame Excel NBB."""
    data    = get_data(df)
    market  = data.get('market', 'Market')
    period  = data.get('period', 'NBB 2025')
    data_js = _build_data_js(data)
    thr     = float(threshold)

    agencies  = data['agencies']
    total_nbb = sum(a['nbb']  for a in agencies)
    total_w   = sum(a['wins'] for a in agencies)
    total_d   = sum(a['deps'] for a in agencies)
    total_r   = sum(a['rets'] for a in agencies)
    n_w = sum(len(a['wins_rows']) for a in agencies)
    n_d = sum(len(a['dep_rows'])  for a in agencies)
    n_r = sum(len(a['ret_rows'])  for a in agencies)
    nbb_cls = 'pos' if total_nbb >= 0 else 'neg'

    sections = '\n'.join([
        build_cover(data),
        build_top_moves(data),
        build_agencies_overview(data, threshold=thr),
        build_groups_overview(data),
        build_retentions(data),
        build_agency_details(data),
    ])

    nav_html = f'''<header class="masthead">
  <div class="masthead-inner">
    <div class="mast-brand">
      <div class="mast-recma">RECMA</div>
      <div class="mast-title">New Business Balance</div>
      <div class="mast-sub">{market} · {period}</div>
    </div>
    <nav class="mast-nav">
      <a href="#section-0" class="active">01 Key Findings</a>
      <a href="#section-1">02 TOP moves</a>
      <a href="#section-2">03 Agencies</a>
      <a href="#section-3">04 Groups</a>
      <a href="#section-4">05 Retentions</a>
      <a href="#section-5">06 Details</a>
    </nav>
    <div class="mast-badge">
      <span class="mast-market">{market.upper()}</span>
      <span class="mast-date">{period}</span>
    </div>
  </div>
</header>'''

    rail_html = f'''<aside class="left-rail">
  <div>
    <div class="rail-section-title">Market KPIs</div>
    <div class="kpi-card">
      <div class="kpi-label">Total NBB</div>
      <div class="kpi-value {nbb_cls}">{fmt(total_nbb)}</div>
      <div class="kpi-sub">{len(agencies)} agencies · {len(data["group_stats"])} groups</div>
      <div class="mini-bar"><div class="mini-bar-fill {nbb_cls}" style="width:70%"></div></div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Wins</div>
      <div class="kpi-value pos" style="font-size:18px">{fmt(total_w)}</div>
      <div class="kpi-sub">{n_w} moves WIN</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Departures</div>
      <div class="kpi-value neg" style="font-size:18px">{fmt(total_d)}</div>
      <div class="kpi-sub">{n_d} moves DEPARTURE</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Retentions</div>
      <div class="kpi-value" style="color:var(--gold);font-size:18px">{fmt(total_r)}</div>
      <div class="kpi-sub">{n_r} moves</div>
    </div>
  </div>
  <div>
    <div class="rail-section-title">Top agencies</div>
    <div id="railAgencies"></div>
    <a href="#section-2" class="rail-more"
       onclick="event.preventDefault();window.scrollTo({{top:document.getElementById('section-2').offsetTop-56,behavior:'smooth'}})">
      Voir toutes →
    </a>
  </div>
</aside>'''

    # Concaténation directe pour éviter le double-échappement du JS statique
    head = (
        "<!DOCTYPE html>\n<html lang=\"fr\">\n<head>\n"
        "<meta charset=\"UTF-8\">\n"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n"
        f"<title>NBB Report {period} · {market}</title>\n"
        "<link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">\n"
        "<link href=\"https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700;900"
        "&family=IBM+Plex+Sans:wght@300;400;500;600"
        "&family=IBM+Plex+Mono:wght@400;500;600&display=swap\" rel=\"stylesheet\">\n"
        "<style>"
    )
    html = (
        head
        + PREMIUM_CSS.replace("{{", "{").replace("}}", "}")
        + "</style>\n</head>\n<body>\n"
        + nav_html + "\n"
        + "<div class=\"layout\">\n"
        + rail_html + "\n"
        + "<main class=\"main\">\n"
        + sections + "\n"
        + "</main>\n</div>\n"
        + "<script>\n"
        + data_js
        + PREMIUM_JS_STATIC
        + "\n</script>\n"
        + PREMIUM_TOOLBAR + "\n"
        + "</body>\n</html>"
    )
    return html.encode('utf-8')




# ─────────────────────────────────────────────────────────────
# STANDALONE
# ─────────────────────────────────────────────────────────────
if __name__ == '__main__':
    EXCEL  = sys.argv[1] if len(sys.argv) > 1 else '/mnt/user-data/uploads/Newbiz_Balance_DB_Report__1_.xlsx'
    OUTPUT = sys.argv[2] if len(sys.argv) > 2 else '/mnt/user-data/outputs/NBB_Report.html'

    df     = pd.read_excel(EXCEL)
    result = build_report_html(df)

    with open(OUTPUT, 'wb') as f:
        f.write(result)
    print(f'✅ {OUTPUT}  ({len(result)//1024} KB)')
