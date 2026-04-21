import io
import copy
import pandas as pd
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# --- CONFIGURATION DU TEMPLATE ---
# Index de la slide modèle pour les détails (Slide 06 -> Index 5)
SLIDE_INDEX_DETAILS = 5  
# Nombre d'agences par slide dans votre design "Glass"
MAX_AGENTS_PER_DETAIL_SLIDE = 4 

def format_nbb(val):
    """Formate les montants avec signe, 1 décimale et suffixe 'm'."""
    if pd.isna(val) or val == 0:
        return "0m"
    sign = "+" if val > 0 else ""
    return f"{sign}{val:.1f}m"

def duplicate_slide(prs, source_slide):
    """Duplique une slide en conservant exactement le layout et les éléments graphiques."""
    slide_layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)
    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape.element)
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    return new_slide

def replace_tags_in_shape(shape, replacements):
    """Remplace les balises dans les formes simples et les groupes d'objets."""
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for s in shape.shapes:
            replace_tags_in_shape(s, replacements)
    elif shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                for tag, value in replacements.items():
                    if tag in run.text:
                        run.text = run.text.replace(tag, str(value))
    elif shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                replace_tags_in_shape(cell, replacements)

def build_agency_pptx(df_excel, template_path, prefilled_prs_bytes=None):
    """Moteur principal de génération PPTX."""
    # 1. PRÉPARATION DES DONNÉES
    df = df_excel.copy()
    df['Integrated Spends'] = pd.to_numeric(df['Integrated Spends'], errors='coerce').fillna(0)
    
    # Calcul du classement NBB par agence
    summary = df.groupby('Agency').agg({'Integrated Spends': 'sum'}).reset_index()
    summary = summary.sort_values('Integrated Spends', ascending=False).reset_index(drop=True)

    # 2. CHARGEMENT DU TEMPLATE
    if prefilled_prs_bytes:
        prs = Presentation(io.BytesIO(prefilled_prs_bytes))
    else:
        prs = Presentation(template_path)
    
    # 3. GÉNÉRATION DU CLASSEMENT (Slides 02 à 05)
    replacements = {}
    for i, row in summary.head(14).iterrows():
        idx = i + 1
        ag_name = row['Agency']
        replacements[f"{{{{AG_{idx}}}}}"] = ag_name.upper()
        replacements[f"{{{{NBB_{idx}}}}}"] = format_nbb(row['Integrated Spends'])
        replacements[f"{{{{RANK_{idx}}}}}"] = str(idx)

        # Extraction des Top Moves pour la Slide 04
        ag_data = df[df['Agency'] == ag_name]
        wins = ag_data[ag_data['NewBiz'] == 'WIN'].sort_values('Integrated Spends', ascending=False).head(3)
        deps = ag_data[ag_data['NewBiz'] == 'DEPARTURE'].sort_values('Integrated Spends', ascending=True).head(3)
        rets = ag_data[ag_data['NewBiz'] == 'RETENTION'].head(3)

        replacements[f"{{{{TOPWINS_{idx}}}}}"] = " · ".join([f"{r.Advertiser} {format_nbb(r['Integrated Spends'])}" for r in wins.itertuples()])
        replacements[f"{{{{TOPDEPS_{idx}}}}}"] = " · ".join([f"{r.Advertiser} {format_nbb(r['Integrated Spends'])}" for r in deps.itertuples()])
        replacements[f"{{{{TOPRET_{idx}}}}}"] = " · ".join([str(r.Advertiser) for r in rets.itertuples()])

    # 4. GÉNÉRATION DYNAMIQUE DES DÉTAILS (Slide 06)
    total_agencies = len(summary)
    pages_needed = (total_agencies // MAX_AGENTS_PER_DETAIL_SLIDE) + (1 if total_agencies % MAX_AGENTS_PER_DETAIL_SLIDE > 0 else 0)
    
    # On isole la slide 6 d'origine
    detail_template_slide = prs.slides[SLIDE_INDEX_DETAILS]
    
    for p in range(pages_needed):
        # Utilise la slide existante pour la p0, sinon duplique
        current_slide = detail_template_slide if p == 0 else duplicate_slide(prs, detail_template_slide)
        
        start_idx = p * MAX_AGENTS_PER_DETAIL_SLIDE
        page_replacements = {}
        
        for n in range(MAX_AGENTS_PER_DETAIL_SLIDE):
            ag_idx = start_idx + n
            tag_num = n + 1
            
            if ag_idx < total_agencies:
                row = summary.iloc[ag_idx]
                page_replacements[f"{{{{D_AG_{tag_num}}}}}"] = row['Agency'].upper()
                page_replacements[f"{{{{D_NBB_{tag_num}}}}}"] = format_nbb(row['Integrated Spends'])
            else:
                # Nettoyage si moins de 4 agences sur la slide
                page_replacements[f"{{{{D_AG_{tag_num}}}}}"] = ""
                page_replacements[f"{{{{D_NBB_{tag_num}}}}}"] = ""

        # Application des remplacements sur la slide de détails actuelle
        for shape in current_slide.shapes:
            replace_tags_in_shape(shape, {**replacements, **page_replacements})

    # 5. REMPLACEMENT FINAL SUR TOUTES LES AUTRES SLIDES (01-05)
    for i, slide in enumerate(prs.slides):
        # On saute les slides de détails car déjà traitées
        if i >= SLIDE_INDEX_DETAILS: continue 
        for shape in slide.shapes:
            replace_tags_in_shape(shape, replacements)

    # 6. EXPORT
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output.getvalue()
