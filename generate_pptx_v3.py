import io
import copy 
import pandas as pd
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# Configuration Slide 06 (Détails)
SLIDE_INDEX_DETAILS = 5  
MAX_AGENTS_PER_DETAIL_SLIDE = 4 

def format_nbb(val):
    if pd.isna(val) or val == 0: return "0m"
    sign = "+" if val > 0 else ""
    return f"{sign}{val:.1f}m"

def duplicate_slide(prs, source_slide):
    slide_layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)
    for shape in source_slide.shapes:
        # Utilisation d'une méthode plus robuste pour copier les formes
        new_el = copy.deepcopy(shape.element)
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    return new_slide

def replace_tags_in_shape(shape, replacements):
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
                if cell.has_text_frame:
                    replace_tags_in_shape(cell, replacements)

def build_agency_pptx(df_excel, template_path):
    df = df_excel.copy()
    # Nettoyage des colonnes critiques
    df['Agency'] = df['Agency'].astype(str).str.strip()
    df['Integrated Spends'] = pd.to_numeric(df['Integrated Spends'], errors='coerce').fillna(0)
    
    summary = df.groupby('Agency').agg({'Integrated Spends': 'sum'}).reset_index()
    summary = summary.sort_values('Integrated Spends', ascending=False).reset_index(drop=True)

    prs = Presentation(template_path)
    
    # 1. Préparation des balises globales (Top 14)
    replacements = {}
    for i, row in summary.head(14).iterrows():
        idx = i + 1
        ag_name = row['Agency']
        replacements[f"{{{{AG_{idx}}}}}"] = str(ag_name).upper()
        replacements[f"{{{{NBB_{idx}}}}}"] = format_nbb(row['Integrated Spends'])
        
        ag_data = df[df['Agency'] == ag_name]
        wins = ag_data[ag_data['NewBiz'] == 'WIN'].sort_values('Integrated Spends', ascending=False).head(3)
        deps = ag_data[ag_data['NewBiz'] == 'DEPARTURE'].sort_values('Integrated Spends', ascending=True).head(3)
        
        # Correction ici : accès sécurisé aux attributs du tuple (itertuples)
        replacements[f"{{{{TOPWINS_{idx}}}}}"] = " · ".join([f"{getattr(r, 'Advertiser', 'N/A')} {format_nbb(getattr(r, 'Integrated_Spends', 0))}" for r in wins.itertuples()])
        replacements[f"{{{{TOPDEPS_{idx}}}}}"] = " · ".join([f"{getattr(r, 'Advertiser', 'N/A')} {format_nbb(getattr(r, 'Integrated_Spends', 0))}" for r in deps.itertuples()])

    # 2. Duplication et remplissage de la Slide 06 (Détails)
    total_agencies = len(summary)
    pages_needed = (total_agencies // MAX_AGENTS_PER_DETAIL_SLIDE) + (1 if total_agencies % MAX_AGENTS_PER_DETAIL_SLIDE > 0 else 0)
    
    if len(prs.slides) > SLIDE_INDEX_DETAILS:
        detail_template_slide = prs.slides[SLIDE_INDEX_DETAILS]
        
        for p in range(pages_needed):
            current_slide = detail_template_slide if p == 0 else duplicate_slide(prs, detail_template_slide)
            start_idx = p * MAX_AGENTS_PER_DETAIL_SLIDE
            page_replacements = {}
            
            for n in range(MAX_AGENTS_PER_DETAIL_SLIDE):
                ag_idx = start_idx + n
                tag_num = n + 1
                if ag_idx < total_agencies:
                    row_ag = summary.iloc[ag_idx]
                    page_replacements[f"{{{{D_AG_{tag_num}}}}}"] = str(row_ag['Agency']).upper()
                    page_replacements[f"{{{{D_NBB_{tag_num}}}}}"] = format_nbb(row_ag['Integrated Spends'])
                else:
                    page_replacements[f"{{{{D_AG_{tag_num}}}}}"] = ""
                    page_replacements[f"{{{{D_NBB_{tag_num}}}}}"] = ""

            for shape in current_slide.shapes:
                replace_tags_in_shape(shape, {**replacements, **page_replacements})

    # 3. Remplissage des slides statiques (0-4)
    for i, slide in enumerate(prs.slides):
        if i >= SLIDE_INDEX_DETAILS: continue 
        for shape in slide.shapes:
            replace_tags_in_shape(shape, replacements)

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()
