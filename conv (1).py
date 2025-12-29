import streamlit as st
import os
import shutil
import pandas as pd
import xlwt # For writing XLS-compatible files
import numpy as np
import csv
import io # For handling uploaded file bytes
from datetime import datetime # Needed for date formatting

# --- Configuration & Setup: Hardcoded paths ---
BASE_DIR = "C:/MADINA" # !!! IMPORTANT: This directory MUST exist and be writable !!!
STATIC_DIR = os.path.join(BASE_DIR, "static")
DATA_DIR = os.path.join(STATIC_DIR, "DATA")
CONF_DIR = os.path.join(STATIC_DIR, "CONF")
GLOBALE_DIR = os.path.join(STATIC_DIR, "GLOBALE")
IMAGES_DIR = os.path.join(BASE_DIR, "images") # Path for product images
NUMBER_FILE_PATH = os.path.join(STATIC_DIR, "number.txt")
NUMBER2_FILE_PATH = os.path.join(STATIC_DIR, "number2.txt")
LOGO_PATH = os.path.join(STATIC_DIR, "logo_Madina.png")

DEFAULT_IMAGE_BASE_URL = "https://www.madina.dz/wp-content/uploads/Images_Madinadz/Brand/Breakout/"

def ensure_app_filesystem():
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(CONF_DIR, exist_ok=True)
    os.makedirs(GLOBALE_DIR, exist_ok=True)
    os.makedirs(IMAGES_DIR, exist_ok=True)
    if not os.path.exists(NUMBER_FILE_PATH):
        with open(NUMBER_FILE_PATH, "w") as f: f.write("0")
    if not os.path.exists(NUMBER2_FILE_PATH):
        with open(NUMBER2_FILE_PATH, "w") as f: f.write("0")
ensure_app_filesystem()

# --- Streamlit UI Setup ---
st.set_page_config(page_title="Madina Convertisseur", layout="wide")
st.title("Madina Convertisseur")
st.markdown('[By Abdellah Zouggari](https://www.linkedin.com/in/abdellah-zouggari-59195a245)')
if os.path.exists(LOGO_PATH): st.image(LOGO_PATH, width=150)
else: st.warning(f"Logo non trouvé à : {LOGO_PATH}")

# Initialize session state
if 'data_file_count' not in st.session_state:
    try: st.session_state.data_file_count = int(open(NUMBER_FILE_PATH, "r").read().strip())
    except: st.session_state.data_file_count = 0
if 'conf_file_count' not in st.session_state:
    try: st.session_state.conf_file_count = int(open(NUMBER2_FILE_PATH, "r").read().strip())
    except: st.session_state.conf_file_count = 0
if 'last_generated_csv' not in st.session_state: st.session_state.last_generated_csv = None
if 'processing_log' not in st.session_state: st.session_state.processing_log = []
if 'image_base_url' not in st.session_state: st.session_state.image_base_url = DEFAULT_IMAGE_BASE_URL
if 'preview_df' not in st.session_state: st.session_state.preview_df = None
if 'latest_data_file_name' not in st.session_state: st.session_state.latest_data_file_name = None
if 'latest_conf_file_name' not in st.session_state: st.session_state.latest_conf_file_name = None


# --- Sidebar ---
st.sidebar.header("Configuration Globale")
st.session_state.image_base_url = st.sidebar.text_input(
    "URL de base pour les images:", value=st.session_state.image_base_url
)
st.sidebar.caption(f"Noms d'images trouvés dans `{IMAGES_DIR}` seront ajoutés.")
st.sidebar.caption(f"Ex: `{st.session_state.image_base_url}image.jpg`")

st.sidebar.markdown("---")
st.sidebar.header("1. Chargement des Fichiers")

# Data File Uploader
uploaded_data_file = st.sidebar.file_uploader("Charger le fichier DATA (.xlsx)", type="xlsx", key="data_uploader")
if uploaded_data_file is not None:
    try:
        with open(NUMBER_FILE_PATH, "r") as f_num: num = int(f_num.read().strip()) + 1
        save_path = os.path.join(DATA_DIR, f'data_{num}.xlsx')
        with open(save_path, "wb") as f: f.write(uploaded_data_file.getbuffer())
        with open(NUMBER_FILE_PATH, "w") as f_num: f_num.write(str(num))
        st.session_state.data_file_count = num
        st.session_state.latest_data_file_name = f'data_{num}.xlsx'
        st.sidebar.success(f"Fichier data '{st.session_state.latest_data_file_name}' chargé!")
    except Exception as e: st.sidebar.error(f"Erreur sauvegarde data: {e}")

if st.session_state.latest_data_file_name:
    st.sidebar.caption(f"Fichier data prêt pour traitement: **{st.session_state.latest_data_file_name}**")
st.sidebar.info(f"Total fichiers data sur serveur: {st.session_state.data_file_count}")

# Config File Uploader
uploaded_conf_file = st.sidebar.file_uploader("Charger le fichier CONFIG (.xlsx)", type="xlsx", key="conf_uploader")
if uploaded_conf_file is not None:
    try:
        with open(NUMBER2_FILE_PATH, "r") as f_num: num_conf = int(f_num.read().strip()) + 1
        save_path_conf = os.path.join(CONF_DIR, f'conf_{num_conf}.xlsx')
        with open(save_path_conf, "wb") as f: f.write(uploaded_conf_file.getbuffer())
        with open(NUMBER2_FILE_PATH, "w") as f_num: f_num.write(str(num_conf))
        st.session_state.conf_file_count = num_conf
        st.session_state.latest_conf_file_name = f'conf_{num_conf}.xlsx'
        st.sidebar.success(f"Fichier config '{st.session_state.latest_conf_file_name}' chargé!")
    except Exception as e: st.sidebar.error(f"Erreur sauvegarde config: {e}")

if st.session_state.latest_conf_file_name:
    st.sidebar.caption(f"Fichier config prêt pour traitement: **{st.session_state.latest_conf_file_name}**")
st.sidebar.info(f"Total fichiers config sur serveur: {st.session_state.conf_file_count}")


# --- Core Processing Logic (Unchanged as per request) ---
def process_files_core(data_file_num_str, conf_file_num_str, image_base_url_param):
    log_messages = []
    log_messages.append(f"Début: data_{data_file_num_str}.xlsx, conf_{conf_file_num_str}.xlsx")
    log_messages.append(f"URL base images: {image_base_url_param}")
    if image_base_url_param and not image_base_url_param.endswith('/'):
        image_base_url_param += '/'; log_messages.append(f"URL base ajustée: {image_base_url_param}")

    data_file_path = os.path.join(DATA_DIR, f'data_{data_file_num_str}.xlsx')
    conf_file_path = os.path.join(CONF_DIR, f'conf_{conf_file_num_str}.xlsx')

    if not os.path.exists(data_file_path): log_messages.append(f"ERREUR: Data non trouvé!"); st.session_state.processing_log.extend(log_messages); st.session_state.preview_df=None; return None
    if not os.path.exists(conf_file_path): log_messages.append(f"ERREUR: Config non trouvé!"); st.session_state.processing_log.extend(log_messages); st.session_state.preview_df=None; return None
    try:
        pandas_data = pd.read_excel(data_file_path, engine='openpyxl').fillna('')
        pandas_conf = pd.read_excel(conf_file_path, engine='openpyxl').fillna('')
        if "Date de début de promo" not in pandas_data.columns: pandas_data["Date de début de promo"] = ''
        if "Date de fin de promo" not in pandas_data.columns: pandas_data["Date de fin de promo"] = ''
    except Exception as e: log_messages.append(f"ERREUR lecture Excel: {e}"); st.session_state.processing_log.extend(log_messages); st.session_state.preview_df=None; return None

    data_list_full = []
    data_var_aggregated = {}
    try:
        list_images_fs = os.listdir(IMAGES_DIR); log_messages.append(f"{len(list_images_fs)} images trouvées: {IMAGES_DIR}")
    except Exception as e: log_messages.append(f"ERREUR accès images {IMAGES_DIR}: {e}"); list_images_fs = []

    date_format_output = '%Y-%m-%d %H:%M:%S'
    input_date_formats = ['%d/%m/%Y %H:%M:%S', '%d/%m/%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d']

    for i in range(len(pandas_data)):
        lign = {}
        try:
            code_a_barre = str(pandas_data.at[i, "CODE A BARRE"]) if "CODE A BARRE" in pandas_data.columns and pd.notna(pandas_data.at[i, "CODE A BARRE"]) else ''
            code_ref_orig = pandas_data.at[i, "Code * Références"] if "Code * Références" in pandas_data.columns and pd.notna(pandas_data.at[i, "Code * Références"]) else ''
            code_ref_processed = str(code_ref_orig).replace('.0', '') if isinstance(code_ref_orig, (float, int)) else str(code_ref_orig)
            designation = str(pandas_data.at[i, "DESIGNATION"]) if "DESIGNATION" in pandas_data.columns and pd.notna(pandas_data.at[i, "DESIGNATION"]) else ''
            age_orig = str(pandas_data.at[i, "Age"]) if "Age" in pandas_data.columns and pd.notna(pandas_data.at[i, "Age"]) else ''
            taille_orig = str(pandas_data.at[i, "Taille"]) if "Taille" in pandas_data.columns and pd.notna(pandas_data.at[i, "Taille"]) else ''

            date_debut_promo_str, date_fin_promo_str = "", ""
            try:
                raw_start_date = pandas_data.at[i, "Date de début de promo"]
                if pd.notna(raw_start_date) and raw_start_date != '':
                     dt_obj_start = None
                     if isinstance(raw_start_date, datetime): dt_obj_start = raw_start_date
                     else:
                         for fmt in input_date_formats:
                             try: dt_obj_start = datetime.strptime(str(raw_start_date).split(" ")[0], fmt); break
                             except: continue
                     if dt_obj_start: date_debut_promo_str = dt_obj_start.strftime(date_format_output)
            except Exception as date_ex: log_messages.append(f"AVERT: Format date début promo non reconnu ligne {i+2}: {pandas_data.at[i, 'Date de début de promo']} ({date_ex})")
            try:
                raw_end_date = pandas_data.at[i, "Date de fin de promo"]
                if pd.notna(raw_end_date) and raw_end_date != '':
                    dt_obj_end = None
                    if isinstance(raw_end_date, datetime): dt_obj_end = raw_end_date
                    else:
                         for fmt in input_date_formats:
                             try: dt_obj_end = datetime.strptime(str(raw_end_date).split(" ")[0], fmt); break
                             except: continue
                    if dt_obj_end: date_fin_promo_str = dt_obj_end.strftime(date_format_output)
            except Exception as date_ex: log_messages.append(f"AVERT: Format date fin promo non reconnu ligne {i+2}: {pandas_data.at[i, 'Date de fin de promo']} ({date_ex})")

            try: prix_promo = float(pandas_data.at[i, "Prix Promo"]) if "Prix Promo" in pandas_data.columns and pandas_data.at[i, "Prix Promo"] != '' else None
            except: prix_promo = None
            try: prix = float(pandas_data.at[i, "Prix"]) if "Prix" in pandas_data.columns and pandas_data.at[i, "Prix"] != '' else None
            except: prix = None
            try: quantite = int(pandas_data.at[i, "Quantité"]) if "Quantité" in pandas_data.columns and pandas_data.at[i, "Quantité"] != '' else 0
            except: quantite = None

            lign.update({
                'CODE A BARRE': code_a_barre, 'Code * Références_orig': code_ref_orig,
                'Code  Références': code_ref_processed, 'DESIGNATION': designation,
                'Age': "Adulte" if age_orig == "Homme" or age_orig == "Femme" else age_orig,
                'Taille': taille_orig.replace("\\", '/').replace('.6666666666666664', ' ⅔').replace('.666666666666664', ' ⅔').replace('2/3',' ⅔').replace('.5', ' ½').replace('1/2', '½').replace('.3333333333333336', ' ⅓').replace('.333333333333336', ' ⅓').replace('.0', ''),
                'Marques': str(pandas_data.at[i, "Marques"]) if "Marques" in pandas_data.columns and pd.notna(pandas_data.at[i, "Marques"]) else '',
                'Main Division': str(pandas_data.at[i, "Main Division"]) if "Main Division" in pandas_data.columns and pd.notna(pandas_data.at[i, "Main Division"]) else '',
                'Rayon': str(pandas_data.at[i, "Rayon"]) if "Rayon" in pandas_data.columns and pd.notna(pandas_data.at[i, "Rayon"]) else '',
                'Gender': str(pandas_data.at[i, "Gender"]) if "Gender" in pandas_data.columns and pd.notna(pandas_data.at[i, "Gender"]) else '',
                'Prix promos': prix_promo, 'Prix': prix, 'Quantité': quantite,
                'Description': str(pandas_data.at[i, "Description"]) if "Description" in pandas_data.columns and pd.notna(pandas_data.at[i, "Description"]) else '',
                'Couleurs': str(pandas_data.at[i, "Couleurs"]) if "Couleurs" in pandas_data.columns and pd.notna(pandas_data.at[i, "Couleurs"]) else '',
                'Couleur P': str(pandas_data.at[i, "Couleurs 2nd"]) if "Couleurs 2nd" in pandas_data.columns and pd.notna(pandas_data.at[i, "Couleurs 2nd"]) else '',
                'Date de début de promo': date_debut_promo_str,
                'Date de fin de promo': date_fin_promo_str
            })
            data_list_full.append(lign)
            data_var_key = designation + lign['Gender']
            data_var_aggregated[data_var_key] = lign.copy()
        except KeyError as ke: log_messages.append(f"AVERT: Colonne manquante '{ke}' data.xlsx ligne {i+2}."); continue
        except Exception as e_parse: log_messages.append(f"ERREUR parsing data.xlsx ligne {i+2}: {e_parse}"); continue
    log_messages.append(f"Lignes data: {len(pandas_data)}, Articles uniques (pre-agg): {len(data_var_aggregated)}")

    for key_agg, art_base_val in data_var_aggregated.items():
        art_base_val.update({'Agg_Taille':set(), 'Agg_Couleurs':set(), 'Agg_Age':set(), 'Agg_Couleur_P':set(), 'Agg_Quantité':0})
        art_base_val['Agg_Taille'].add(str(art_base_val['Taille'])); art_base_val['Agg_Couleurs'].add(str(art_base_val['Couleurs']))
        art_base_val['Agg_Age'].add(str(art_base_val['Age'])); art_base_val['Agg_Couleur_P'].add(str(art_base_val['Couleur P']))
        current_q = 0
        for art_detail in data_list_full:
            if art_base_val['DESIGNATION']==art_detail['DESIGNATION'] and art_base_val['Gender']==art_detail['Gender']:
                art_base_val['Agg_Taille'].add(str(art_detail['Taille'])); art_base_val['Agg_Couleurs'].add(str(art_detail['Couleurs']))
                art_base_val['Agg_Age'].add(str(art_detail['Age'])); art_base_val['Agg_Couleur_P'].add(str(art_detail['Couleur P']))
                current_q += art_detail.get('Quantité', 0)
        art_base_val['Agg_Quantité'] = current_q
    for k_final_agg in data_var_aggregated:
        data_var_aggregated[k_final_agg]['Taille'] = ",".join(sorted(filter(None, data_var_aggregated[k_final_agg]['Agg_Taille'])))
        data_var_aggregated[k_final_agg]['Couleurs'] = ",".join(sorted(filter(None, data_var_aggregated[k_final_agg]['Agg_Couleurs'])))
        data_var_aggregated[k_final_agg]['Age'] = ",".join(sorted(filter(None, data_var_aggregated[k_final_agg]['Agg_Age'])))
        data_var_aggregated[k_final_agg]['Couleur P'] = ",".join(sorted(filter(None, data_var_aggregated[k_final_agg]['Agg_Couleur_P'])))
        data_var_aggregated[k_final_agg]['Quantité'] = data_var_aggregated[k_final_agg]['Agg_Quantité']
    log_messages.append(f"Articles agrégés (parents) = {len(data_var_aggregated)}")

    workbook = xlwt.Workbook(encoding='utf-8'); worksheet = workbook.add_sheet('Globale')
    style = xlwt.easyxf("protection: cell_locked 0;")
    header = [
        "Type","UGS","Nom","Publié","Mis en avant ?","Visibilité dans le catalogue","Description courte","Description","Date de début de promo","Date de fin de promo",
        "État de la TVA","Classe TVA","En stock ?","Stock","Montant de stock faible","Autoriser les commandes de produits en rupture ?","Vendre individuellement ?",
        "Autoriser les avis clients ?","Note de l'achat","Prix de vente","Prix de base","Catégories","Tags","Classe de livraison","Images","Limite de téléchargement",
        "Jours d'expiration du téléchargement","Parent","Groupes de produits","Produits suggérés","Ventes croisées","URL externe","Libellé du bouton",
        "Brand","EAN","Woo Variation Gallery Images","Swatches Attributes","Nom de l'attribut 1","Valeur(s) de l'attribut 1","Attribut 1 visible","Attribut 1 global",
        "Nom de l'attribut 2","Valeur(s) de l'attribut 2","Attribut 2 visible","Attribut 2 global","Attribut 2 par défaut",
        "Nom de l'attribut 4","Valeur(s) de l'attribut 4","Attribut 4 visible","Attribut 4 global","Attribut 4 par défaut",
        "Méta : _wcj_purchase_price","Nom de l'attribut 5","Valeur(s) de l'attribut 5","Attribut 5 visible","Attribut 5 global","Attribut 5 par défaut",
        "Attribut 1 par défaut","Nom de l'attribut 6","Valeur(s) de l'attribut 6","Attribut 6 global","Attribut 6 visible",
        "Unnamed: 62", "Unnamed: 63"
    ]
    for col_num, value in enumerate(header): worksheet.write(0, col_num, value, style)
    row_idx = 1; nb_img_introuvable_total = 0; nb_cate_introuvable_total = 0

    for art_parent in data_var_aggregated.values():
        worksheet.write(row_idx,0,"variable",style); code_ref_p=str(art_parent.get('Code  Références','')); worksheet.write(row_idx,1,code_ref_p,style)
        worksheet.write(row_idx,2,art_parent.get('DESIGNATION',''),style); worksheet.write(row_idx,3,1,style); worksheet.write(row_idx,4,0,style)
        worksheet.write(row_idx,5,"visible",style); worksheet.write(row_idx,6,art_parent.get('Description',''),style)
        worksheet.write(row_idx,7,"",style)
        worksheet.write(row_idx,8,"",style); worksheet.write(row_idx,9,"",style)
        worksheet.write(row_idx,10,"taxable",style); worksheet.write(row_idx,11,"",style)
        worksheet.write(row_idx,12,1 if art_parent.get('Quantité',0)>0 else 0,style); worksheet.write(row_idx,13,art_parent.get('Quantité',0),style)
        worksheet.write(row_idx,14,"",style); worksheet.write(row_idx,15,0,style); worksheet.write(row_idx,16,0,style); worksheet.write(row_idx,17,0,style)
        worksheet.write(row_idx,18,"",style); worksheet.write(row_idx,19,"",style); worksheet.write(row_idx,20,"",style)
        final_cate = ""
        conf_cols=["Catégorie Woocommerce","Division","Breakout","Age","Sexe"]
        for col in conf_cols:
            if col not in pandas_conf.columns: pandas_conf[col] = ''
            pandas_conf[col] = pandas_conf[col].astype(str)
        for k_conf in range(len(pandas_conf)):
            match_division, match_breakout, match_age, match_sexe = False, False, False, False
            cat_woo = pandas_conf.at[k_conf,"Catégorie Woocommerce"]
            conf_div_str = pandas_conf.at[k_conf,"Division"].lower().strip()
            conf_breakout = pandas_conf.at[k_conf,"Breakout"].lower().strip()
            conf_age_str = pandas_conf.at[k_conf,"Age"].lower().strip()
            conf_sexe_str = pandas_conf.at[k_conf,"Sexe"].lower().strip()
            art_p_main_div_str = str(art_parent.get('Main Division', '')).lower()
            art_p_rayon = str(art_parent.get('Rayon','')).lower().strip()
            art_p_age_agg = str(art_parent.get('Age','')).lower()
            art_p_gender = str(art_parent.get('Gender','')).lower().strip()
            if not conf_div_str: match_division = True
            else:
                conf_div_list = [d.strip() for d in conf_div_str.split(',') if d.strip()]
                if conf_div_list:
                    all_rule_terms_found = True
                    for rule_term in conf_div_list:
                        if rule_term not in art_p_main_div_str: all_rule_terms_found = False; break
                    if all_rule_terms_found: match_division = True
            if not conf_breakout or conf_breakout == art_p_rayon: match_breakout = True
            if not conf_age_str: match_age = True
            else:
                 art_p_age_list = [a.strip() for a in art_p_age_agg.split(',') if a.strip()]
                 conf_age_list = [a.strip() for a in conf_age_str.split(',') if a.strip()]
                 if art_p_age_list and conf_age_list:
                     if any(any(a_art in a_conf for a_conf in conf_age_list) for a_art in art_p_age_list): match_age = True
            if not conf_sexe_str: match_sexe = True
            else:
                conf_sexe_list_split = [s.strip() for s in conf_sexe_str.split(',') if s.strip()]
                product_gender_list_split = [g.strip() for g in art_p_gender.split(',') if g.strip()] # Split product gender too!

                # Check if any product gender term is present as a substring in any config gender term
                # This matches the logic used for Age
                if product_gender_list_split and conf_sexe_list_split:
                     match_sexe = any(any(pg_term in cs_term for cs_term in conf_sexe_list_split) for pg_term in product_gender_list_split)



            if match_division and match_breakout and match_age and match_sexe and cat_woo:
                final_cate=f"{final_cate},{cat_woo}" if final_cate else cat_woo
        if not final_cate:log_messages.append(f"Catégorie non trouvée: UGS {code_ref_p}");nb_cate_introuvable_total+=1
        worksheet.write(row_idx,21,final_cate,style)
        worksheet.write(row_idx,22,"",style); worksheet.write(row_idx,23,"",style)
        parent_img_urls=[]; search_p_img=str(art_parent.get('Code  Références','')).lower()
        if search_p_img:
            matched_p_imgs_fs=sorted([img_f for img_f in list_images_fs if search_p_img in str(img_f).lower()])
            for img_fn in matched_p_imgs_fs:
                if len(parent_img_urls)<8:parent_img_urls.append(f"{image_base_url_param}{img_fn}")
        worksheet.write(row_idx,24,",".join(parent_img_urls),style)
        if not parent_img_urls and list_images_fs and search_p_img :log_messages.append(f"Image parent non trouvée: {code_ref_p}");nb_img_introuvable_total+=1
        for c_idx in range(25, 33): worksheet.write(row_idx, c_idx, "", style)
        worksheet.write(row_idx,33,art_parent.get('Marques',''),style); worksheet.write(row_idx,34,"",style); worksheet.write(row_idx,35,"",style); worksheet.write(row_idx,36,"",style)
        worksheet.write(row_idx,37,"Couleur",style);worksheet.write(row_idx,38,art_parent.get('Couleurs',''),style); worksheet.write(row_idx,39,1,style);worksheet.write(row_idx,40,1,style)
        art_p_ray_l=str(art_parent.get('Rayon','')).lower(); p_taille_agg=art_parent.get('Taille',''); def_taille_p=p_taille_agg.split(',')[0] if ',' in p_taille_agg else p_taille_agg
        if "chaussure" in art_p_ray_l:
            for c_i in range(41,46): worksheet.write(row_idx,c_i,"",style)
            worksheet.write(row_idx,46,"Pointure",style); worksheet.write(row_idx,47,p_taille_agg,style); worksheet.write(row_idx,48,1,style); worksheet.write(row_idx,49,1,style); worksheet.write(row_idx,50,def_taille_p,style)
            worksheet.write(row_idx,51,"",style)
            for c_i in range(52,57): worksheet.write(row_idx,c_i,"",style)
        elif "chaussette" in art_p_ray_l:
            worksheet.write(row_idx,41,"Chaussettes",style); worksheet.write(row_idx,42,p_taille_agg,style); worksheet.write(row_idx,43,1,style); worksheet.write(row_idx,44,1,style); worksheet.write(row_idx,45,def_taille_p,style)
            for c_i in range(46,51): worksheet.write(row_idx,c_i,"",style)
            worksheet.write(row_idx,51,"",style)
            for c_i in range(52,57): worksheet.write(row_idx,c_i,"",style)
        else:
            for c_i in range(41,46): worksheet.write(row_idx,c_i,"",style)
            for c_i in range(46,51): worksheet.write(row_idx,c_i,"",style)
            worksheet.write(row_idx,51,"",style)
            worksheet.write(row_idx,52,"Taille",style); worksheet.write(row_idx,53,p_taille_agg,style); worksheet.write(row_idx,54,1,style); worksheet.write(row_idx,55,1,style); worksheet.write(row_idx,56,def_taille_p,style)
        p_couleurs_agg=art_parent.get('Couleurs','');def_couleur_p=p_couleurs_agg.split(',')[0] if ',' in p_couleurs_agg else p_couleurs_agg
        worksheet.write(row_idx,57,def_couleur_p,style)
        worksheet.write(row_idx,58,"Couleur P",style);worksheet.write(row_idx,59,art_parent.get('Couleur P',''),style)
        worksheet.write(row_idx,60,1,style);worksheet.write(row_idx,61,0,style)
        worksheet.write(row_idx, 62, art_parent.get('Date de début de promo', ''), style)
        worksheet.write(row_idx, 63, art_parent.get('Date de fin de promo', ''), style)
        row_idx+=1

        for art_var in data_list_full:
            if art_parent.get('DESIGNATION','')==art_var.get('DESIGNATION','') and art_parent.get('Gender','')==art_var.get('Gender',''):
                worksheet.write(row_idx,0,"variation",style)
                var_taille = art_var.get('Taille','')
                var_code_ref = art_var.get('Code  Références','')
                var_ugs = f"{var_code_ref}-{var_taille}" if var_taille else var_code_ref
                worksheet.write(row_idx,1,var_ugs,style)
                worksheet.write(row_idx,2,art_var.get('DESIGNATION',''),style); worksheet.write(row_idx,3,1,style); worksheet.write(row_idx,5,"visible",style)
                worksheet.write(row_idx,4,0,style)
                worksheet.write(row_idx,6,"",style); worksheet.write(row_idx,7,"",style)
                worksheet.write(row_idx,8,"",style); worksheet.write(row_idx,9,"",style)
                worksheet.write(row_idx,10,"taxable",style); worksheet.write(row_idx,11,"parent",style)
                worksheet.write(row_idx,12,1 if art_var.get('Quantité',0)>0 else 0,style); worksheet.write(row_idx,13,art_var.get('Quantité',0),style)
                worksheet.write(row_idx,14,"",style)
                worksheet.write(row_idx,15,0,style); worksheet.write(row_idx,16,0,style); worksheet.write(row_idx,17,0,style)
                worksheet.write(row_idx,18,"",style)
                worksheet.write(row_idx,19,art_var.get('Prix promos',0.0),style); worksheet.write(row_idx,20,art_var.get('Prix',0.0),style)
                worksheet.write(row_idx,21,"",style); worksheet.write(row_idx,22,"",style); worksheet.write(row_idx,23,"",style)
                var_main_img_url="";search_v_img=str(var_code_ref).lower()
                var_gallery_urls=[]
                if search_v_img:
                    matched_v_imgs_fs=sorted([img_f for img_f in list_images_fs if search_v_img in str(img_f).lower()])
                    if matched_v_imgs_fs:var_main_img_url=f"{image_base_url_param}{matched_v_imgs_fs[0]}"
                    if len(matched_v_imgs_fs)>1:
                        for img_idx_g in range(1,min(len(matched_v_imgs_fs),6)):var_gallery_urls.append(f"{image_base_url_param}{matched_v_imgs_fs[img_idx_g]}")
                worksheet.write(row_idx,24,var_main_img_url,style)
                worksheet.write(row_idx,35,",".join(var_gallery_urls),style)
                worksheet.write(row_idx,25,"",style); worksheet.write(row_idx,26,"",style)
                worksheet.write(row_idx,27,code_ref_p,style)
                for c_idx in range(28, 34): worksheet.write(row_idx, c_idx, "", style)
                worksheet.write(row_idx,34,str(art_var.get('CODE A BARRE','')).replace('.0',''),style)
                worksheet.write(row_idx,36,"",style)
                worksheet.write(row_idx,37,"Couleur",style);worksheet.write(row_idx,38,art_var.get('Couleurs',''),style)
                worksheet.write(row_idx,39,"",style);worksheet.write(row_idx,40,1,style)
                if "chaussure" in art_p_ray_l:
                    worksheet.write(row_idx,41,"",style); worksheet.write(row_idx,42,"",style); worksheet.write(row_idx,43,"",style); worksheet.write(row_idx,44,"",style); worksheet.write(row_idx,45,"",style)
                    worksheet.write(row_idx,46,"Pointure",style);worksheet.write(row_idx,47,var_taille,style); worksheet.write(row_idx,48,"",style);worksheet.write(row_idx,49,1,style);worksheet.write(row_idx,50,"",style)
                    worksheet.write(row_idx,51,"",style)
                    worksheet.write(row_idx,52,"",style); worksheet.write(row_idx,53,"",style); worksheet.write(row_idx,54,"",style); worksheet.write(row_idx,55,"",style); worksheet.write(row_idx,56,"",style)
                elif "chaussette" in art_p_ray_l:
                    worksheet.write(row_idx,41,"Chaussettes",style);worksheet.write(row_idx,42,var_taille,style); worksheet.write(row_idx,43,"",style);worksheet.write(row_idx,44,1,style);worksheet.write(row_idx,45,"",style)
                    worksheet.write(row_idx,46,"",style); worksheet.write(row_idx,47,"",style); worksheet.write(row_idx,48,"",style); worksheet.write(row_idx,49,"",style); worksheet.write(row_idx,50,"",style)
                    worksheet.write(row_idx,51,"",style)
                    worksheet.write(row_idx,52,"",style); worksheet.write(row_idx,53,"",style); worksheet.write(row_idx,54,"",style); worksheet.write(row_idx,55,"",style); worksheet.write(row_idx,56,"",style)
                else:
                    worksheet.write(row_idx,41,"",style); worksheet.write(row_idx,42,"",style); worksheet.write(row_idx,43,"",style); worksheet.write(row_idx,44,"",style); worksheet.write(row_idx,45,"",style)
                    worksheet.write(row_idx,46,"",style); worksheet.write(row_idx,47,"",style); worksheet.write(row_idx,48,"",style); worksheet.write(row_idx,49,"",style); worksheet.write(row_idx,50,"",style)
                    worksheet.write(row_idx,51,"",style)
                    worksheet.write(row_idx,52,"Taille",style);worksheet.write(row_idx,53,var_taille,style); worksheet.write(row_idx,54,"",style);worksheet.write(row_idx,55,1,style);worksheet.write(row_idx,56,"",style)
                worksheet.write(row_idx,57,"",style)
                worksheet.write(row_idx,58,"Couleur P",style);worksheet.write(row_idx,59,art_var.get('Couleur P',''),style)
                worksheet.write(row_idx,60,"",style);worksheet.write(row_idx,61,0,style)
                worksheet.write(row_idx, 62, art_var.get('Date de début de promo', ''), style)
                worksheet.write(row_idx, 63, art_var.get('Date de fin de promo', ''), style)
                row_idx+=1

    log_messages.append("--- Statistiques ---")
    log_messages.append(f"Images introuvables: {nb_img_introuvable_total}")
    log_messages.append(f"Catégories non trouvées: {nb_cate_introuvable_total}")
    output_excel_fn=f'globale_{data_file_num_str}.xls';output_excel_path=os.path.join(GLOBALE_DIR,output_excel_fn)
    output_csv_fn=f'globale_{data_file_num_str}.csv';output_csv_path=os.path.join(GLOBALE_DIR,output_csv_fn)
    try:
        workbook.save(output_excel_path);log_messages.append(f"Excel '{output_excel_fn}' généré.")
        read_file_for_csv=pd.read_excel(output_excel_path,dtype=str,keep_default_na=False)
        if "Unnamed: 62" not in read_file_for_csv.columns: read_file_for_csv["Unnamed: 62"] = ""
        if "Unnamed: 63" not in read_file_for_csv.columns: read_file_for_csv["Unnamed: 63"] = ""
        read_file_for_csv.to_csv(output_csv_path,index=None,header=True,encoding="utf-8-sig")
        log_messages.append(f"CSV '{output_csv_fn}' généré.")
        st.session_state.preview_df=read_file_for_csv
        st.session_state.processing_log.extend(log_messages)
        return output_csv_path
    except Exception as e:
        log_messages.append(f"ERREUR sauvegarde/conversion Excel/CSV: {e}")
        st.session_state.preview_df=None
        st.session_state.processing_log.extend(log_messages)
        return None


# --- Sidebar Actions & Download ---
st.sidebar.markdown("---")
st.sidebar.header("2. Actions")
if st.sidebar.button("Exécuter le script", key="execute_button", type="primary"):
    st.session_state.processing_log = []; st.session_state.preview_df = None # Clear previous results
    try:
        # Use the counters from session state which reflect the latest successful upload
        data_f_num = str(st.session_state.data_file_count)
        conf_f_num = str(st.session_state.conf_file_count)
    except Exception as e:
        st.error(f"Erreur lecture numéros de fichiers depuis l'état: {e}")
        data_f_num=None;conf_f_num=None

    if data_f_num and conf_f_num and data_f_num!="0" and conf_f_num!="0":
        with st.spinner("Traitement en cours..."):
            gen_csv_path=process_files_core(data_f_num,conf_f_num,st.session_state.image_base_url)
        if gen_csv_path:
            st.success(f"Terminé! Fichier '{os.path.basename(gen_csv_path)}' prêt.")
            st.session_state.last_generated_csv=gen_csv_path
        else:
            st.error("Echec du traitement. Veuillez consulter les logs pour plus de détails.")
            st.session_state.last_generated_csv=None
    else:
        st.warning("Veuillez charger un fichier data ET un fichier de configuration avant d'exécuter (les compteurs doivent être > 0).")
        st.session_state.preview_df = None # Ensure no stale preview if files not ready

if st.sidebar.button("Réinitialiser l'état de l'application", key="clear_state_button"):
    st.session_state.last_generated_csv = None
    st.session_state.preview_df = None
    st.session_state.processing_log = []
    # st.session_state.image_base_url = DEFAULT_IMAGE_BASE_URL # Uncomment to reset URL too
    if 'latest_data_file_name' in st.session_state:
        st.session_state.latest_data_file_name = None # Clear displayed name
    if 'latest_conf_file_name' in st.session_state:
        st.session_state.latest_conf_file_name = None # Clear displayed name
    st.sidebar.info("État de l'application réinitialisé (logs, aperçus, dernier CSV effacés).")
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.header("3. Téléchargement")
if st.session_state.last_generated_csv and os.path.exists(st.session_state.last_generated_csv):
    with open(st.session_state.last_generated_csv,"rb") as fp:
        st.sidebar.download_button("Télécharger dernier fichier généré (.csv)",fp,os.path.basename(st.session_state.last_generated_csv),"text/csv",key="dl_csv_btn")
elif st.session_state.data_file_count > 0 : # Check if any data file has ever been uploaded
    if st.sidebar.button("Préparer dernier fichier connu pour téléchargement",key="prep_dl_btn"):
        try:
            # Use session state counters as they reflect the latest known successful upload
            latest_data_num = str(st.session_state.data_file_count)
            if latest_data_num!="0":
                potential_csv=os.path.join(GLOBALE_DIR,f'globale_{latest_data_num}.csv')
                if os.path.exists(potential_csv):
                    st.session_state.last_generated_csv=potential_csv
                    try:
                        st.session_state.preview_df=pd.read_csv(potential_csv,dtype=str,keep_default_na=False,na_filter=False)
                        st.sidebar.success("Fichier CSV précédent trouvé et chargé pour aperçu/téléchargement.")
                        st.rerun() # Rerun to update UI with download button and preview
                    except Exception as e_read:
                        st.session_state.preview_df=None
                        st.sidebar.warning(f"Erreur lors de la lecture du fichier CSV pour l'aperçu: {e_read}")
                else:
                    st.sidebar.warning(f"Le fichier CSV '{os.path.basename(potential_csv)}' correspondant au dernier compteur n'a pas été trouvé. Veuillez (ré)exécuter le script.")
                    st.session_state.preview_df=None
            else:
                st.sidebar.warning("Le compteur de fichiers data est à 0. Aucun fichier à préparer.")
                st.session_state.preview_df=None
        except Exception as e:
            st.sidebar.error(f"Erreur lors de la recherche du fichier: {e}")
            st.session_state.preview_df=None

# --- Main Page Tabs ---
tab_log, tab_stats, tab_preview_detail, tab_preview_raw = st.tabs([
    "📝 Log du Traitement",
    "📊 Statistiques",
    "🔍 Aperçu Détaillé des Produits",
    "📄 Aperçu CSV Brut"
])

with tab_log:
    st.subheader("Log du Traitement :")
    if st.session_state.processing_log:
        log_html="<div style='max-height:400px;overflow-y:auto;border:1px solid #ccc;padding:10px;border-radius:5px;background-color:#f9f9f9;'>"
        for msg in st.session_state.processing_log:
            color="black";
            if "ERREUR" in msg:color="red"; msg = f"🔴 {msg}"
            elif "AVERTISSEMENT" in msg or "non trouvée" in msg or "AVERT:" in msg: color="orange"; msg = f"🟠 {msg}"
            elif "--- Statistiques ---" in msg or "généré" in msg or "Début:" in msg:color="green"; msg = f"🟢 {msg}"
            log_html+=f"<p style='color:{color};margin-bottom:5px;font-family:monospace;font-size:0.9em;'>{msg}</p>"
        log_html+="</div>"; st.markdown(log_html, unsafe_allow_html=True)
    else:
        st.info("Aucun log de traitement disponible. Exécutez le script pour générer des logs.")

with tab_stats:
    st.subheader("Statistiques du Fichier Généré")
    if st.session_state.preview_df is not None and not st.session_state.preview_df.empty:
        df_stats = st.session_state.preview_df
        total_rows = len(df_stats)
        parent_products_stats = df_stats[df_stats['Type'].astype(str).str.lower() == 'variable']
        variation_products_stats = df_stats[df_stats['Type'].astype(str).str.lower() == 'variation']
        num_parents = len(parent_products_stats)
        num_variations = len(variation_products_stats)
        parents_missing_images = parent_products_stats[parent_products_stats['Images'].astype(str).fillna('') == ''].shape[0]
        parents_missing_categories = parent_products_stats[parent_products_stats['Catégories'].astype(str).fillna('') == ''].shape[0]

        # Ensure 'Stock' column exists and handle potential errors converting to numeric
        if 'Stock' in df_stats.columns:
            df_stats['Stock_numeric'] = pd.to_numeric(df_stats['Stock'], errors='coerce').fillna(0)
            total_stock = int(df_stats['Stock_numeric'].sum())
        else:
            total_stock = "N/A (Colonne 'Stock' manquante)"
            st.warning("La colonne 'Stock' est manquante dans le fichier généré. Le stock total ne peut être calculé.")


        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Lignes Totales", f"{total_rows}")
        col2.metric("Produits Parents", f"{num_parents}")
        col3.metric("Variations", f"{num_variations}")
        col4.metric("Stock Total (Toutes lignes)", f"{total_stock}")

        if num_parents > 0:
            # Calculate percentage of parents with images/categories
            parents_with_images = num_parents - parents_missing_images
            parents_with_categories = num_parents - parents_missing_categories

            st.metric(
                "Parents avec Images",
                f"{parents_with_images}/{num_parents}",
                delta=f"{(parents_with_images/num_parents*100):.1f}% trouvées",
                delta_color="normal" if parents_missing_images == 0 else ("inverse" if parents_missing_images > num_parents * 0.5 else "off") # More nuanced color
            )
            st.metric(
                "Parents avec Catégories",
                f"{parents_with_categories}/{num_parents}",
                delta=f"{(parents_with_categories/num_parents*100):.1f}% trouvées",
                delta_color="normal" if parents_missing_categories == 0 else ("inverse" if parents_missing_categories > num_parents * 0.5 else "off")
            )
        else:
            st.metric("Parents avec Images", "N/A"); st.metric("Parents avec Catégories", "N/A")
        st.markdown("---")
    else:
        st.info("Aucunes données disponibles pour les statistiques. Exécutez le script pour générer un fichier.")

with tab_preview_detail:
    st.subheader("Rapport d'Aperçu Détaillé des Produits")
    if st.session_state.preview_df is not None and not st.session_state.preview_df.empty:
        preview_data = st.session_state.preview_df.copy()
        preview_data['UGS'] = preview_data['UGS'].astype(str)
        preview_data['Parent'] = preview_data['Parent'].astype(str).fillna('')
        unnamed_col_62 = "Unnamed: 62" if "Unnamed: 62" in preview_data.columns else None
        unnamed_col_63 = "Unnamed: 63" if "Unnamed: 63" in preview_data.columns else None

        parent_products = preview_data[preview_data['Type'].astype(str).str.lower() == 'variable']
        if parent_products.empty:
            st.warning("Aucun produit parent (type 'variable') trouvé dans le fichier généré.")
        else:
            if len(parent_products) == 1:
                num_parents_to_preview = 1
                st.write(f"Affichage du seul produit parent trouvé:")
            else:
                num_parents_to_preview = st.slider("Nombre de produits parents à afficher en détail:",1,len(parent_products),min(3,len(parent_products)),1,key="parent_preview_slider")

            for i, parent_row in parent_products.head(num_parents_to_preview).iterrows():
                st.markdown("---"); st.markdown(f"#### Produit Parent: **{parent_row.get('Nom','N/A')}** (UGS: `{parent_row.get('UGS','N/A')}`)")
                col1,col2=st.columns([2,3])
                with col1:
                    st.markdown(f"**Marque:** {parent_row.get('Brand','N/A')}")
                    st.markdown(f"**Stock Total (Parent):** {parent_row.get('Stock','N/A')}")

                    promo_start_val = parent_row.get(unnamed_col_62, '') if unnamed_col_62 else ''
                    promo_end_val = parent_row.get(unnamed_col_63, '') if unnamed_col_63 else ''
                    if promo_start_val or promo_end_val:
                        st.markdown(f"**Dates de Promo (issues des colonnes brutes 62/63):**")
                        st.markdown(f"  Début: `{promo_start_val if promo_start_val else 'N/A'}`")
                        st.markdown(f"  Fin: `{promo_end_val if promo_end_val else 'N/A'}`")

                    st.markdown(f"**Catégories:**")
                    cats = parent_row.get('Catégories', '');
                    if cats:
                        for cat in cats.split(','):
                             if cat.strip(): st.caption(f"- {cat.strip()}")
                    else: st.caption("_Aucune_")

                    imgs_p=parent_row.get('Images','');
                    st.markdown(f"**Images (Parent):**")
                    if imgs_p:
                        img_list=[img.strip() for img in imgs_p.split(',') if img.strip()]
                        if img_list:
                            try: st.image(img_list[0], width=100, caption=f"Image principale ({len(img_list)} au total)")
                            except Exception as img_ex:
                                st.caption(f"! Erreur affichage Image: {img_list[0]}")
                                st.markdown(f"  URL: [{img_list[0]}]({img_list[0]})")
                            if len(img_list)>1:
                               with st.expander(f"Voir les {len(img_list)-1} autres images"):
                                   for il_idx, il in enumerate(img_list[1:]):
                                       st.markdown(f"- [{il}]({il})")
                                       # Optionally display more images:
                                       # try: st.image(il, width=80, caption=f"Image {il_idx+2}")
                                       # except: pass
                        else: st.caption("_Aucune URL d'image listée._")
                    else: st.caption("_Aucune_")
                with col2:
                    st.markdown("**Attributs (Parent):**")
                    att1_n=parent_row.get('Nom de l\'attribut 1','');att1_v=parent_row.get('Valeur(s) de l\'attribut 1','')
                    if att1_n and att1_v:st.markdown(f"- **{att1_n}:** {att1_v} (Défaut: `{parent_row.get('Attribut 1 par défaut','')}`)")
                    s_att_n,s_att_v,s_att_d="","",""
                    if parent_row.get('Nom de l\'attribut 2','') and parent_row.get('Valeur(s) de l\'attribut 2',''): s_att_n,s_att_v,s_att_d=parent_row.get('Nom de l\'attribut 2'),parent_row.get('Valeur(s) de l\'attribut 2'),parent_row.get('Attribut 2 par défaut')
                    elif parent_row.get('Nom de l\'attribut 4','') and parent_row.get('Valeur(s) de l\'attribut 4',''): s_att_n,s_att_v,s_att_d=parent_row.get('Nom de l\'attribut 4'),parent_row.get('Valeur(s) de l\'attribut 4'),parent_row.get('Attribut 4 par défaut')
                    elif parent_row.get('Nom de l\'attribut 5','') and parent_row.get('Valeur(s) de l\'attribut 5',''): s_att_n,s_att_v,s_att_d=parent_row.get('Nom de l\'attribut 5'),parent_row.get('Valeur(s) de l\'attribut 5'),parent_row.get('Attribut 5 par défaut')
                    if s_att_n and s_att_v:st.markdown(f"- **{s_att_n}:** {s_att_v} (Défaut: `{s_att_d}`)")
                    att6_n=parent_row.get('Nom de l\'attribut 6','');att6_v=parent_row.get('Valeur(s) de l\'attribut 6','')
                    if att6_n and att6_v: st.markdown(f"- **{att6_n}:** {att6_v}")
                    st.markdown("**Description courte:**");st.caption(f"{parent_row.get('Description courte','_Aucune_')}")

                parent_ugs_str=str(parent_row['UGS'])
                variations=preview_data[(preview_data['Type'].astype(str).str.lower()=='variation')&(preview_data['Parent'].astype(str)==parent_ugs_str)]
                if not variations.empty:
                    st.markdown("--- \n ##### Variations de ce Parent:")
                    var_disp_data=[]
                    for _, vr in variations.iterrows():
                        taille_val_2 = vr.get("Valeur(s) de l'attribut 2", '')
                        taille_val_4 = vr.get("Valeur(s) de l'attribut 4", '')
                        taille_val_5 = vr.get("Valeur(s) de l'attribut 5", '')
                        taille_pointure_preview = (taille_val_2 or taille_val_4 or taille_val_5)
                        if not taille_pointure_preview: taille_pointure_preview = 'N/A'
                        var_data_entry = {
                            "UGS Var.": vr.get('UGS','N/A'), "Taille/Pointure": taille_pointure_preview,
                            "Couleur": vr.get('Valeur(s) de l\'attribut 1','N/A'), "Couleur P": vr.get('Valeur(s) de l\'attribut 6','N/A'),
                            "Prix Vente": vr.get('Prix de vente','N/A'), "Prix Base": vr.get('Prix de base','N/A'),
                            "Stock": vr.get('Stock','N/A'), "EAN": vr.get('EAN','N/A'), "Image Var.": vr.get('Images','N/A')
                        }
                        var_disp_data.append(var_data_entry)
                    st.dataframe(pd.DataFrame(var_disp_data),hide_index=True, use_container_width=True)
                else:st.caption("_Aucune variation trouvée pour ce parent._")
            if len(parent_products)>num_parents_to_preview:st.info(f"Affichage détaillé des {num_parents_to_preview} premiers produits parents sur {len(parent_products)}. Modifiez le curseur ci-dessus pour en voir plus.")
    else:
        st.info("Aucunes données disponibles pour l'aperçu détaillé. Exécutez le script pour générer un fichier.")

with tab_preview_raw:
    st.subheader("Aperçu Tabulaire Complet (CSV Brut)")
    if st.session_state.preview_df is not None and not st.session_state.preview_df.empty:
        df_to_show = st.session_state.preview_df
        min_raw_rows=min(5,len(df_to_show));max_raw_rows=len(df_to_show);def_raw_rows=min(10,len(df_to_show))
        if max_raw_rows<min_raw_rows:min_raw_rows=max_raw_rows # Ensure min is not greater than max
        if def_raw_rows<min_raw_rows:def_raw_rows=min_raw_rows # Ensure default is not less than min
        if def_raw_rows > max_raw_rows:def_raw_rows = max_raw_rows # Ensure default is not more than max


        if max_raw_rows<=min_raw_rows and max_raw_rows>0:
            num_rows_raw_preview=max_raw_rows
            st.write(f"Affichage de toutes les {max_raw_rows} lignes de l'aperçu brut:")
        elif max_raw_rows > min_raw_rows:
            num_rows_raw_preview=st.slider("Nombre de lignes à afficher dans l'aperçu brut:",min_raw_rows,max_raw_rows,def_raw_rows, max(1, (max_raw_rows-min_raw_rows)//10),key="raw_preview_slider")
        else: # max_raw_rows is 0 or invalid scenario
            num_rows_raw_preview=max_raw_rows

        if num_rows_raw_preview > 0:
            st.dataframe(df_to_show.head(num_rows_raw_preview), use_container_width=True)
        else:
            st.info("Le fichier généré est vide.")

    elif st.session_state.processing_log and not st.session_state.last_generated_csv:
         st.info("Le traitement a été exécuté, mais aucun aperçu CSV n'est disponible (peut-être une erreur). Consultez le log.")
    else:
         st.info("Aucun fichier traité ou chargé pour l'aperçu brut.")


# --- Footer & Help Section ---
st.markdown("---")
st.caption(f"NOTE: La structure de dossiers de base est configurée sur le serveur à `{BASE_DIR}`. Les images doivent être placées dans `{IMAGES_DIR}`. Assurez-vous que l'URL de base pour les images est correctement configurée dans la barre latérale.")
st.caption("Built By Abdellah Zouggari")

with st.sidebar.expander("ℹ️ Aide et Informations", expanded=False):
    st.markdown(f"""
    **Bienvenue sur Madina Convertisseur!**

    Cette application transforme vos fichiers Excel de données produits et de configuration en un fichier CSV importable pour WooCommerce.

    **Instructions:**
    1.  **Configuration Globale (Barre Latérale):**
        *   **URL de base pour les images:** Entrez l'URL où vos images produits sont hébergées (ex: `https://www.votre-site.com/images/`). Les noms de fichiers images trouvés dans `{IMAGES_DIR}` seront ajoutés à cette URL.
        *   Le dossier pour les images sur le serveur est: `{IMAGES_DIR}`.
    2.  **Chargement des Fichiers (Barre Latérale):**
        *   **Fichier DATA (.xlsx):** Contient les détails des produits (ex: `CODE A BARRE`, `DESIGNATION`, `Prix`, `Quantité`).
        *   **Fichier CONFIG (.xlsx):** Pour le mappage des catégories (ex: `Catégorie Woocommerce`, `Division`, `Breakout`).
        *   Les fichiers chargés sont stockés sur le serveur et un compteur est incrémenté. L'application utilise toujours les derniers fichiers indiqués par ces compteurs pour le traitement.
    3.  **Actions (Barre Latérale):**
        *   **Exécuter le script:** Lance la conversion.
        *   **Réinitialiser l'état de l'application:** Efface les logs, aperçus, et le lien vers le dernier CSV généré de la session actuelle. Ne supprime pas les fichiers déjà chargés sur le serveur ni ne réinitialise les compteurs de fichiers.
    4.  **Téléchargement (Barre Latérale):**
        *   Téléchargez le fichier CSV généré.
        *   "Préparer dernier fichier connu..." tente de localiser le CSV basé sur le dernier compteur de fichiers data.

    **Onglets Principaux:**
    *   **Log du Traitement:** Affiche les messages du processus de conversion, y compris les erreurs et avertissements.
    *   **Statistiques:** Résumé du fichier généré (nombre de produits, stock, etc.).
    *   **Aperçu Détaillé:** Visualisation détaillée de quelques produits parents et leurs variations.
    *   **Aperçu CSV Brut:** Tableau des premières lignes du fichier CSV généré.

    **Structure des Dossiers (Côté Serveur):**
    *   Répertoire de Base: `{BASE_DIR}`
    *   Données Téléchargées: `{DATA_DIR}`
    *   Configurations Téléchargées: `{CONF_DIR}`
    *   Fichiers Générés: `{GLOBALE_DIR}`
    """)