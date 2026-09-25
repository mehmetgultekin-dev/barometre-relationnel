import streamlit as st
import pandas as pd
import openpyxl
import json
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, ColumnsAutoSizeMode
import hashlib
import io
import zipfile
from barometre_core import classer_relation, longueur_max_colonne, valider_donnees_projet

# Récupération sécurisée depuis les secrets Streamlit
USERNAME = st.secrets["auth"]["username"]
PASSWORD_HASH = st.secrets["auth"]["password_hash"]

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.title("🔒 Connexion sécurisée")
    login = st.text_input("Identifiant")
    password = st.text_input("Mot de passe", type="password")
    if st.button("Se connecter"):
        if login == USERNAME and hash_password(password) == PASSWORD_HASH:
            st.session_state.logged_in = True
            st.rerun()
        else:
            st.error("Identifiants incorrects. Veuillez réessayer.")
    st.stop()



# === INITIALISATION DES ÉTATS STREAMLIT ======================================
default_states = {
    "etat": "menu",
    "participants": [],
    "participant_ajoute_message": None,
    "services": [],
    "relations_saisies": [],
    "relation_a_modifier": None,
    "affiche_formulaire_participant": False,
    "participant_a_modifier": None,
    "selected_relations": pd.DataFrame(),
    "nombre_total_personnes": 0,
}
for k, v in default_states.items():
    st.session_state.setdefault(k, v)

# === UTILITAIRES D’IMPORT / EXPORT ===========================================
def exporter_json_data() -> str:
    """Exporte les données actuelles de la session en une chaîne JSON."""
    data = {
        "participants": st.session_state.participants,
        "services": st.session_state.services,
        "relations_saisies": st.session_state.relations_saisies,
        "nombre_total_personnes": st.session_state.nombre_total_personnes,
    }
    return json.dumps(data, indent=4, ensure_ascii=False)


def exporter_excel_data() -> bytes:
    """
    Exporte les relations saisies dans un fichier Excel avec un ordre de colonnes spécifique
    et ajuste automatiquement la largeur des colonnes. Ajoute également une feuille de calcul
    avec un récapitulatif de la vigilance des relations, une feuille pour les relations
    unidirectionnelles et une feuille pour les relations croisées négatives et positives.
    Inclut TOUTES les relations possibles (saisies et neutres, y compris celles
    impliquant des participants non nommés) dans la feuille 'Relations'.
    """
    all_relations_for_excel = []

    # Construire la liste de TOUS les participants (nommés + anonymes)
    all_person_names = [p['nom'] for p in st.session_state.participants]
    num_named_participants = len(all_person_names)
    total_persons_in_project = st.session_state.nombre_total_personnes

    if total_persons_in_project > num_named_participants:
        for i in range(total_persons_in_project - num_named_participants):
            all_person_names.append(f"Personne Anonyme {i+1}")

    # Créer un dictionnaire de hachages pour les relations saisies pour une recherche rapide
    saisies_map = {}
    for rel in st.session_state.relations_saisies:
        key = (rel["Émetteur"], rel["Récepteur"])
        saisies_map[key] = rel

    # Définir l'ordre exact des colonnes pour les feuilles Excel
    colonnes_relations_excel = [
        "Émetteur", "Récepteur", "Date", "Début", "Fin", "Service",
        "P+", "P-", "I+", "I-", "C+", "C-",
        "Score Pic Positif", "Score Pic Négatif", "Score Net",
        "Vigilance", "Commentaire"
    ]

    # Générer TOUTES les combinaisons bidirectionnelles possibles pour la feuille 'Relations'
    for emetteur in all_person_names:
        for recepteur in all_person_names:
            if emetteur != recepteur:
                current_key = (emetteur, recepteur)
                
                # Vérifier si cette relation a été saisie
                if current_key in saisies_map:
                    all_relations_for_excel.append(saisies_map[current_key])
                else:
                    # C'est une relation neutre (non saisie)
                    # Déterminer le service de l'émetteur si l'émetteur est nommé
                    service_emetteur = "RAS"
                    for p in st.session_state.participants:
                        if p["nom"] == emetteur:
                            service_emetteur = p["service"]
                            break

                    all_relations_for_excel.append({
                        "Émetteur": emetteur,
                        "Récepteur": recepteur,
                        "Date": "RAS", # Pas de date pour une relation non saisie
                        "Début": "RAS", # Pas d'heure
                        "Fin": "RAS",   # Pas d'heure
                        "Service": service_emetteur, # Service de l'émetteur (si nommé), sinon RAS
                        "P+": 0, "P-": 0, "I+": 0, "I-": 0, "C+": 0, "C-": 0,
                        "Score Pic Positif": 0,
                        "Score Pic Négatif": 0,
                        "Score Net": 0,
                        "Vigilance": "Neutre", # Marqué comme neutre
                        "Commentaire": "Relation non renseignée ou non applicable" # Commentaire explicatif
                    })
            
    df_relations = pd.DataFrame(all_relations_for_excel)

    # S'assurer que toutes les colonnes définies existent dans le DataFrame
    for col in colonnes_relations_excel:
        if col not in df_relations.columns:
            df_relations[col] = None

    # Réorganiser le DataFrame selon l'ordre des colonnes définies
    df_relations = df_relations[colonnes_relations_excel]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # --- Première feuille : Relations (Toutes les combinaisons, saisies et neutres) ---
        df_relations.to_excel(writer, index=False, sheet_name='Relations')

        workbook = writer.book
        worksheet_relations = writer.sheets['Relations']

        for col_idx, column_name in enumerate(colonnes_relations_excel):
            max_length = len(str(column_name))
            if not df_relations.empty:
                max_length = max(max_length, longueur_max_colonne(df_relations[column_name]))

            adjusted_width = (max_length + 2)
            if column_name == "Commentaire":
                adjusted_width = min(adjusted_width, 100) # Limite pour le commentaire
            elif column_name in ["Date", "Début", "Fin", "Service", "Vigilance"]:
                 adjusted_width = min(adjusted_width, 25) # Limite pour ces colonnes
            else:
                 adjusted_width = min(adjusted_width, 15) # Limite générale pour les autres

            worksheet_relations.column_dimensions[openpyxl.utils.get_column_letter(col_idx + 1)].width = adjusted_width

        # --- Deuxième feuille : Récapitulatif Vigilance (Résumé de la vigilance) ---
        worksheet_vigilance = workbook.create_sheet(title='Récapitulatif Vigilance')

        # Calcul des statistiques de vigilance
        relations_enregistrees_count = len(st.session_state.relations_saisies)
        
        # Obtenir les comptages de chaque type de vigilance existant dans les relations SAISIES
        df_saisies_for_stats = pd.DataFrame(st.session_state.relations_saisies)
        vigilance_counts = df_saisies_for_stats['Vigilance'].value_counts().to_dict() if not df_saisies_for_stats.empty else {}
        
        # Initialiser les types de relations pour s'assurer qu'ils apparaissent même s'ils sont à 0
        types_relation_stats = ["Positif pur", "Positif", "Mixte positif", "Mixte tendu", "Négatif pur", "Négatif", "Aucune donnée"] # Updated types
        stats_data = []

        # Calcul du nombre total de combinaisons possibles (basé sur nombre_total_personnes)
        nombre_total_personnes_app = st.session_state.nombre_total_personnes 
        nombre_combinaisons_possibles = 0
        if nombre_total_personnes_app > 1:
            nombre_combinaisons_possibles = nombre_total_personnes_app * (nombre_total_personnes_app - 1)
        
        if nombre_combinaisons_possibles == 0 and nombre_total_personnes_app > 1:
            st.error("ERREUR DE CALCUL : Le nombre de combinaisons possibles est zéro. Veuillez saisir un nombre total de personnes supérieur à 1.")
        
        # Calcul des relations neutres GLOBALES (nombre total possible - nombre de relations enregistrées)
        relations_neutres_globales = max(0, nombre_combinaisons_possibles - relations_enregistrees_count)

        # Ajouter les relations spécifiques (enregistrées)
        for rel_type in types_relation_stats:
            count = vigilance_counts.get(rel_type, 0)
            percentage = (count / nombre_combinaisons_possibles) if nombre_combinaisons_possibles > 0 else 0.0
            stats_data.append({"Type de relation": rel_type, "Nombre de cas": count, "Pourcentage": percentage})

        # Ajouter les relations neutres globales
        percentage_neutre_globale = (relations_neutres_globales / nombre_combinaisons_possibles) if nombre_combinaisons_possibles > 0 else 0.0
        stats_data.append({"Type de relation": "Neutre (Global)", "Nombre de cas": relations_neutres_globales, "Pourcentage": percentage_neutre_globale})
        
        # Ajouter la ligne "Total des combinaisons possibles" à la fin
        total_pourcentage = 1.0 if nombre_combinaisons_possibles > 0 else 0.0
        stats_data.append({"Type de relation": "Total des combinaisons possibles", "Nombre de cas": nombre_combinaisons_possibles, "Pourcentage": total_pourcentage})


        # Créer un DataFrame pour les statistiques
        df_stats = pd.DataFrame(stats_data)

        # Écrire le DataFrame des stats dans la deuxième feuille
        worksheet_vigilance.append(df_stats.columns.tolist())
        for row in df_stats.itertuples(index=False):
            worksheet_vigilance.append(list(row))

        # Ajuster la largeur des colonnes pour la feuille 'Récapitulatif Vigilance'
        for col_idx, column_name in enumerate(df_stats.columns):
            max_length = len(str(column_name))
            if not df_stats.empty:
                max_length = max(max_length, longueur_max_colonne(df_stats[column_name]))
            
            adjusted_width = (max_length + 2)
            if column_name == "Type de relation":
                adjusted_width = min(adjusted_width, 40)
            elif column_name == "Nombre de cas":
                adjusted_width = min(adjusted_width, 20)
            elif column_name == "Pourcentage":
                adjusted_width = min(adjusted_width, 20)
            
            worksheet_vigilance.column_dimensions[openpyxl.utils.get_column_letter(col_idx + 1)].width = adjusted_width
            
            # Format pour la colonne Pourcentage
            if column_name == "Pourcentage":
                for row_idx in range(2, worksheet_vigilance.max_row + 1):
                    cell = worksheet_vigilance.cell(row=row_idx, column=col_idx + 1)
                    cell.number_format = '0.00%'

        # --- Nouvelle feuille : Relations Unidirectionnelles ---
        worksheet_unidirectional = workbook.create_sheet(title='Relations Unidirectionnelles')
        
        # Filtrer relations_saisies pour inclure seulement les relations avec une vigilance définie
        # (excluant 'Aucune donnée' qui résulterait de P+=0 et P-=0)
        df_unidirectional = pd.DataFrame(st.session_state.relations_saisies)
        if not df_unidirectional.empty:
            df_unidirectional = df_unidirectional[df_unidirectional['Vigilance'] != 'Aucune donnée']
            # S'assurer de l'ordre des colonnes
            for col in colonnes_relations_excel:
                if col not in df_unidirectional.columns:
                    df_unidirectional[col] = None
            df_unidirectional = df_unidirectional[colonnes_relations_excel]
        
        df_unidirectional.to_excel(writer, index=False, sheet_name='Relations Unidirectionnelles')

        # Ajuster la largeur des colonnes pour la feuille 'Relations Unidirectionnelles'
        for col_idx, column_name in enumerate(colonnes_relations_excel):
            max_length = len(str(column_name))
            if not df_unidirectional.empty:
                max_length = max(max_length, longueur_max_colonne(df_unidirectional[column_name]))

            adjusted_width = (max_length + 2)
            if column_name == "Commentaire":
                adjusted_width = min(adjusted_width, 100)
            elif column_name in ["Date", "Début", "Fin", "Service", "Vigilance"]:
                 adjusted_width = min(adjusted_width, 25)
            else:
                 adjusted_width = min(adjusted_width, 15)

            worksheet_unidirectional.column_dimensions[openpyxl.utils.get_column_letter(col_idx + 1)].width = adjusted_width

        # --- Nouvelle feuille : Relations Croisées Négatives ---
        worksheet_negative_cross = workbook.create_sheet(title='Relations Croisées Négatives')

        negative_cross_relations_data = []
        
        # Créer un mapping pour une recherche rapide des relations saisies
        recorded_relations_lookup = {}
        for rel in st.session_state.relations_saisies:
            recorded_relations_lookup[(rel["Émetteur"], rel["Récepteur"])] = rel

        # Définir les types de vigilance considérés comme "négatifs" pour cette analyse
        # NOTE: "Mixte tendu" indique un "pic négatif" car p_moins >= p_plus
        # Ajout de "Négatif" aux types de vigilance négatifs pour les relations croisées
        negative_vigilance_types_for_cross = {"Négatif pur", "Négatif", "Mixte tendu"}
        
        # Itérer sur les paires uniques de participants nommés pour trouver les relations croisées
        processed_pairs = set() # Pour éviter de traiter (A,B) et (B,A) comme des paires différentes pour la logique
        for p1_obj in st.session_state.participants:
            for p2_obj in st.session_state.participants:
                p1 = p1_obj["nom"]
                p2 = p2_obj["nom"]
                
                if p1 == p2:
                    continue

                # S'assurer de traiter chaque paire une seule fois (ex: A,B mais pas B,A) pour la logique de paire
                pair_key = tuple(sorted((p1, p2)))
                if pair_key in processed_pairs:
                    continue
                processed_pairs.add(pair_key)

                # Obtenir les relations pour les deux directions
                rel_p1_to_p2 = recorded_relations_lookup.get((p1, p2))
                rel_p2_to_p1 = recorded_relations_lookup.get((p2, p1))

                # Vérifier si les deux relations existent et sont d'un type de vigilance négatif
                if (rel_p1_to_p2 and rel_p1_to_p2["Vigilance"] in negative_vigilance_types_for_cross and
                    rel_p2_to_p1 and rel_p2_to_p1["Vigilance"] in negative_vigilance_types_for_cross):
                    
                    # Déterminer le type de relation croisée ("Conflit" ou "Tension relationnelle")
                    type_de_croise = ""
                    is_p1_p2_pure_negative = rel_p1_to_p2["Vigilance"] == "Négatif pur"
                    is_p2_p1_pure_negative = rel_p2_to_p1["Vigilance"] == "Négatif pur"

                    # "Conflit" si les deux sont "Négatif pur" (3 pics négatifs de chaque côté)
                    if is_p1_p2_pure_negative and is_p2_p1_pure_negative:
                        type_de_croise = "Conflit"
                    # "Tension relationnelle" si au moins l'un des deux est un "pic négatif" (Négatif pur, Négatif ou Mixte tendu)
                    else:
                        type_de_croise = "Tension relationnelle"
                    
                    # Ajouter les deux relations à la liste, avec la nouvelle classification
                    # Copier la relation pour éviter de modifier l'objet original dans relations_saisies
                    rel_p1_to_p2_copy = rel_p1_to_p2.copy()
                    rel_p1_to_p2_copy["Type de Croisé"] = type_de_croise
                    negative_cross_relations_data.append(rel_p1_to_p2_copy)

                    rel_p2_to_p1_copy = rel_p2_to_p1.copy()
                    rel_p2_to_p1_copy["Type de Croisé"] = type_de_croise
                    negative_cross_relations_data.append(rel_p2_to_p1_copy)
        
        df_negative_cross = pd.DataFrame(negative_cross_relations_data)
        
        # Nouvel ordre de colonnes pour les relations croisées négatives, incluant le type de croisé
        colonnes_negative_cross_excel = colonnes_relations_excel + ["Type de Croisé"]

        if not df_negative_cross.empty:
            # S'assurer que toutes les colonnes définies existent dans le DataFrame
            for col in colonnes_negative_cross_excel:
                if col not in df_negative_cross.columns:
                    df_negative_cross[col] = None
            # Réorganiser le DataFrame selon l'ordre des colonnes définies
            df_negative_cross = df_negative_cross[colonnes_negative_cross_excel]

        df_negative_cross.to_excel(writer, index=False, sheet_name='Relations Croisées Négatives')

        # Ajuster la largeur des colonnes pour la feuille 'Relations Croisées Négatives'
        for col_idx, column_name in enumerate(colonnes_negative_cross_excel):
            max_length = len(str(column_name))
            if not df_negative_cross.empty:
                max_length = max(max_length, longueur_max_colonne(df_negative_cross[column_name]))

            adjusted_width = (max_length + 2)
            if column_name == "Commentaire":
                adjusted_width = min(adjusted_width, 100)
            elif column_name in ["Date", "Début", "Fin", "Service", "Vigilance", "Type de Croisé"]:
                 adjusted_width = min(adjusted_width, 25)
            else:
                 adjusted_width = min(adjusted_width, 15)

            worksheet_negative_cross.column_dimensions[openpyxl.utils.get_column_letter(col_idx + 1)].width = adjusted_width


        # --- Nouvelle feuille : Relations Croisées Positives ---
        worksheet_positive_cross = workbook.create_sheet(title='Relations Croisées Positives')

        positive_cross_relations_data = []
        
        # Définir les types de vigilance considérés comme "positifs" pour cette analyse
        positive_vigilance_types_for_cross = {"Positif pur", "Positif"}
        
        # Itérer sur les paires uniques de participants nommés pour trouver les relations croisées
        # Nous réutilisons processed_pairs pour éviter les doublons A-B et B-A pour les relations positives
        processed_pairs_positive = set()
        for p1_obj in st.session_state.participants:
            for p2_obj in st.session_state.participants:
                p1 = p1_obj["nom"]
                p2 = p2_obj["nom"]
                
                if p1 == p2:
                    continue

                pair_key_positive = tuple(sorted((p1, p2)))
                if pair_key_positive in processed_pairs_positive:
                    continue
                processed_pairs_positive.add(pair_key_positive)

                # Obtenir les relations pour les deux directions
                rel_p1_to_p2 = recorded_relations_lookup.get((p1, p2))
                rel_p2_to_p1 = recorded_relations_lookup.get((p2, p1))

                # Vérifier si les deux relations existent et sont d'un type de vigilance positif
                if (rel_p1_to_p2 and rel_p1_to_p2["Vigilance"] in positive_vigilance_types_for_cross and
                    rel_p2_to_p1 and rel_p2_to_p1["Vigilance"] in positive_vigilance_types_for_cross):
                    
                    # Déterminer le type de relation croisée positive
                    type_de_croise = ""
                    is_p1_p2_pure_positive = rel_p1_to_p2["Vigilance"] == "Positif pur"
                    is_p2_p1_pure_positive = rel_p2_to_p1["Vigilance"] == "Positif pur"

                    # "Harmonie Parfaite" si les deux sont "Positif pur"
                    if is_p1_p2_pure_positive and is_p2_p1_pure_positive:
                        type_de_croise = "Harmonie Parfaite"
                    # "Harmonie Relationnelle" si au moins l'un des deux est "Positif" ou un mix avec "Positif pur"
                    else:
                        type_de_croise = "Harmonie Relationnelle"
                    
                    # Ajouter les deux relations à la liste, avec la nouvelle classification
                    rel_p1_to_p2_copy = rel_p1_to_p2.copy()
                    rel_p1_to_p2_copy["Type de Croisé"] = type_de_croise
                    positive_cross_relations_data.append(rel_p1_to_p2_copy)

                    rel_p2_to_p1_copy = rel_p2_to_p1.copy()
                    rel_p2_to_p1_copy["Type de Croisé"] = type_de_croise
                    positive_cross_relations_data.append(rel_p2_to_p1_copy)
        
        df_positive_cross = pd.DataFrame(positive_cross_relations_data)
        
        # Nouvel ordre de colonnes pour les relations croisées positives, incluant le type de croisé
        colonnes_positive_cross_excel = colonnes_relations_excel + ["Type de Croisé"]

        if not df_positive_cross.empty:
            # S'assurer que toutes les colonnes définies existent dans le DataFrame
            for col in colonnes_positive_cross_excel:
                if col not in df_positive_cross.columns:
                    df_positive_cross[col] = None
            # Réorganiser le DataFrame selon l'ordre des colonnes définies
            df_positive_cross = df_positive_cross[colonnes_positive_cross_excel]

        df_positive_cross.to_excel(writer, index=False, sheet_name='Relations Croisées Positives')

        # Ajuster la largeur des colonnes pour la feuille 'Relations Croisées Positives'
        for col_idx, column_name in enumerate(colonnes_positive_cross_excel):
            max_length = len(str(column_name))
            if not df_positive_cross.empty:
                max_length = max(max_length, longueur_max_colonne(df_positive_cross[column_name]))

            adjusted_width = (max_length + 2)
            if column_name == "Commentaire":
                adjusted_width = min(adjusted_width, 100)
            elif column_name in ["Date", "Début", "Fin", "Service", "Vigilance", "Type de Croisé"]:
                 adjusted_width = min(adjusted_width, 25)
            else:
                 adjusted_width = min(adjusted_width, 15)

            worksheet_positive_cross.column_dimensions[openpyxl.utils.get_column_letter(col_idx + 1)].width = adjusted_width


        # --- Nouvelle feuille : Récap ---
        worksheet_recap = workbook.create_sheet(title='Récap')

        recap_dataframes = []

        if not df_unidirectional.empty:
            df_temp_uni = df_unidirectional.copy()
            df_temp_uni['Type de Croisé'] = None  # Pas de type de croisé pour les unidirectionnelles
            df_temp_uni['Type de Récap'] = 'Unidirectionnelle'
            recap_dataframes.append(df_temp_uni)
        
        if not df_negative_cross.empty:
            df_temp_neg = df_negative_cross.copy()
            df_temp_neg['Type de Récap'] = 'Négative Croisée'
            recap_dataframes.append(df_temp_neg)

        if not df_positive_cross.empty:
            df_temp_pos = df_positive_cross.copy()
            df_temp_pos['Type de Récap'] = 'Positive Croisée'
            recap_dataframes.append(df_temp_pos)
        
        colonnes_recap_excel = colonnes_relations_excel + ["Type de Croisé", "Type de Récap"]
        df_recap = pd.DataFrame(columns=colonnes_recap_excel)
        if recap_dataframes:
            df_recap = pd.concat(recap_dataframes, ignore_index=True)
            
            # S'assurer que toutes les colonnes définies existent dans le DataFrame récapitulatif
            for col in colonnes_recap_excel:
                if col not in df_recap.columns:
                    df_recap[col] = None
            
            # Réorganiser le DataFrame selon l'ordre des colonnes définies
            df_recap = df_recap[colonnes_recap_excel]


        df_recap.to_excel(writer, index=False, sheet_name='Récap')

        # Ajuster la largeur des colonnes pour la feuille 'Récap'
        for col_idx, column_name in enumerate(colonnes_recap_excel):
            max_length = len(str(column_name))
            if not df_recap.empty:
                max_length = max(max_length, longueur_max_colonne(df_recap[column_name]))

            adjusted_width = (max_length + 2)
            if column_name == "Commentaire":
                adjusted_width = min(adjusted_width, 100)
            elif column_name in ["Date", "Début", "Fin", "Service", "Vigilance", "Type de Croisé", "Type de Récap"]:
                 adjusted_width = min(adjusted_width, 25)
            else:
                 adjusted_width = min(adjusted_width, 15)

            worksheet_recap.column_dimensions[openpyxl.utils.get_column_letter(col_idx + 1)].width = adjusted_width


    output.seek(0)
    return output.getvalue()


def exporter_zip():
    """Crée un fichier ZIP contenant le JSON et l'Excel des données du projet."""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        # Ajouter le fichier JSON
        json_data = exporter_json_data()
        zf.writestr("barometre_projet.json", json_data)

        # Ajouter le fichier Excel
        excel_data = exporter_excel_data()
        zf.writestr("relations_barometre.xlsx", excel_data)

    zip_buffer.seek(0)
    return zip_buffer


def importer_json():
    st.markdown("### 📁 Glissez-déposez un fichier JSON ici ou cliquez pour le sélectionner :")
    fichier = st.file_uploader("Charger un fichier JSON", type=["json"], label_visibility="collapsed")
    if fichier is not None:
        try:
            contenu = json.load(fichier)
            projet = valider_donnees_projet(contenu)
            st.session_state.participants = projet["participants"]
            st.session_state.services = projet["services"]
            st.session_state.relations_saisies = projet["relations_saisies"]
            st.session_state.nombre_total_personnes = projet["nombre_total_personnes"]
            st.success("Projet chargé avec succès !")
            st.session_state.etat = "relations"
            st.rerun()
        except Exception as e:
            st.error(f"Erreur lors du chargement du fichier JSON : {e}")


# === TITRE PRINCIPAL =========================================================
st.title("Baromètre Relationnel — Réalisé par Hatice Gultekin")

# === MENU PRINCIPAL ==========================================================
if st.session_state.etat == "menu":
    st.subheader("Gestion de projet")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Démarrer un nouveau projet"):
            st.session_state.participants         = []
            st.session_state.services             = []
            st.session_state.relations_saisies = []
            st.session_state.relation_a_modifier = None
            st.session_state.selected_relations = pd.DataFrame()
            st.session_state.nombre_total_personnes = 0
            st.session_state.etat = "participants"
            st.rerun()
    # Appel direct de importer_json() en dehors de la colonne pour optimiser le glisser-déposer
    importer_json()

# === ETAPE 1 : PARTICIPANTS ==================================================
elif st.session_state.etat == "participants":
    st.subheader("Étape 1 : Saisie des participants")

    # Formulaire d’ajout
    with st.form("form_ajout_participant"):
        col1, col2 = st.columns(2)
        nom = col1.text_input("Nom du participant")
        nouveau_service = col2.text_input("Service associé")
        ajouter = st.form_submit_button("Ajouter")

    if ajouter:
        st.session_state.participant_ajoute_message = None
        if nom.strip() and nouveau_service.strip():
            if all(nom != p["nom"] for p in st.session_state.participants):
                if nouveau_service not in st.session_state.services:
                    st.session_state.services.append(nouveau_service)
                st.session_state.participants.append(
                    {"nom": nom.strip(), "service": nouveau_service.strip()}
                )
                st.session_state.participant_ajoute_message = (
                    f"Participant « {nom.strip()} » ajouté."
                )
                st.rerun()
            else:
                st.warning("Participant déjà ajouté.")
        else:
            st.warning("Veuillez saisir un nom ET un service valides.")

    if st.session_state.participant_ajoute_message:
        st.info(st.session_state.participant_ajoute_message)
        st.session_state.participant_ajoute_message = None

    if not st.session_state.participants:
        st.info("Ajoutez au moins deux participants pour continuer.")

    # Sélecteur + bouton de modification
    if st.session_state.participants:
        noms_services = [f"{p['nom']} ({p['service']})"
                                 for p in st.session_state.participants]
        index_to_modify = st.selectbox(
            "Modifier un participant existant",
            options=noms_services,
            index=0 if noms_services else None
        )

        if index_to_modify and st.button("✏️ Modifier le participant", key="modifier_participant_menu"):
            st.session_state.participant_a_modifier = index_to_modify.split(" (")[0]
            st.rerun()

    if st.session_state.participants and st.button(
        "🗑️ Supprimer le participant", key="supprimer_participant_menu"
    ):
        if index_to_modify:
            participant_nom = index_to_modify.split(" (")[0]
            st.session_state.participants = [
                p for p in st.session_state.participants if p["nom"] != participant_nom
            ]
            st.session_state.relations_saisies = [
                r for r in st.session_state.relations_saisies
                if r["Émetteur"] != participant_nom and r["Récepteur"] != participant_nom
            ]
            st.success(f"Participant « {participant_nom} » et ses relations ont été supprimés.")
            st.rerun()
        else:
            st.warning("Veuillez sélectionner un participant à supprimer.")


    # Formulaire de modification (étape 1)
    if st.session_state.participant_a_modifier:
        data = next((p for p in st.session_state.participants
                                 if p["nom"] == st.session_state.participant_a_modifier), None)
        if data:
            with st.form("modif_form_etape1"):
                new_nom = st.text_input("Nouveau nom", value=data["nom"])
                new_service = st.text_input("Nouveau service", value=data["service"])
                submit_modif = st.form_submit_button("Valider les modifications")

            if submit_modif:
                for rel in st.session_state.relations_saisies:
                    if rel["Émetteur"] == data["nom"]:
                        rel["Émetteur"] = new_nom.strip()
                    if rel["Récepteur"] == data["nom"]:
                        rel["Récepteur"] = new_nom.strip()
                data["nom"] = new_nom.strip()
                data["service"] = new_service.strip()
                st.success("Participant modifié avec succès.")
                st.session_state.participant_a_modifier = None
                st.rerun()

    # Bouton suivant
    if len(st.session_state.participants) >= 2:
        if st.button("Passer à l'étape suivante"):
            st.session_state.etat = "relations"
            st.rerun()
    elif st.session_state.participants:
        st.info("Ajoutez au moins un autre participant pour continuer.")

# === ETAPE 2 : RELATIONS =====================================================
elif st.session_state.etat == "relations":
    st.subheader("Étape 2 : Saisie des relations")

    # Le total suit automatiquement les participants saisis à l'étape 1.
    st.session_state.nombre_total_personnes = len(st.session_state.participants)
    st.markdown("---")
    st.markdown("### Nombre total de personnes concernées par l'analyse")
    st.write(f"{st.session_state.nombre_total_personnes} participant(s) enregistré(s) à l’étape 1")
    st.markdown("---")


    # Sélecteur + bouton de modification (étape 2)
    if st.session_state.participants:
        noms_services = [f"{p['nom']} ({p['service']})"
                                 for p in st.session_state.participants]
        index_to_modify_rel = st.selectbox(
            "Modifier un participant existant",
            options=noms_services,
            key="select_modifier_rel",
            index=0 if noms_services else None
        )
        if index_to_modify_rel and st.button("✏️ Modifier le participant", key="modifier_participant_relations"):
            st.session_state.participant_a_modifier = index_to_modify_rel.split(" (")[0]
            st.rerun()

        if st.button("🗑️ Supprimer le participant", key="supprimer_participant_menu"):
            if index_to_modify_rel:
                participant_nom = index_to_modify_rel.split(" (")[0]
                st.session_state.participants = [
                    p for p in st.session_state.participants if p["nom"] != participant_nom
                ]
                st.session_state.relations_saisies = [
                    r for r in st.session_state.relations_saisies
                    if r["Émetteur"] != participant_nom and r["Récepteur"] != participant_nom
                ]
                st.success(f"Participant « {participant_nom} » et ses relations ont été supprimés.")
                st.rerun()
            else:
                st.warning("Veuillez sélectionner un participant à supprimer.")


    # Formulaire de modification (affiché dans l’étape 2)
    if st.session_state.participant_a_modifier:
        data = next((p for p in st.session_state.participants
                                 if p["nom"] == st.session_state.participant_a_modifier), None)
        if data:
            with st.form("modif_form_relations"):
                new_nom = st.text_input("Nouveau nom", value=data["nom"])
                new_service = st.text_input("Nouveau service", value=data["service"])
                submit_modif = st.form_submit_button("Valider les modifications")

            if submit_modif:
                for rel in st.session_state.relations_saisies:
                    if rel["Émetteur"] == data["nom"]:
                        rel["Émetteur"] = new_nom.strip()
                    if rel["Récepteur"] == data["nom"]:
                        rel["Récepteur"] = new_nom.strip()
                data["nom"] = new_nom.strip()
                data["service"] = new_service.strip()
                st.success("Participant modifié avec succès.")
                st.session_state.participant_a_modifier = None
                st.rerun()

    # --- Ajout rapide d’un participant --------------------------------------
    with st.expander("➕ Ajouter un participant oublié"):
        with st.form("form_ajout_participant_rapide"):
            col1, col2 = st.columns(2)
            nom_rapide = col1.text_input("Nom du participant oublié")
            service_rapide = col2.text_input("Service associé")
            ajouter_rapide = st.form_submit_button("Ajouter")

        if ajouter_rapide:
            if nom_rapide.strip() and service_rapide.strip():
                if all(nom_rapide != p["nom"]
                                 for p in st.session_state.participants):
                    if service_rapide not in st.session_state.services:
                        st.session_state.services.append(service_rapide)
                    st.session_state.participants.append(
                        {"nom": nom_rapide.strip(),
                         "service": service_rapide.strip()}
                    )
                    st.success("Participant ajouté avec succès.")
                    st.rerun()
                else:
                    st.warning("Ce participant existe déjà.")
            else:
                st.warning("Veuillez saisir un nom ET un service valides.")

    # --- Liste des relations possibles --------------------------------------
    noms = [p["nom"] for p in st.session_state.participants]
    relations_possibles = [(e, r) for e in noms for r in noms if e != r]
    numeros_relations = {relation: index + 1 for index, relation in enumerate(relations_possibles)}
    recherche_relation = st.text_input(
        "Rechercher une personne dans les relations :",
        placeholder="Exemple : Bob",
        key="recherche_relation_input",
    ).strip().casefold()
    relations_filtrees = [
        relation for relation in relations_possibles
        if not recherche_relation or any(
            recherche_relation in nom.casefold() for nom in relation
        )
    ]
    relation_textes = [
        f"{numeros_relations[relation]}. {relation[0]} → {relation[1]}"
        for relation in relations_filtrees
    ]

    if relations_possibles:
        if relation_textes:
            current_selection_index = 0
            selected_relation_index = st.session_state.get("relation_choisie_index")
            if selected_relation_index is not None:
                for index, relation in enumerate(relations_filtrees):
                    if numeros_relations[relation] - 1 == selected_relation_index:
                        current_selection_index = index
                        break

            if st.session_state.get("relation_choisie_select") not in relation_textes:
                st.session_state.pop("relation_choisie_select", None)

            relation_choisie = st.selectbox(
                "Relation à enregistrer :",
                relation_textes,
                index=current_selection_index,
                key="relation_choisie_select"
            )

            relation_par_texte = dict(zip(relation_textes, relations_filtrees))
            relation_selectionnee = relation_par_texte[relation_choisie]
            st.session_state.relation_choisie_index = numeros_relations[relation_selectionnee] - 1
        else:
            relation_choisie = None
            st.info("Aucune relation ne correspond à cette personne.")


        if relation_choisie:
            try:
                emetteur, recepteur = relation_selectionnee
                service_emetteur = next(
                    (p["service"] for p in st.session_state.participants
                     if p["nom"] == emetteur), ""
                )

                # Formulaire d’enregistrement
                date_default = datetime.now().date()
                date  = st.date_input("Date", value=date_default)
                debut = st.text_input("Heure début", key="heure_debut_input")
                fin   = st.text_input("Heure fin", key="heure_fin_input")
                st.markdown(f"**Service détecté automatiquement :** {service_emetteur}")

                indicateurs = ["P+", "P-", "I+", "I-", "C+", "C-"]
                indicateurs_values = {i: st.checkbox(i, key=f"indic_{i}_checkbox") for i in indicateurs}
                commentaire = st.text_area("Commentaire global", key="commentaire_input")

                if st.button("💾 Enregistrer la relation"):
                    erreurs = []
                    if not debut.strip():
                        erreurs.append("Heure de début manquante")
                    if not fin.strip():
                        erreurs.append("Heure de fin manquante")
                    if not any(indicateurs_values.values()):
                        erreurs.append("Aucun indicateur sélectionné")

                    is_duplicate = False
                    for rel in st.session_state.relations_saisies:
                        if (rel["Émetteur"] == emetteur and
                            rel["Récepteur"] == recepteur and
                            rel["Date"] == date.strftime("%d/%m/%Y") and
                            rel["Début"] == debut and
                            rel["Fin"] == fin):
                            is_duplicate = True
                            break
                    if is_duplicate:
                        erreurs.append("Une relation avec le même émetteur, récepteur, date, heure de début et heure de fin existe déjà.")


                    if erreurs:
                        for e in erreurs:
                            st.warning(e)
                    else:
                        p_plus  = sum(indicateurs_values[i]
                                      for i in ["P+", "I+", "C+"])
                        p_moins = sum(indicateurs_values[i]
                                      for i in ["P-", "I-", "C-"])
                        vigilance = classer_relation(p_plus, p_moins)

                        st.session_state.relations_saisies.append({
                            "Émetteur": emetteur,
                            "Récepteur": recepteur,
                            "Date": date.strftime("%d/%m/%Y"),
                            "Début": debut,
                            "Fin": fin,
                            "Service": service_emetteur,
                            "P+": int(indicateurs_values["P+"]),
                            "P-": int(indicateurs_values["P-"]),
                            "I+": int(indicateurs_values["I+"]),
                            "I-": int(indicateurs_values["I-"]),
                            "C+": int(indicateurs_values["C+"]),
                            "C-": int(indicateurs_values["C-"]),
                            "Score Pic Positif": p_plus,
                            "Score Pic Négatif": p_moins,
                            "Score Net": p_plus - p_moins,
                            "Vigilance": vigilance,
                            "Commentaire": commentaire,
                        })
                        st.success("Relation enregistrée.")
                        st.rerun()
            except Exception as e:
                st.warning(f"Erreur lors de l'extraction des données de relation. Veuillez vérifier la sélection: {e}")
        else:
            st.warning("Aucune relation possible sélectionnée. Veuillez en ajouter.")
    else:
        st.warning("Aucune relation possible ; ajoutez plus de participants.")


    # --- Affichage des Relations dans AgGrid ---
    st.markdown("---")
    st.subheader("Relations enregistrées")
    if st.session_state.relations_saisies:
        colonnes_ordonnees = [
            "N° relation", "Émetteur", "Récepteur", "Date", "Début", "Fin", "Service",
            "P+", "P-", "I+", "I-", "C+", "C-",
            "Score Pic Positif", "Score Pic Négatif", "Score Net",
            "Vigilance", "Commentaire"
        ]
        df = pd.DataFrame(st.session_state.relations_saisies)
        df.insert(0, "N° relation", [
            numeros_relations.get((emetteur, recepteur), "?")
            for emetteur, recepteur in zip(df["Émetteur"], df["Récepteur"])
        ])

        for col in colonnes_ordonnees:
            if col not in df.columns:
                df[col] = None

        df = df[colonnes_ordonnees]

        gb = GridOptionsBuilder.from_dataframe(df)
        gb.configure_selection(selection_mode="multiple", use_checkbox=True)

        # Configuration des colonnes pour AgGrid
        gb.configure_column("N° relation", header_name="N°", width=65, minWidth=65, maxWidth=75)
        gb.configure_column("Émetteur", header_name="Émetteur", wrapText=True, autoHeight=True, minWidth=100)
        gb.configure_column("Récepteur", header_name="Récepteur", wrapText=True, autoHeight=True, minWidth=100)
        gb.configure_column("Date", header_name="Date", wrapText=True, autoHeight=True, minWidth=80)
        gb.configure_column("Début", header_name="Début", wrapText=True, autoHeight=True, minWidth=70)
        gb.configure_column("Fin", header_name="Fin", wrapText=True, autoHeight=True, minWidth=70)
        gb.configure_column("Service", header_name="Service", wrapText=True, autoHeight=True, minWidth=100)

        for col_name in ["P+", "P-", "I+", "I-", "C+", "C-", "Score Pic Positif", "Score Pic Négatif", "Score Net"]:
            gb.configure_column(col_name, header_name=col_name, type=["numericColumn", "numberColumnFilter", "customNumericFormat"], precision=0, minWidth=50)

        gb.configure_column("Vigilance", header_name="Vigilance", wrapText=True, autoHeight=True, minWidth=120)

        # Configuration spécifique pour la colonne "Commentaire"
        gb.configure_column("Commentaire", header_name="Commentaire", wrapText=True, autoHeight=True, minWidth=200, flex=1)

        # Configuration de la grille globale
        gb.configure_grid_options(domLayout='normal')

        grid_options = gb.build()

        grid_response = AgGrid(
            df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            data_return_mode=DataReturnMode.AS_INPUT,
            height=400,
            allow_unsafe_jscode=True,
            enable_enterprise_modules=False,
            key="grid_relations_display",
            columns_auto_size_mode=ColumnsAutoSizeMode.FIT_CONTENTS,
            js_code="""
            function(params) {
                params.api.resetRowHeights();
                params.api.autoSizeAllColumns();
            }
            """
        )

        selected_rows = grid_response.get("selected_rows")
        if isinstance(selected_rows, pd.DataFrame):
            st.session_state.selected_relations = selected_rows
        elif isinstance(selected_rows, list):
            st.session_state.selected_relations = pd.DataFrame(selected_rows)
        else:
            st.session_state.selected_relations = pd.DataFrame()
        if st.session_state.selected_relations.empty:
            st.session_state.selected_relations = pd.DataFrame()

        if st.button("🗑️ Supprimer les relations sélectionnées", key="delete_selected_relations_button"):
            if not st.session_state.selected_relations.empty:
                selected_ids = set()
                for s_row in st.session_state.selected_relations.to_dict(orient='records'):
                    # Crée un hash basé sur les champs identifiants pour trouver les doublons uniques
                    row_hash = hashlib.md5(json.dumps({k: s_row[k] for k in ["Émetteur", "Récepteur", "Date", "Début", "Fin"]}, sort_keys=True).encode('utf-8')).hexdigest()
                    selected_ids.add(row_hash)

                new_relations_saisies = []
                for r in st.session_state.relations_saisies:
                    current_row_hash = hashlib.md5(json.dumps({k: r[k] for k in ["Émetteur", "Récepteur", "Date", "Début", "Fin"]}, sort_keys=True).encode('utf-8')).hexdigest()
                    if current_row_hash not in selected_ids:
                        new_relations_saisies.append(r)
                
                st.session_state.relations_saisies = new_relations_saisies
                st.success("Relations sélectionnées supprimées.")
                st.rerun()
            else:
                st.warning("Veuillez sélectionner des relations à supprimer.")

    else:
        st.info("Aucune relation enregistrée pour l'instant. Saisissez de nouvelles relations ci-dessus.")

    # --- Boutons de navigation ---
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Retour au menu principal"):
            st.session_state.etat = "menu"
            st.rerun()
    with col2:
        # Bouton d'exportation pour l'ensemble du projet au format ZIP
        download_zip_filename = f"barometre_projet_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        st.download_button(
            label="📦 Télécharger le projet (JSON + Excel)",
            data=exporter_zip(),
            file_name=download_zip_filename,
            mime="application/zip",
            key="download_project_zip_button"
        )
