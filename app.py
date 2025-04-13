import streamlit as st
import pandas as pd
import json
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import io
import os
import sys
import uuid
import glob

# Configuration de base de l'interface Streamlit
st.set_page_config(page_title="Atterrissage VL", page_icon="📊", layout="wide")

# Système d'authentification simple avec persistance de 30 jours
def check_password():
    """Retourne `True` si le mot de passe est correct ou déjà validé, `False` sinon."""
    # Vérifier si l'authentification a déjà été validée dans un cookie
    if "password_validated" in st.session_state:
        expiry_date = st.session_state["password_validated"]
        current_date = datetime.now()
        # Si le cookie est toujours valide (moins de 30 jours)
        if (current_date - expiry_date).days < 30:
            return True

    def password_entered():
        """Vérifie si le mot de passe entré par l'utilisateur est correct."""
        if st.session_state["password"] == "VL2025":
            st.session_state["password_correct"] = True
            # Enregistrer la date de validation pour 30 jours
            st.session_state["password_validated"] = datetime.now()
            del st.session_state["password"]  # Ne pas stocker le mot de passe
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # Premier affichage, afficher l'entrée du mot de passe
        st.text_input(
            "Mot de passe", type="password", on_change=password_entered, key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        # Mot de passe incorrect, nouvelle tentative
        st.text_input(
            "Mot de passe", type="password", on_change=password_entered, key="password"
        )
        st.error("😕 Mot de passe incorrect")
        return False
    else:
        # Mot de passe correct
        return True

# Vérifie si l'utilisateur est authentifié
if not check_password():
    st.stop()  # Arrête l'exécution si le mot de passe est incorrect

# Style CSS custom
st.markdown("""
<style>
    .main .block-container {padding-top: 2rem;}
    .stButton>button {background-color: #0000DC; color: white; font-weight: bold;}
    .stButton>button:hover {background-color: #00008C;}
    h1, h2, h3 {color: #0000DC;}
    .stTabs [data-baseweb="tab-list"] {gap: 20px;}
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f0f2f6;
        border-radius: 4px 4px 0px 0px;
        gap: 1px;
        padding-top: 10px;
        padding-bottom: 10px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0000DC;
        color: white;
    }
    .big-button {
        font-size: 1.2rem !important;
        padding: 0.8rem !important;
        width: 100%;
    }
    .simulation-card {
        background-color: #f8f9fa;
        border-radius: 5px;
        padding: 15px;
        margin-bottom: 10px;
        border-left: 5px solid #0000DC;
    }
</style>
""", unsafe_allow_html=True)

# === FONCTION D'UTILITAIRES ===
def format_fr_euro(valeur):
    """Formater un nombre en euros format français"""
    try:
        valeur_arrondie = round(float(valeur), 2)
        return f"{valeur_arrondie:,.2f} €".replace(",", " ").replace(".", ",")
    except (ValueError, TypeError):
        return "0,00 €"

def champ_numerique(label, valeur, conteneur=st.sidebar):
    """Gérer la saisie d'une valeur numérique au format français"""
    try:
        champ = conteneur.text_input(label, value=format_fr_euro(valeur))
        try:
            champ = champ.replace(" ", "").replace(",", ".").replace("€", "")
            return float(champ) if champ else 0.0
        except ValueError:
            st.warning(f"Valeur non numérique pour {label}, utilisation de 0.0")
            return 0.0
    except Exception as e:
        st.warning(f"Erreur avec le champ {label}: {str(e)}")
        return 0.0

# === INITIALISATION DU STOCKAGE JSON ===
def init_storage():
    """Créer le répertoire de stockage des fichiers JSON si nécessaire"""
    try:
        # Créer le répertoire data s'il n'existe pas
        os.makedirs('data/simulations', exist_ok=True)
        
    except Exception as e:
        # En cas d'erreur, informer clairement l'utilisateur
        st.error(f"Erreur lors de l'initialisation du stockage: {str(e)}")
        import traceback
        st.error(traceback.format_exc())

# === PARAMÈTRES INITIAUX ===
default_params = {
    "nom_fonds": "Nom du Fonds",
    "nom_scenario": "Base case",
    "date_vl_connue": "31/12/2024",
    "date_fin_fonds": "31/12/2027",
    "anr_derniere_vl": 10_000_000.0,
    "nombre_parts": 10_000.0,
    "impacts": [
        ("Frais corporate", -50_000.0)
    ],
    "impacts_multidates": [
        {
            "libelle": "Honoraires NIV",
            "montants": [
                {"date": "30/06/2025", "montant": -30_000.0},
                {"date": "31/12/2025", "montant": -15_000.0}
            ]
        }
    ],
    "actifs": [
        {
            "nom": "Actif 1",
            "pct_detention": 1.0,
            "valeur_actuelle": 5_000_000.0,
            "valeur_projetee": 5_250_000.0,
            "is_a_provisionner": True,
            "variation": 187_500.0,
            "variation_brute": 250_000.0
        }
    ]
}

# Fonctions pour la gestion des simulations en JSON
def sauvegarder_simulation(params, commentaire=""):
    """Sauvegarder une simulation dans un fichier JSON"""
    try:
        # S'assurer que les valeurs numériques sont bien des nombres
        try:
            anr = float(params['anr_derniere_vl'])
            parts = float(params['nombre_parts'])
        except (ValueError, KeyError):
            # Si conversion impossible, utiliser des valeurs par défaut
            st.warning("Problème avec les valeurs numériques, utilisation de valeurs par défaut")
            anr = 10000000.0
            parts = 10000.0
        
        # Récupérer le nom du scénario, utiliser "Base case" par défaut
        nom_scenario = params.get('nom_scenario', 'Base case')
        nom_fonds = params.get('nom_fonds', 'Fonds sans nom')
        
        # Créer un identifiant unique pour cette simulation
        simulation_id = str(uuid.uuid4())
        
        # Préparer le dictionnaire complet de la simulation
        simulation_data = {
            "id": simulation_id,
            "nom_fonds": nom_fonds,
            "nom_scenario": nom_scenario,
            "date_vl_connue": params.get('date_vl_connue', '31/12/2023'),
            "date_fin_fonds": params.get('date_fin_fonds', '31/12/2026'),
            "anr_derniere_vl": anr,
            "nombre_parts": parts,
            "date_creation": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "commentaire": commentaire,
            "impacts": [],
            "impacts_multidates": [],
            "actifs": []
        }
        
        # Traiter les impacts récurrents
        impacts = params.get('impacts', [])
        for impact in impacts:
            try:
                if isinstance(impact, tuple) and len(impact) == 2:
                    libelle, montant = impact
                elif isinstance(impact, list) and len(impact) == 2:
                    libelle, montant = impact
                elif isinstance(impact, dict) and 'libelle' in impact and 'montant' in impact:
                    libelle, montant = impact['libelle'], impact['montant']
                else:
                    continue  # Ignorer les impacts mal formatés
                
                montant_float = float(montant)
                simulation_data['impacts'].append({
                    "libelle": libelle,
                    "montant": montant_float
                })
            except (ValueError, TypeError) as e:
                st.warning(f"Problème avec un impact récurrent: {str(e)}")
        
        # Traiter les impacts multidates
        impacts_multidates = params.get('impacts_multidates', [])
        for impact in impacts_multidates:
            try:
                impact_dict = {
                    "libelle": impact.get('libelle', 'Impact sans nom'),
                    "montants": []
                }
                
                # Ajouter les occurrences de cet impact
                for occurrence in impact.get('montants', []):
                    try:
                        montant_float = float(occurrence.get('montant', 0))
                        date = occurrence.get('date', '01/01/2024')
                        impact_dict["montants"].append({
                            "date": date,
                            "montant": montant_float
                        })
                    except (ValueError, TypeError) as e:
                        st.warning(f"Problème avec une occurrence d'impact: {str(e)}")
                
                simulation_data['impacts_multidates'].append(impact_dict)
            except Exception as e:
                st.warning(f"Problème avec un impact multidate: {str(e)}")
        
        # Traiter les actifs
        actifs = params.get('actifs', [])
        for actif in actifs:
            try:
                pct = float(actif.get('pct_detention', 1.0))
                val_act = float(actif.get('valeur_actuelle', 1000000.0))
                val_proj = float(actif.get('valeur_projetee', 1050000.0))
                is_prov = bool(actif.get('is_a_provisionner', False))
                
                # Calculer les valeurs dérivées
                variation_brute = (val_proj - val_act) * pct
                
                # Appliquer la règle de l'IS (75% de l'impact en cas de plus-value)
                if is_prov and variation_brute > 0:
                    variation = variation_brute * 0.75  # Réduction de 25% pour l'IS
                else:
                    variation = variation_brute
                
                simulation_data['actifs'].append({
                    "nom": actif.get('nom', 'Actif sans nom'),
                    "pct_detention": pct,
                    "valeur_actuelle": val_act,
                    "valeur_projetee": val_proj,
                    "is_a_provisionner": is_prov,
                    "variation": variation,
                    "variation_brute": variation_brute
                })
            except (ValueError, TypeError, KeyError) as e:
                st.warning(f"Problème avec un actif: {str(e)}")
        
        # Sauvegarder les données dans un fichier JSON
        filename = f"data/simulations/{simulation_id}.json"
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(simulation_data, f, ensure_ascii=False, indent=4)
        
        return simulation_id
        
    except Exception as e:
        st.error(f"Erreur lors de la sauvegarde: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None

def charger_simulation(simulation_id):
    """Charger une simulation depuis un fichier JSON"""
    try:
        # Charger les données depuis le fichier JSON
        filename = f"data/simulations/{simulation_id}.json"
        with open(filename, 'r', encoding='utf-8') as f:
            simulation_data = json.load(f)
        
        # Structure pour stocker les paramètres nécessaires à l'application
        params = {
            'nom_fonds': simulation_data.get('nom_fonds', 'Fonds sans nom'),
            'nom_scenario': simulation_data.get('nom_scenario', 'Base case'),
            'date_vl_connue': simulation_data.get('date_vl_connue', '31/12/2023'),
            'date_fin_fonds': simulation_data.get('date_fin_fonds', '31/12/2026'),
            'anr_derniere_vl': float(simulation_data.get('anr_derniere_vl', 10000000.0)),
            'nombre_parts': float(simulation_data.get('nombre_parts', 10000.0)),
            'impacts': [],
            'impacts_multidates': [],
            'actifs': []
        }
        
        # Récupérer les impacts récurrents
        for impact in simulation_data.get('impacts', []):
            libelle = impact.get('libelle', 'Sans nom')
            montant = float(impact.get('montant', 0.0))
            params['impacts'].append((libelle, montant))
        
        # Récupérer les impacts multidates
        params['impacts_multidates'] = simulation_data.get('impacts_multidates', [])
        
        # Récupérer les actifs
        params['actifs'] = simulation_data.get('actifs', [])
        
        return params
    except Exception as e:
        st.error(f"Erreur lors du chargement: {str(e)}")
        return None

def lister_simulations():
    """Lister toutes les simulations sauvegardées"""
    try:
        simulations = []
        # Trouver tous les fichiers JSON dans le répertoire des simulations
        json_files = glob.glob('data/simulations/*.json')
        
        for file_path in json_files:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    simulation_data = json.load(f)
                
                # Créer un dictionnaire avec les informations importantes
                sim_dict = {
                    'id': simulation_data.get('id', os.path.basename(file_path).replace('.json', '')),
                    'nom_fonds': simulation_data.get('nom_fonds', 'Fonds sans nom'),
                    'nom_scenario': simulation_data.get('nom_scenario', 'Base case'),
                    'date_vl_connue': simulation_data.get('date_vl_connue', '31/12/2023'),
                    'date_creation': simulation_data.get('date_creation', ''),
                    'commentaire': simulation_data.get('commentaire', '')
                }
                
                # Nettoyer le nom du fonds (enlever les dates potentielles)
                if sim_dict['nom_fonds'] and '(' in sim_dict['nom_fonds']:
                    # Si le nom contient une parenthèse (comme une date), prendre juste la partie avant
                    sim_dict['nom_fonds'] = sim_dict['nom_fonds'].split('(')[0].strip()
                
                simulations.append(sim_dict)
            except Exception as e:
                st.warning(f"Problème lors de la lecture du fichier {file_path}: {str(e)}")
        
        # Trier par date de création (du plus récent au plus ancien)
        simulations.sort(key=lambda x: x.get('date_creation', ''), reverse=True)
        
        return simulations
    except Exception as e:
        st.error(f"Erreur lors de la lecture des simulations: {str(e)}")
        return []

def supprimer_simulation(simulation_id):
    """Supprimer une simulation (fichier JSON)"""
    try:
        filename = f"data/simulations/{simulation_id}.json"
        if os.path.exists(filename):
            os.remove(filename)
            return True
        else:
            st.warning(f"Simulation avec ID {simulation_id} introuvable.")
            return False
    except Exception as e:
        st.error(f"Erreur lors de la suppression: {str(e)}")
        return False

# Initialiser le stockage au démarrage
try:
    init_storage()
except Exception as e:
    st.error(f"Erreur critique lors de l'initialisation: {str(e)}")

# === TITRE ET LAYOUT PRINCIPAL ===
st.title("Atterrissage VL")

# Afficher la barre latérale avec les simulations chargées
st.sidebar.title("Simulations sauvegardées")
simulations = lister_simulations()
if simulations:
    st.sidebar.markdown("### Charger une simulation")
    for sim in simulations:
        if st.sidebar.button(f"📂 {sim['nom_fonds']} - {sim['nom_scenario']}", key=f"sidebar_load_{sim['id']}"):
            params_charges = charger_simulation(sim['id'])
            if params_charges:
                st.session_state.params = params_charges
                st.sidebar.success(f"Simulation '{sim['nom_scenario']}' chargée avec succès")
                st.rerun()
    st.sidebar.markdown("---")
else:
    st.sidebar.info("Aucune simulation sauvegardée")
    st.sidebar.markdown("---")

# Interface principale avec onglets
tab1, tab2, tab3 = st.tabs(["📊 Projection VL", "💾 Gestion des simulations", "ℹ️ Aide"])

with tab1:
    # === INITIALISATION DES PARAMÈTRES DE SESSION ===
    if 'params' not in st.session_state:
        st.session_state.params = default_params.copy()
    
    params = st.session_state.params
    
    # === CRÉATION DE COLONNES POUR LE LAYOUT ===
    col_param, col_impacts = st.columns([1, 1])
    
    with col_param:
        st.subheader("Paramètres généraux")
        nom_fonds = st.text_input("Nom du fonds", params.get('nom_fonds', default_params['nom_fonds']))
        nom_scenario = st.text_input("Nom du scénario", params.get('nom_scenario', default_params['nom_scenario']))
        
        col1, col2 = st.columns(2)
        with col1:
            date_vl_connue_str = st.text_input("Date dernière VL connue (jj/mm/aaaa)", 
                                              params.get('date_vl_connue', default_params['date_vl_connue']))
        with col2:
            date_fin_fonds_str = st.text_input("Date fin de fonds (jj/mm/aaaa)", 
                                              params.get('date_fin_fonds', default_params['date_fin_fonds']))
        
        col3, col4 = st.columns(2)
        with col3:
            anr_derniere_vl = champ_numerique("ANR dernière VL connue (€)", 
                                              params.get('anr_derniere_vl', default_params['anr_derniere_vl']), st)
        with col4:
            # Utilisation d'un champ spécifique pour le nombre de parts (sans format euro)
            try:
                nombre_parts_str = st.text_input("Nombre de parts", 
                                              value=f"{params.get('nombre_parts', default_params['nombre_parts']):,.2f}".replace(",", " ").replace(".", ","))
                nombre_parts = float(nombre_parts_str.replace(" ", "").replace(",", "."))
            except ValueError:
                st.warning(f"Valeur non numérique pour le nombre de parts, utilisation de {default_params['nombre_parts']}")
                nombre_parts = default_params['nombre_parts']
    
    with col_impacts:
        st.subheader("Impacts semestriels récurrents")
        
        # Utilisation d'un expander pour montrer/cacher les impacts
        with st.expander("Gérer les impacts récurrents", expanded=True):
            # Nombre d'impacts
            nb_impacts = st.number_input("Nombre d'impacts récurrents", min_value=0, 
                                         value=len(params.get('impacts', [])), step=1)
            
            # Container pour les impacts
            impacts = []
            for i in range(nb_impacts):
                st.markdown(f"##### Impact {i+1}")
                col1, col2 = st.columns([2, 1])
                
                if i < len(params.get('impacts', [])):
                    try:
                        if isinstance(params['impacts'][i], tuple) and len(params['impacts'][i]) == 2:
                            libelle_defaut, montant_defaut = params['impacts'][i]
                        elif isinstance(params['impacts'][i], list) and len(params['impacts'][i]) == 2:
                            libelle_defaut, montant_defaut = params['impacts'][i]
                        elif isinstance(params['impacts'][i], dict) and 'libelle' in params['impacts'][i] and 'montant' in params['impacts'][i]:
                            libelle_defaut = params['impacts'][i]['libelle']
                            montant_defaut = params['impacts'][i]['montant']
                        else:
                            libelle_defaut = f"Impact {i+1}"
                            montant_defaut = 0.0
                    except (IndexError, TypeError, ValueError):
                        libelle_defaut = f"Impact {i+1}"
                        montant_defaut = 0.0
                else:
                    libelle_defaut = f"Impact {i+1}"
                    montant_defaut = 0.0
                    
                with col1:
                    libelle = st.text_input(f"Libellé impact {i+1}", libelle_defaut, key=f"imp_lib_{i}")
                with col2:
                    montant = champ_numerique(f"Montant semestriel (€)", montant_defaut, st)
                    
                impacts.append((libelle, montant))
                st.markdown("---")
    
    # === DATES POUR LA PROJECTION ===
    try:
        date_vl_connue = datetime.strptime(date_vl_connue_str, "%d/%m/%Y")
        date_fin_fonds = datetime.strptime(date_fin_fonds_str, "%d/%m/%Y")
        
        dates_semestres = [date_vl_connue]
        y = date_vl_connue.year
        while datetime(y, 12, 31) <= date_fin_fonds:
            if datetime(y, 6, 30) > date_vl_connue:
                dates_semestres.append(datetime(y, 6, 30))
            if datetime(y, 12, 31) > date_vl_connue:
                dates_semestres.append(datetime(y, 12, 31))
            y += 1
        
        # Liste des dates formatées pour le selectbox
        dates_semestres_str = [d.strftime("%d/%m/%Y") for d in dates_semestres]
    except ValueError as e:
        st.error(f"Erreur de format de date: {str(e)}")
        dates_semestres = [datetime.now()]
        dates_semestres_str = [datetime.now().strftime("%d/%m/%Y")]
    
    # === IMPACTS MULTIDATES ET ACTIFS ===
    st.subheader("Impacts multidates")
    with st.expander("Gérer les impacts à dates spécifiques", expanded=False):
        impacts_multidates = []
        nb_impacts_multidates = st.number_input("Nombre d'impacts multidates", min_value=0, 
                                            value=len(params.get('impacts_multidates', [])), step=1)
        
        for i in range(nb_impacts_multidates):
            st.markdown(f"##### Impact multidate {i+1}")
            
            # Récupération des paramètres par défaut s'ils existent
            if i < len(params.get('impacts_multidates', [])):
                impact_default = params['impacts_multidates'][i]
                libelle_defaut = impact_default.get('libelle', f"Impact multidate {i+1}")
                montants_defaut = impact_default.get('montants', [])
            else:
                libelle_defaut = f"Impact multidate {i+1}"
                montants_defaut = []
            
            # Libellé de l'impact multidate
            libelle = st.text_input(f"Libellé impact multidate {i+1}", libelle_defaut, key=f"multi_lib_{i}")
            
            # Nombre d'occurrences pour cet impact
            nb_occurrences = st.number_input(
                f"Nombre d'occurrences pour '{libelle}'", 
                min_value=0, 
                value=len(montants_defaut), 
                step=1,
                key=f"multi_nb_{i}"
            )
            
            # Liste pour stocker les montants et dates de cet impact
            montants = []
            
            # Interface pour chaque occurrence
            for j in range(nb_occurrences):
                col1, col2 = st.columns([1, 1])
                
                # Valeurs par défaut pour cette occurrence
                if j < len(montants_defaut):
                    montant_defaut = montants_defaut[j].get('montant', 0.0)
                    date_defaut = montants_defaut[j].get('date', dates_semestres_str[0])
                    try:
                        date_index = dates_semestres_str.index(date_defaut)
                    except ValueError:
                        date_index = 0
                else:
                    montant_defaut = 0.0
                    date_index = 0
                
                with col1:
                    # Selectbox pour la date
                    date_str = st.selectbox(
                        f"Date occurrence {j+1}",
                        options=dates_semestres_str,
                        index=date_index,
                        key=f"multi_date_{i}_{j}"
                    )
                
                with col2:
                    # Champ pour le montant
                    montant = champ_numerique(f"Montant (€)", montant_defaut, st)
                
                montants.append({
                    "date": date_str,
                    "montant": montant
                })
            
            impacts_multidates.append({
                "libelle": libelle,
                "montants": montants
            })
            
            st.markdown("---")
    
    # Actifs
    st.subheader("Actifs")
    with st.expander("Gérer les actifs du portefeuille", expanded=True):
        actifs = []
        nb_actifs = st.number_input("Nombre d'actifs", min_value=1, 
                                    value=max(1, len(params.get('actifs', []))), step=1)
        
        for i in range(nb_actifs):
            st.markdown(f"##### Actif {i+1}")
            
            if i < len(params.get('actifs', [])):
                a = params['actifs'][i]
                nom_defaut = a.get('nom', f"Actif {i+1}")
                pct_defaut = a.get('pct_detention', 1.0) * 100
                val_actuelle = a.get('valeur_actuelle', 1_000_000.0)
                val_proj = a.get('valeur_projetee', val_actuelle + 50_000)
                is_prov_defaut = a.get('is_a_provisionner', False)
            else:
                nom_defaut = f"Actif {i+1}"
                pct_defaut = 100.0
                val_actuelle = 1_000_000.0
                val_proj = 1_050_000.0
                is_prov_defaut = False
            
            col1, col2 = st.columns([2, 1])
            with col1:
                nom_actif = st.text_input(f"Nom de l'Actif", nom_defaut, key=f"actif_nom_{i}")
                
                col_pct, col_is = st.columns([2, 1])
                with col_pct:
                    # Champ de saisie directe du pourcentage au format français
                    pct_text = st.text_input(f"% Détention", value=f"{pct_defaut:.2f}".replace(".", ","), key=f"actif_pct_{i}")
                    try:
                        pct_detention = float(pct_text.replace(",", "."))
                    except ValueError:
                        st.warning(f"Valeur non numérique pour le pourcentage, utilisation de {pct_defaut}%")
                        pct_detention = pct_defaut
                
                with col_is:
                    # Option pour provisionner l'IS
                    is_a_provisionner = st.checkbox(f"IS à provisionner", value=is_prov_defaut, key=f"actif_is_{i}")
            
            with col2:
                valeur_actuelle = champ_numerique(f"Valeur actuelle (€)", val_actuelle, st)
                valeur_projetee = champ_numerique(f"Valeur projetée (€)", val_proj, st)
            
            # Calcul de la variation avec prise en compte de l'IS si nécessaire
            variation_brute = (valeur_projetee - valeur_actuelle) * (pct_detention / 100)
            
            # Appliquer la règle de l'IS (75% de l'impact en cas de plus-value)
            if is_a_provisionner and variation_brute > 0:
                variation = variation_brute * 0.75  # Réduction de 25% pour l'IS
            else:
                variation = variation_brute  # Pas de modification en cas de moins-value ou si IS non provisionné
            
            col_var1, col_var2 = st.columns(2)
            with col_var1:
                st.metric("Variation brute", format_fr_euro(variation_brute), delta=None)
            with col_var2:
                st.metric("Variation nette d'IS", format_fr_euro(variation), 
                         delta=f"-{format_fr_euro(variation_brute - variation)}" if variation != variation_brute else None)
            
            actifs.append({
                "nom": nom_actif,
                "pct_detention": pct_detention / 100,
                "valeur_actuelle": valeur_actuelle,
                "valeur_projetee": valeur_projetee,
                "variation": variation,
                "is_a_provisionner": is_a_provisionner,
                "variation_brute": variation_brute
            })
            
            st.markdown("---")
    
    # Boutons rapides pour sauvegarder et charger
    col_save1, col_save2 = st.columns(2)
    with col_save1:
        if st.button("💾 SAUVEGARDER CETTE SIMULATION", key="quick_save", help="Sauvegarde rapide de la simulation actuelle"):
            # Date au format jour/mois/année
            date_formatee = datetime.now().strftime("%d/%m/%Y")
            # Commentaire
            commentaire = f"{nom_scenario} - {date_formatee}"
            # Préparer les données à sauvegarder
            current_params = {
                "nom_fonds": nom_fonds,
                "nom_scenario": nom_scenario,
                "date_vl_connue": date_vl_connue_str,
                "date_fin_fonds": date_fin_fonds_str,
                "anr_derniere_vl": anr_derniere_vl,
                "nombre_parts": nombre_parts,
                "impacts": impacts,
                "impacts_multidates": impacts_multidates,
                "actifs": actifs
            }
            # Sauvegarder
            simulation_id = sauvegarder_simulation(current_params, commentaire)
            if simulation_id:
                st.success(f"Simulation '{nom_scenario}' sauvegardée avec succès")
            
    with col_save2:
        if st.button("🔄 ACTUALISER", key="refresh_calc", help="Recalculer la projection"):
            st.session_state.params = {
                "nom_fonds": nom_fonds,
                "nom_scenario": nom_scenario,
                "date_vl_connue": date_vl_connue_str,
                "date_fin_fonds": date_fin_fonds_str,
                "anr_derniere_vl": anr_derniere_vl,
                "nombre_parts": nombre_parts,
                "impacts": impacts,
                "impacts_multidates": impacts_multidates,
                "actifs": actifs
            }
            st.rerun()
            
    try:
        # === CALCUL PROJECTION DÉTAILLÉE ===
        vl_semestres = []
        anr_courant = anr_derniere_vl
        projection_rows = []
        
        for i, date in enumerate(dates_semestres):
            row = {"Date": date.strftime('%d/%m/%Y')}
            date_str = date.strftime('%d/%m/%Y')
        
            # Variation par actif (S+1 uniquement)
            total_var_actifs = 0
            for a in actifs:
                var = a.get('variation', 0) if i == 1 else 0
                
                # Si IS à provisionner et il s'agit d'une plus-value, afficher uniquement le montant net
                if i == 1 and a.get('is_a_provisionner', False) and a.get('variation_brute', 0) > 0:
                    # Afficher seulement le montant net, sans texte supplémentaire
                    row[f"Actif - {a.get('nom', 'Sans nom')}"] = format_fr_euro(var)
                else:
                    row[f"Actif - {a.get('nom', 'Sans nom')}"] = format_fr_euro(var)
                    
                total_var_actifs += var
        
            # Impacts récurrents
            total_impacts_recurrents = 0
            for libelle, montant in impacts:
                if i > 0:  # Appliquer les impacts récurrents à partir de S+1
                    row[f"Impact récurrent - {libelle}"] = format_fr_euro(montant)
                    total_impacts_recurrents += montant
                else:
                    row[f"Impact récurrent - {libelle}"] = format_fr_euro(0)
        
            # Impacts multidates pour cette date spécifique
            total_impacts_multidates = 0
            for impact in impacts_multidates:
                libelle = impact.get('libelle', 'Sans nom')
                # Initialiser la valeur à 0 par défaut
                impact_valeur = 0
                
                # Parcourir toutes les occurrences de cet impact
                for occurrence in impact.get('montants', []):
                    if occurrence.get('date') == date_str:
                        impact_valeur += occurrence.get('montant', 0)
                
                # Ajouter au total
                total_impacts_multidates += impact_valeur
                
                # Afficher dans le tableau
                row[f"Impact multidate - {libelle}"] = format_fr_euro(impact_valeur)
        
            if i > 0:
                anr_courant += total_var_actifs + total_impacts_recurrents
            # Appliquer les impacts multidates
            anr_courant += total_impacts_multidates
        
            vl = anr_courant / nombre_parts if nombre_parts else 0
            # Arrondir à deux décimales
            vl = round(vl, 2)
            vl_semestres.append(vl)
            row["VL prévisionnelle (€)"] = format_fr_euro(vl)
            row["ANR (€)"] = format_fr_euro(anr_courant)  # Ajout de l'ANR au tableau
        
            projection_rows.append(row)
        
        # === AFFICHAGE TABLEAU ===
        st.subheader("VL prévisionnelle")
        projection = pd.DataFrame(projection_rows)
        st.dataframe(projection, use_container_width=True)
        
        # === GRAPHIQUE BLEU STYLÉ ===
        st.subheader("Graphique d'évolution de la VL")
        
        couleur_bleue = "#0000DC"
        
        fig, ax = plt.subplots(figsize=(10, 5))
        
        # Courbe
        ax.plot(
            dates_semestres,
            vl_semestres,
            linewidth=2.5,
            marker='o',
            markersize=7,
            color=couleur_bleue,
            markerfacecolor=couleur_bleue,  # Points remplis de couleur bleue
            markeredgewidth=1,
            markeredgecolor=couleur_bleue
        )
        
        # Annotations de chaque point
        for i, txt in enumerate(vl_semestres):
            # S'assurer que chaque valeur est arrondie à 2 décimales
            txt_arrondi = round(txt, 2)
            ax.annotate(
                format_fr_euro(txt_arrondi),
                (dates_semestres[i], vl_semestres[i]),
                textcoords="offset points",
                xytext=(0, 10),
                ha='center',
                fontsize=9,
                color='white',
                bbox=dict(boxstyle="round,pad=0.3", fc=couleur_bleue, ec=couleur_bleue, alpha=0.9)
            )
        
        # Titres et axes
        ax.set_title(f"Atterrissage VL - {nom_fonds}", fontsize=16, fontweight='bold', color=couleur_bleue, pad=20)
        ax.set_ylabel("VL (€)", fontsize=12, color=couleur_bleue)
        
        # Ticks
        ax.set_xticks(dates_semestres)
        ax.set_xticklabels(
            [d.strftime('%b-%y').capitalize() for d in dates_semestres],
            rotation=45,
            ha='right',
            fontsize=10,
            color=couleur_bleue
        )
        ax.tick_params(axis='y', labelcolor=couleur_bleue)
        
        ax.yaxis.set_major_formatter(
            ticker.FuncFormatter(lambda x, _: f"{round(x, 2):,.2f} €".replace(",", " ").replace(".", ","))
        )
        
        # Fond
        fig.patch.set_facecolor('white')
        ax.set_facecolor('white')
        
        # Supprimer les contours inutiles
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        
        st.pyplot(fig)
        
        # === EXPORT EXCEL AVEC GRAPHIQUE ===
        try:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Exporter les données - nous n'utilisons pas to_excel directement
                # car nous voulons créer une marge d'une ligne et une colonne
                workbook = writer.book
                
                # Créer un onglet "Atterrissage VL" à la place de "Projection"
                worksheet = workbook.add_worksheet('Atterrissage VL')
                
                # Formats pour l'Excel
                header_format = workbook.add_format({
                    'bold': True,
                    'font_color': couleur_bleue,
                    'bg_color': '#F0F0F0',
                    'border': 0,
                    'align': 'center',
                    'valign': 'vcenter'
                })
                
                # Format pour cellules texte alignées à droite (pour les valeurs monétaires)
                text_right_format = workbook.add_format({
                    'align': 'right'
                })
                
                # Définir une marge pour laisser une ligne vide en haut et une colonne vide à gauche
                row_offset = 1  # Décalage d'une ligne
                col_offset = 1  # Décalage d'une colonne
                
                # Appliquer le format d'en-tête (avec décalage)
                for col_num, value in enumerate(projection.columns.values):
                    worksheet.write(row_offset, col_num + col_offset, value, header_format)
                
                # Enlever le quadrillage
                worksheet.hide_gridlines(2)  # 2 = enlever complètement le quadrillage
                
                # Créer des formats pour les nombres et montants
                number_format = workbook.add_format({
                    'align': 'right',
                    'num_format': '#,##0.00',  # Format standard pour les nombres avec décimales
                })
                
                # Format pour les nombres négatifs en rouge
                negative_number_format = workbook.add_format({
                    'align': 'right',
                    'num_format': '#,##0.00',  # Format standard pour les nombres avec décimales
                    'font_color': 'red',
                })
                
                # Créer un format pour les cellules monétaires
                money_format = workbook.add_format({
                    'align': 'right',
                    'num_format': '#,##0.00 €',  # Format monétaire sans zéros superflus
                })
                
                # Format pour les montants négatifs en rouge
                negative_money_format = workbook.add_format({
                    'align': 'right',
                    'num_format': '#,##0.00 €',  # Format monétaire sans zéros superflus
                    'font_color': 'red',
                })
                
                # Mettre en forme les colonnes monétaires et ajuster les largeurs (avec décalage)
                for idx, col in enumerate(projection.columns):
                    # Calculer la largeur optimale (plus précise)
                    col_values = projection[col].astype(str)
                    max_len = max(
                        max([len(str(s)) for s in col_values]),
                        len(col)
                    ) + 3  # Marge supplémentaire
                    
                    # Limiter la largeur maximum
                    max_len = min(max_len, 40)
                    
                    # Appliquer la largeur (avec décalage)
                    worksheet.set_column(idx + col_offset, idx + col_offset, max_len)
                    
                    # Appliquer le format monétaire si la colonne contient "€"
                    if "€" in col or any(s in col for s in ["Impact", "Actif"]):
                        # Écrire comme nombres réels avec format monétaire
                        for row_num in range(len(projection)):
                            cell_value = projection.iloc[row_num, idx]
                            
                            # Extraction de la valeur numérique avec gestion d'erreurs
                            try:
                                if isinstance(cell_value, str):
                                    if "€" in cell_value:
                                        # Format standard avec euro
                                        numeric_value = float(cell_value.replace(" ", "").replace("€", "").replace(",", "."))
                                    else:
                                        # Autre texte, essayer de convertir directement
                                        numeric_value = float(cell_value.replace(",", "."))
                                elif isinstance(cell_value, (int, float)):
                                    # Déjà un nombre
                                    numeric_value = float(cell_value)
                                else:
                                    # Si c'est un autre type (None, etc.)
                                    numeric_value = 0
                            except Exception:
                                # En cas d'erreur, mettre 0
                                numeric_value = 0
                            
                            # Appliquer le format approprié selon si le nombre est négatif ou positif (avec décalage)
                            if numeric_value < 0:
                                worksheet.write_number(row_num + row_offset + 1, idx + col_offset, numeric_value, negative_money_format)
                            else:
                                worksheet.write_number(row_num + row_offset + 1, idx + col_offset, numeric_value, money_format)
                    # Format spécial pour les colonnes VL - nombre à deux décimales
                    elif "VL" in col:
                        for row_num in range(len(projection)):
                            cell_value = projection.iloc[row_num, idx]
                            
                            # Extraction de la valeur numérique avec gestion d'erreurs
                            try:
                                if isinstance(cell_value, str):
                                    if "€" in cell_value:
                                        # Format standard avec euro
                                        numeric_value = float(cell_value.replace(" ", "").replace("€", "").replace(",", "."))
                                    else:
                                        # Autre texte, essayer de convertir directement
                                        numeric_value = float(cell_value.replace(",", "."))
                                elif isinstance(cell_value, (int, float)):
                                    # Déjà un nombre
                                    numeric_value = float(cell_value)
                                else:
                                    # Si c'est un autre type (None, etc.)
                                    numeric_value = 0
                            except Exception:
                                # En cas d'erreur, mettre 0
                                numeric_value = 0
                            
                            # Format nombre à deux décimales pour VL
                            if numeric_value < 0:
                                worksheet.write_number(row_num + row_offset + 1, idx + col_offset, numeric_value, negative_number_format)
                            else:
                                worksheet.write_number(row_num + row_offset + 1, idx + col_offset, numeric_value, number_format)
                    else:
                        # Pour les autres colonnes (non monétaires) - avec décalage
                        for row_num in range(len(projection)):
                            worksheet.write(row_num + row_offset + 1, idx + col_offset, projection.iloc[row_num, idx])
                
                # Aussi ajuster la largeur de la première colonne vide
                worksheet.set_column(0, 0, 3)  # Largeur de 3 pour la colonne vide
                
                # --- AJOUT DU GRAPHIQUE DIRECTEMENT DANS LA FEUILLE DE PROJECTION ---
                
                # Préparation des données pour le graphique (sans nouvelle feuille)
                # Nous allons ajouter ces données directement en-dessous du tableau principal
                
                # Déterminer où commencer à ajouter les données du graphique (après le tableau principal)
                # Décalé en fonction de l'ajout de la ligne vide
                start_row = len(projection) + 3 + row_offset  # +3 pour laisser un peu d'espace
                
                # Ajouter un titre pour la section graphique
                title_format = workbook.add_format({
                    'bold': True,
                    'font_size': 16,
                    'font_color': couleur_bleue,
                    'align': 'left'
                })
                worksheet.write(start_row, col_offset, f'Atterrissage VL - {nom_fonds}', title_format)
                
                # En-têtes pour les données du graphique (avec décalage)
                headers_format = workbook.add_format({
                    'bold': True,
                    'font_color': couleur_bleue,
                    'bg_color': '#F0F0F0',
                    'border': 0,
                    'align': 'center'
                })
                worksheet.write(start_row + 1, col_offset, 'Date', headers_format)
                worksheet.write(start_row + 1, col_offset + 1, 'VL', headers_format)
                
                # Format pour les dates
                date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
                
                # Écrire les données pour le graphique (avec décalage)
                for i, (date, vl) in enumerate(zip(dates_semestres, vl_semestres)):
                    row = start_row + 2 + i
                    # Écrire la date
                    worksheet.write_datetime(row, col_offset, date, date_format)
                    # Écrire la VL comme nombre
                    worksheet.write_number(row, col_offset + 1, float(vl))
                
                # Créer le graphique simple avec style similaire à l'application
                chart = workbook.add_chart({'type': 'line'})
                
                # Configurer la série de données pour ressembler au graphique de l'application
                # Note: mise à jour des références aux cellules avec le décalage
                chart.add_series({
                    'categories': [worksheet.name, start_row + 2, col_offset, start_row + 1 + len(dates_semestres), col_offset],
                    'values': [worksheet.name, start_row + 2, col_offset + 1, start_row + 1 + len(dates_semestres), col_offset + 1],
                    'marker': {
                        'type': 'circle',
                        'size': 8,
                        'border': {'color': couleur_bleue},
                        'fill': {'color': couleur_bleue}  # Points remplis de couleur bleue
                    },
                    'line': {
                        'color': couleur_bleue,
                        'width': 2.5
                    },
                    'data_labels': {
                        'value': True,
                        'position': 'above',
                        'font': {
                            'color': 'white',
                            'bold': True,
                            'size': 9
                        },
                        'num_format': '#,##0.00 "€"',  # Format sans zéros superflus
                        'border': {'color': couleur_bleue},
                        'fill': {'color': couleur_bleue}
                    },
                })
                
                # Configurer le graphique pour qu'il ressemble à celui de l'application
                chart.set_title({
                    'name': f'Atterrissage VL - {nom_fonds}',  # Même titre que dans l'interface
                    'name_font': {
                        'size': 16,  # Même taille que dans l'interface
                        'color': couleur_bleue,
                        'bold': True
                    }
                })
                
                chart.set_x_axis({
                    'name': '',  # Pas de titre d'axe
                    'num_format': 'mmm-yy',  # Format "jun-25", "dec-25"
                    'num_font': {
                        'rotation': 45,
                        'color': couleur_bleue
                    },
                    'line': {'color': couleur_bleue},
                    'major_gridlines': {'visible': False}
                })
                
                chart.set_y_axis({
                    'name': 'VL (€)',
                    'name_font': {
                        'color': couleur_bleue,
                        'bold': True
                    },
                    'num_format': '#,##0.00 "€"',  # Format sans zéros superflus
                    'num_font': {'color': couleur_bleue},
                    'line': {'color': couleur_bleue},
                    'major_gridlines': {'visible': True, 'line': {'color': '#E0E0E0', 'width': 0.5}}  # Lignes de grille légères
                })
                
                chart.set_legend({'none': True})
                chart.set_chartarea({'border': {'none': True}, 'fill': {'color': 'white'}})
                chart.set_plotarea({'border': {'none': True}, 'fill': {'color': 'white'}})
                
                # Insérer le graphique dans la feuille (avec décalage approprié pour la position du graphique)
                worksheet.insert_chart(start_row + 2, col_offset + 3, chart, {'x_scale': 1.5, 'y_scale': 1.2})
            
            buffer.seek(0)
            
            # Boutons d'exportation
            export_col1, export_col2 = st.columns(2)
            
            with export_col1:
                # Export Excel
                date_aujourd_hui = datetime.now().strftime("%Y%m%d")
                nom_fichier_excel = f"{date_aujourd_hui} - Atterrissage VL - {nom_fonds}.xlsx"
                
                st.download_button(
                    label="📥 Exporter en Excel",
                    data=buffer,
                    file_name=nom_fichier_excel,
                    mime="application/vnd.ms-excel"
                )
                
            with export_col2:
                # Export JSON
                export_data = {
                    "nom_fonds": nom_fonds,
                    "nom_scenario": nom_scenario,
                    "date_vl_connue": date_vl_connue_str,
                    "date_fin_fonds": date_fin_fonds_str,
                    "anr_derniere_vl": anr_derniere_vl,
                    "nombre_parts": nombre_parts,
                    "impacts": impacts,
                    "impacts_multidates": impacts_multidates,
                    "actifs": actifs
                }
                date_aujourd_hui = datetime.now().strftime("%Y%m%d")
                nom_fichier_json = f"{date_aujourd_hui} - {nom_fonds} - {nom_scenario}.json"
                json_export = json.dumps(export_data, indent=2).encode('utf-8')
                
                st.download_button(
                    label="📤 Exporter en JSON",
                    data=json_export,
                    file_name=nom_fichier_json,
                    mime="application/json"
                )
        except Exception as e:
            st.error(f"Erreur lors de la génération de l'export: {str(e)}")
        
        # Sauvegarder dans la session
        st.session_state.params = {
            "nom_fonds": nom_fonds,
            "nom_scenario": nom_scenario,
            "date_vl_connue": date_vl_connue_str,
            "date_fin_fonds": date_fin_fonds_str,
            "anr_derniere_vl": anr_derniere_vl,
            "nombre_parts": nombre_parts,
            "impacts": impacts,
            "impacts_multidates": impacts_multidates,
            "actifs": actifs
        }
    
    except Exception as e:
        st.error(f"Erreur lors du calcul de la projection: {str(e)}")
        import traceback
        st.error(traceback.format_exc())

with tab2:
    # === CHARGEMENT / PERSISTENCE DES PARAMÈTRES ===
    st.header("Gestion des simulations")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Importer / Exporter")
        
        # Import JSON
        params_json = st.file_uploader("Importer des paramètres JSON", type="json")
        if params_json is not None:
            try:
                imported_params = json.load(params_json)
                st.session_state.params.update(imported_params)
                st.success("Paramètres importés avec succès")
                if st.button("⚡ Appliquer les paramètres importés", type="primary"):
                    st.rerun()
            except Exception as e:
                st.error(f"Erreur lors de l'importation du fichier JSON: {str(e)}")
    
    with col2:
        st.subheader("Sauvegarder en base de données")
        
        # Option pour sauvegarder la simulation actuelle
        mode_sauvegarde = st.radio(
            "Mode de sauvegarde",
            ["Nouvelle sauvegarde", "Mettre à jour une sauvegarde existante"]
        )
        
        if mode_sauvegarde == "Nouvelle sauvegarde":
            if st.button("💾 Sauvegarder comme nouvelle simulation", type="primary"):
                # Préparer les données à sauvegarder
                params = st.session_state.params
                
                # Date au format jour/mois/année
                date_formatee = datetime.now().strftime("%d/%m/%Y")
                
                # Sauvegarder dans la BDD avec le commentaire formaté
                commentaire = f"{params.get('nom_scenario', 'Base case')} - {date_formatee}"
                simulation_id = sauvegarder_simulation(params, commentaire)
                
                if simulation_id:
                    st.success(f"Nouvelle simulation '{params.get('nom_scenario', 'Base case')}' sauvegardée avec succès")
                    st.rerun()  # Actualiser pour montrer la nouvelle simulation
                else:
                    st.error("Échec de la sauvegarde, veuillez réessayer")
        else:
            # Option pour mettre à jour une sauvegarde existante
            simulations = lister_simulations()
            
            if simulations:
                # Format d'affichage simplifié pour les simulations existantes
                options = {}
                for s in simulations:
                    # Format simplifié: Nom du fonds - Nom du scénario
                    display_text = f"{s['nom_fonds']} - {s['nom_scenario']}"
                    options[display_text] = s['id']
                    
                sim_a_mettre_a_jour = st.selectbox(
                    "Choisir la simulation à mettre à jour",
                    options=list(options.keys())
                )
                
                if st.button("🔄 Mettre à jour la simulation", type="primary"):
                    simulation_id = options[sim_a_mettre_a_jour]
                    
                    # Préparer les données à sauvegarder
                    params = st.session_state.params
                    
                    # Supprimer l'ancienne simulation
                    supprimer_simulation(simulation_id)
                    
                    # Créer une nouvelle avec les mêmes données
                    date_formatee = datetime.now().strftime("%d/%m/%Y")
                    commentaire = f"{params.get('nom_scenario', 'Base case')} - {date_formatee} (Mise à jour)"
                    new_id = sauvegarder_simulation(params, commentaire)
                    
                    if new_id:
                        st.success(f"Simulation '{params.get('nom_scenario', 'Base case')}' mise à jour avec succès")
                        st.rerun()  # Actualiser pour montrer la mise à jour
                    else:
                        st.error("Échec de la mise à jour, veuillez réessayer")
            else:
                st.info("Aucune simulation existante à mettre à jour")
    
    # Liste des simulations sauvegardées
    st.subheader("Simulations sauvegardées")
    simulations = lister_simulations()
    
    if simulations:
        # Pour chaque ligne, ajouter des boutons d'actions
        for sim in simulations:
            with st.container(border=True):
                col1, col2 = st.columns([3, 1])
                
                with col1:
                    st.markdown(f"**{sim['nom_fonds']} - {sim['nom_scenario']}**")
                    st.caption(f"Créé le {sim['date_creation'].split(' ')[0] if ' ' in sim['date_creation'] else sim['date_creation']}")
                
                with col2:
                    col_load, col_del = st.columns(2)
                    with col_load:
                        if st.button("📂 Charger", key=f"load_{sim['id']}"):
                            params_charges = charger_simulation(sim['id'])
                            if params_charges:
                                st.session_state.params = params_charges
                                st.success(f"Simulation '{sim['nom_scenario']}' chargée avec succès")
                                st.rerun()
                    with col_del:
                        if st.button("🗑️ Supprimer", key=f"del_{sim['id']}"):
                            if supprimer_simulation(sim['id']):
                                st.success("Simulation supprimée avec succès")
                                st.rerun()
    else:
        st.info("Aucune simulation sauvegardée")
    
    # Bouton de réinitialisation
    if st.button("♻️ Réinitialiser tous les paramètres", type="primary"):
        st.session_state.params = default_params.copy()
        st.success("Paramètres réinitialisés aux valeurs par défaut")
        st.rerun()

with tab3:
    st.header("Guide d'utilisation")
    
    st.markdown("""
    ### Présentation
    
    Cette application permet de simuler l'évolution de la Valeur Liquidative (VL) d'un fonds sur plusieurs semestres, en tenant compte de:
    - L'évolution des actifs du portefeuille
    - Des impacts récurrents (frais semestriels)
    - Des impacts ponctuels à des dates spécifiques
    
    ### Paramètres généraux
    
    - **Nom du fonds** : Identifiant du fonds simulé
    - **Nom du scénario** : Nom du scénario de simulation (ex: "Base case", "Stress test", etc.)
    - **Date dernière VL connue** : Point de départ de la simulation au format jj/mm/aaaa
    - **Date fin de fonds** : Horizon de la simulation au format jj/mm/aaaa
    - **ANR dernière VL connue** : Actif Net Réévalué à la date de dernière VL connue
    - **Nombre de parts** : Nombre de parts du fonds
    
    ### Gestion des actifs
    
    Pour chaque actif du portefeuille:
    - **Nom de l'actif** : Identifiant de l'actif
    - **% Détention** : Pourcentage de détention de l'actif (ex: 100% pour détention totale)
    - **IS à provisionner** : Si coché, l'application appliquera un abattement de 25% sur les plus-values
    - **Valeur actuelle** : Valeur de l'actif à la date de dernière VL connue
    - **Valeur projetée** : Valeur estimée de l'actif au semestre suivant (S+1)
    
    ### Impacts récurrents et multidates
    
    - **Impacts récurrents** : Frais ou autres impacts qui se répètent à chaque semestre
    - **Impacts multidates** : Impacts ponctuels à des dates spécifiques
    
    ### Fonctionnalités principales
    
    - Projection de la VL sur plusieurs semestres
    - Visualisation graphique de l'évolution de la VL
    - Export des résultats en Excel ou JSON
    - Sauvegarde et chargement des simulations en base de données
    """)
    
    st.info("Cette application nécessite que les dates soient au format jj/mm/aaaa et les valeurs monétaires au format X XXX,XX €")

# Afficher la version de l'application
st.sidebar.markdown("---")
st.sidebar.caption("Atterrissage VL v2.1")