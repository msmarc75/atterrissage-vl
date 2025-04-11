import streamlit as st
import pandas as pd
import json
import sqlite3
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import io
import os

# === INITIALISATION DE LA BASE DE DONNÉES ===
def init_db():
    try:
        # Vérifier si le fichier de base de données existe et s'il a une taille nulle
        db_file = 'simulations_fonds.db'
        if os.path.exists(db_file) and os.path.getsize(db_file) == 0:
            print("DEBUG: Fichier de BDD corrompu ou vide, suppression")
            os.remove(db_file)
            
        conn = sqlite3.connect('simulations_fonds.db')
        c = conn.cursor()
        
        # Table principale pour les paramètres généraux
        c.execute('''
        CREATE TABLE IF NOT EXISTS simulations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nom_fonds TEXT NOT NULL,
            date_vl_connue TEXT NOT NULL,
            date_fin_fonds TEXT NOT NULL,
            anr_derniere_vl REAL NOT NULL,
            nombre_parts REAL NOT NULL,
            date_creation TIMESTAMP NOT NULL,
            commentaire TEXT
        )
        ''')
        
        # Vérifier si la colonne nom_scenario existe déjà dans la table
        c.execute("PRAGMA table_info(simulations)")
        colonnes = c.fetchall()
        colonnes_noms = [col[1] for col in colonnes]  # Indice 1 pour le nom de colonne
        
        # Si la colonne n'existe pas, l'ajouter
        if 'nom_scenario' not in colonnes_noms:
            c.execute('''
            ALTER TABLE simulations 
            ADD COLUMN nom_scenario TEXT DEFAULT 'Base case'
            ''')
            print("Colonne nom_scenario ajoutée à la table simulations")
        
        # Table pour les impacts récurrents
        c.execute('''
        CREATE TABLE IF NOT EXISTS impacts_recurrents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            simulation_id INTEGER NOT NULL,
            libelle TEXT NOT NULL,
            montant REAL NOT NULL,
            FOREIGN KEY (simulation_id) REFERENCES simulations (id) ON DELETE CASCADE
        )
        ''')
        
        # Table pour les impacts multidates
        c.execute('''
        CREATE TABLE IF NOT EXISTS impacts_multidates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            simulation_id INTEGER NOT NULL,
            libelle TEXT NOT NULL,
            FOREIGN KEY (simulation_id) REFERENCES simulations (id) ON DELETE CASCADE
        )
        ''')
        
        # Table pour les occurrences des impacts multidates
        c.execute('''
        CREATE TABLE IF NOT EXISTS occurrences_impacts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            impact_multidate_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            montant REAL NOT NULL,
            FOREIGN KEY (impact_multidate_id) REFERENCES impacts_multidates (id) ON DELETE CASCADE
        )
        ''')
        
        # Table pour les actifs
        c.execute('''
        CREATE TABLE IF NOT EXISTS actifs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            simulation_id INTEGER NOT NULL,
            nom TEXT NOT NULL,
            pct_detention REAL NOT NULL,
            valeur_actuelle REAL NOT NULL,
            valeur_projetee REAL NOT NULL,
            is_a_provisionner BOOLEAN DEFAULT 0,
            FOREIGN KEY (simulation_id) REFERENCES simulations (id) ON DELETE CASCADE
        )
        ''')
        
        # Activer le support des clés étrangères
        c.execute("PRAGMA foreign_keys = ON")
        
        conn.commit()
        conn.close()
        
        # Log de réussite
        print("Initialisation de la BDD réussie")
        
    except Exception as e:
        # En cas d'erreur, informer clairement l'utilisateur
        st.error(f"Erreur lors de l'initialisation de la base de données: {str(e)}")
        import traceback
        st.error(traceback.format_exc())

# Fonctions pour la gestion des simulations en base de données
def sauvegarder_simulation(params, commentaire=""):
    try:
        conn = sqlite3.connect('simulations_fonds.db')
        # Activer les clés étrangères pour cette connexion
        conn.execute("PRAGMA foreign_keys = ON")
        c = conn.cursor()
        
        # S'assurer que les valeurs numériques sont bien des nombres
        try:
            anr = float(params['anr_derniere_vl'])
            parts = float(params['nombre_parts'])
        except ValueError:
            # Si conversion impossible, utiliser des valeurs par défaut
            st.warning("Problème avec les valeurs numériques, utilisation de valeurs par défaut")
            anr = 10000000.0
            parts = 10000.0
        
        # Récupérer le nom du scénario, utiliser "Base case" par défaut
        nom_scenario = params.get('nom_scenario', 'Base case')
        
        # Vérifier si la colonne nom_scenario existe dans la table
        c.execute("PRAGMA table_info(simulations)")
        colonnes = c.fetchall()
        colonnes_noms = [col[1] for col in colonnes]  # Indice 1 pour le nom de colonne
        
        # Préparer la requête d'insertion en fonction des colonnes disponibles
        if 'nom_scenario' in colonnes_noms:
            # Si la colonne existe, l'inclure dans l'insertion
            c.execute('''
            INSERT INTO simulations (
                nom_fonds, nom_scenario, date_vl_connue, date_fin_fonds, 
                anr_derniere_vl, nombre_parts, date_creation, commentaire
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                params['nom_fonds'],
                nom_scenario, 
                params['date_vl_connue'], 
                params['date_fin_fonds'],
                anr, 
                parts, 
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                commentaire
            ))
        else:
            # Si la colonne n'existe pas, l'omettre de l'insertion
            c.execute('''
            INSERT INTO simulations (
                nom_fonds, date_vl_connue, date_fin_fonds, 
                anr_derniere_vl, nombre_parts, date_creation, commentaire
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (
                params['nom_fonds'], 
                params['date_vl_connue'], 
                params['date_fin_fonds'],
                anr, 
                parts, 
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                commentaire
            ))
            # Stocker le nom du scénario dans le commentaire
            print(f"Colonne nom_scenario manquante, scénario '{nom_scenario}' stocké dans le commentaire")
        
        
        simulation_id = c.lastrowid
        print(f"DEBUG: Simulation créée avec ID={simulation_id}")
        
        # Insérer les impacts récurrents
        for libelle, montant in params['impacts']:
            try:
                montant_float = float(montant)
            except (ValueError, TypeError):
                st.warning(f"Problème avec la valeur de l'impact '{libelle}', utilisation de 0.0")
                montant_float = 0.0
                
            c.execute('''
            INSERT INTO impacts_recurrents (simulation_id, libelle, montant)
            VALUES (?, ?, ?)
            ''', (simulation_id, libelle, montant_float))
        
        # Insérer les impacts multidates
        for impact in params['impacts_multidates']:
            c.execute('''
            INSERT INTO impacts_multidates (simulation_id, libelle)
            VALUES (?, ?)
            ''', (simulation_id, impact['libelle']))
            
            impact_id = c.lastrowid
            
            # Insérer les occurrences de cet impact
            for occurrence in impact['montants']:
                try:
                    montant_float = float(occurrence['montant'])
                except (ValueError, TypeError):
                    st.warning(f"Problème avec la valeur d'occurrence pour '{impact['libelle']}', utilisation de 0.0")
                    montant_float = 0.0
                    
                c.execute('''
                INSERT INTO occurrences_impacts (impact_multidate_id, date, montant)
                VALUES (?, ?, ?)
                ''', (impact_id, occurrence['date'], montant_float))
        
        # Insérer les actifs
        for actif in params['actifs']:
            try:
                pct = float(actif['pct_detention'])
                val_act = float(actif['valeur_actuelle'])
                val_proj = float(actif['valeur_projetee'])
            except (ValueError, TypeError, KeyError):
                st.warning(f"Problème avec les valeurs de l'actif '{actif.get('nom', 'sans nom')}', utilisation de valeurs par défaut")
                pct = 1.0
                val_act = 1000000.0
                val_proj = 1050000.0
                
            c.execute('''
            INSERT INTO actifs (
                simulation_id, nom, pct_detention, 
                valeur_actuelle, valeur_projetee, is_a_provisionner
            ) VALUES (?, ?, ?, ?, ?, ?)
            ''', (
                simulation_id, 
                actif.get('nom', 'Actif sans nom'), 
                pct,
                val_act, 
                val_proj, 
                bool(actif.get('is_a_provisionner', False))
            ))
        
        print("DEBUG: Tentative de commit...")
        conn.commit()
        print("DEBUG: Commit effectué")
        conn.close()
        
        print(f"DEBUG: Sauvegarde complète de la simulation #{simulation_id}")
        return simulation_id
        
    except Exception as e:
        st.error(f"Erreur lors de la sauvegarde: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None

def charger_simulation(simulation_id):
    conn = sqlite3.connect('simulations_fonds.db')
    conn.row_factory = sqlite3.Row  # Pour accéder aux colonnes par nom
    c = conn.cursor()
    
    # Récupérer les paramètres généraux
    c.execute("SELECT * FROM simulations WHERE id = ?", (simulation_id,))
    sim_row = c.fetchone()
    
    # Vérifier que la simulation existe
    if not sim_row:
        st.error(f"Simulation avec ID {simulation_id} introuvable.")
        conn.close()
        return None
    
    # Convertir en dictionnaire
    sim = dict(sim_row)
    
    # Structure pour stocker les paramètres complets
    params = {
        'nom_fonds': sim['nom_fonds'],
        'nom_scenario': sim.get('nom_scenario', 'Base case'),  # Récupérer le nom du scénario
        'date_vl_connue': sim['date_vl_connue'],
        'date_fin_fonds': sim['date_fin_fonds'],
        'anr_derniere_vl': float(sim['anr_derniere_vl']),  # Assurer que c'est un float
        'nombre_parts': float(sim['nombre_parts']),        # Assurer que c'est un float
        'impacts': [],
        'impacts_multidates': [],
        'actifs': []
    }
    
    # Récupérer les impacts récurrents
    c.execute("SELECT libelle, montant FROM impacts_recurrents WHERE simulation_id = ?", (simulation_id,))
    impacts_rows = c.fetchall()
    params['impacts'] = [(row['libelle'], float(row['montant'])) for row in impacts_rows]
    
    # Récupérer les impacts multidates
    c.execute("SELECT id, libelle FROM impacts_multidates WHERE simulation_id = ?", (simulation_id,))
    impacts_multidates = c.fetchall()
    
    for impact in impacts_multidates:
        impact_dict = {'libelle': impact['libelle'], 'montants': []}
        
        # Récupérer les occurrences pour cet impact
        c.execute("""
        SELECT date, montant FROM occurrences_impacts 
        WHERE impact_multidate_id = ?
        """, (impact['id'],))
        
        impact_dict['montants'] = [
            {'date': row['date'], 'montant': float(row['montant'])} 
            for row in c.fetchall()
        ]
        
        params['impacts_multidates'].append(impact_dict)
    
    # Récupérer les actifs
    c.execute("""
    SELECT nom, pct_detention, valeur_actuelle, valeur_projetee, is_a_provisionner 
    FROM actifs WHERE simulation_id = ?
    """, (simulation_id,))
    
    for row in c.fetchall():
        actif = {
            'nom': row['nom'],
            'pct_detention': float(row['pct_detention']),
            'valeur_actuelle': float(row['valeur_actuelle']),
            'valeur_projetee': float(row['valeur_projetee']),
            'is_a_provisionner': bool(row['is_a_provisionner'])
        }
        
        # Calculer les valeurs dérivées qui sont attendues par l'application
        variation_brute = (actif['valeur_projetee'] - actif['valeur_actuelle']) * actif['pct_detention']
        
        # Appliquer la règle de l'IS (75% de l'impact en cas de plus-value)
        if actif['is_a_provisionner'] and variation_brute > 0:
            variation = variation_brute * 0.75  # Réduction de 25% pour l'IS
        else:
            variation = variation_brute  # Pas de modification en cas de moins-value ou si IS non provisionné
        
        actif['variation'] = variation
        actif['variation_brute'] = variation_brute
        
        params['actifs'].append(actif)
    
    conn.close()
    
    # Logging pour debug
    print(f"Paramètres chargés: {params}")
    
    return params

def lister_simulations():
    conn = sqlite3.connect('simulations_fonds.db')
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    
    # Débug - lister toutes les simulations brutes
    c.execute("SELECT * FROM simulations")
    all_rows = c.fetchall()
    print(f"DEBUG - Nombre total de simulations en base: {len(all_rows)}")
    for row in all_rows:
        try:
            print(f"  - ID: {row['id']}, Nom: {row['nom_fonds']}, Scénario: {row.get('nom_scenario', 'N/A')}")
        except Exception as e:
            print(f"  - Erreur lecture row: {e}")
    
    # Vérifier si la colonne nom_scenario existe déjà dans la table
    c.execute("PRAGMA table_info(simulations)")
    colonnes = c.fetchall()
    colonnes_noms = [col[1] for col in colonnes]  # Indice 1 pour le nom de colonne
    
    if 'nom_scenario' in colonnes_noms:
        # Si la colonne existe, on l'inclut dans la requête
        c.execute("""
        SELECT id, nom_fonds, nom_scenario, date_vl_connue, date_creation, commentaire 
        FROM simulations ORDER BY date_creation DESC
        """)
    else:
        # Si la colonne n'existe pas encore, on ne la sélectionne pas
        c.execute("""
        SELECT id, nom_fonds, date_vl_connue, date_creation, commentaire 
        FROM simulations ORDER BY date_creation DESC
        """)
    
    simulations = []
    for row in c.fetchall():
        # Convertir en dictionnaire
        sim_dict = dict(row)
        
        # Nettoyer le nom du fonds (enlever les dates potentielles)
        if sim_dict['nom_fonds'] and '(' in sim_dict['nom_fonds']:
            # Si le nom contient une parenthèse (comme une date), prendre juste la partie avant
            sim_dict['nom_fonds'] = sim_dict['nom_fonds'].split('(')[0].strip()
        
        # Assurer qu'il y a un nom_scenario
        if 'nom_scenario' not in sim_dict or not sim_dict['nom_scenario']:
            # Essayer d'extraire du commentaire
            if sim_dict['commentaire'] and ' - ' in sim_dict['commentaire']:
                sim_dict['nom_scenario'] = sim_dict['commentaire'].split(' - ')[0].strip()
            else:
                sim_dict['nom_scenario'] = 'Base case'
        
        simulations.append(sim_dict)
    
    conn.close()
    
    # Déboguer - afficher les infos importantes pour chaque simulation
    print(f"DEBUG - Nombre de simulations traitées: {len(simulations)}")
    for i, sim in enumerate(simulations):
        print(f"Simulation #{i+1}:")
        print(f"  ID: {sim['id']}")
        print(f"  Nom fonds: '{sim['nom_fonds']}'")
        print(f"  Nom scénario: '{sim['nom_scenario']}'")
        print(f"  Commentaire: '{sim['commentaire']}'")
        print("------------------------")
    
    return simulations

def supprimer_simulation(simulation_id):
    conn = sqlite3.connect('simulations_fonds.db')
    # Activer les clés étrangères pour cette connexion
    conn.execute("PRAGMA foreign_keys = ON")
    c = conn.cursor()
    
    # Supprimer les occurrences d'impacts multidates
    c.execute("""
        DELETE FROM occurrences_impacts
        WHERE impact_multidate_id IN (
            SELECT id FROM impacts_multidates
            WHERE simulation_id = ?
        )
    """, (simulation_id,))
    
    # Supprimer les impacts multidates
    c.execute("DELETE FROM impacts_multidates WHERE simulation_id = ?", (simulation_id,))
    
    # Supprimer les impacts récurrents
    c.execute("DELETE FROM impacts_recurrents WHERE simulation_id = ?", (simulation_id,))
    
    # Supprimer les actifs
    c.execute("DELETE FROM actifs WHERE simulation_id = ?", (simulation_id,))
    
    # Supprimer la simulation elle-même
    c.execute("DELETE FROM simulations WHERE id = ?", (simulation_id,))
    
    conn.commit()
    conn.close()

# Initialiser la base de données au démarrage
init_db()

st.title("Atterrissage VL")

def format_fr_euro(valeur):
    # S'assurer que la valeur est arrondie à deux décimales
    valeur_arrondie = round(valeur, 2)
    return f"{valeur_arrondie:,.2f} €".replace(",", " ").replace(".", ",")

def champ_numerique(label, valeur):
    champ = st.sidebar.text_input(label, value=format_fr_euro(valeur))
    champ = champ.replace(" ", "").replace(",", ".").replace("€", "")
    return float(champ) if champ else 0.0

# === PARAMÈTRES INITIAUX ===
default_params = {
    "nom_fonds": "FPCI IE 1",
    "nom_scenario": "Base case",
    "date_vl_connue": "31/12/2024",
    "date_fin_fonds": "31/12/2028",
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
    "actifs": []
}

# === CHARGEMENT / PERSISTENCE DES PARAMÈTRES ===
st.sidebar.header("Importer / Exporter")
params_json = st.sidebar.file_uploader("Importer des paramètres JSON", type="json")

if 'params' not in st.session_state:
    st.session_state.params = default_params.copy()

if params_json is not None:
    st.session_state.params.update(json.load(params_json))

params = st.session_state.params

# === GESTION DES SIMULATIONS ===
st.sidebar.header("Gestion des simulations en BDD")

# Option pour sauvegarder la simulation actuelle
with st.sidebar.expander("💾 Sauvegarder la simulation"):
    # Récupérer les valeurs actuelles de manière sécurisée
    current_nom_fonds = params.get('nom_fonds', "FPCI IE 1")
    current_nom_scenario = params.get('nom_scenario', "Base case")
    
    # Vérifier si les variables locales existent et les utiliser si c'est le cas
    if 'nom_fonds' in locals() and nom_fonds:
        current_nom_fonds = nom_fonds
    if 'nom_scenario' in locals() and nom_scenario:
        current_nom_scenario = nom_scenario
    
    # Option pour créer une nouvelle sauvegarde ou mettre à jour
    mode_sauvegarde = st.radio(
        "Mode de sauvegarde",
        ["Nouvelle sauvegarde", "Mettre à jour une sauvegarde existante"]
    )
    
    if mode_sauvegarde == "Nouvelle sauvegarde":
        if st.button("Sauvegarder comme nouvelle simulation"):
            # Préparer les données à sauvegarder avec une vérification sécurisée de toutes les variables
            export_data = {
                "nom_fonds": current_nom_fonds,
                "nom_scenario": current_nom_scenario,
                "date_vl_connue": params.get('date_vl_connue', "31/12/2024"),
                "date_fin_fonds": params.get('date_fin_fonds', "31/12/2028"),
                "anr_derniere_vl": params.get('anr_derniere_vl', 10_000_000.0),
                "nombre_parts": params.get('nombre_parts', 10_000.0),
                "impacts": params.get('impacts', []),
                "impacts_multidates": params.get('impacts_multidates', []),
                "actifs": params.get('actifs', [])
            }
            
            # Utiliser les variables locales si elles existent
            if 'date_vl_connue_str' in locals() and date_vl_connue_str:
                export_data["date_vl_connue"] = date_vl_connue_str
            if 'date_fin_fonds_str' in locals() and date_fin_fonds_str:
                export_data["date_fin_fonds"] = date_fin_fonds_str
            if 'anr_derniere_vl' in locals() and anr_derniere_vl:
                export_data["anr_derniere_vl"] = anr_derniere_vl
            if 'nombre_parts' in locals() and nombre_parts:
                export_data["nombre_parts"] = nombre_parts
            if 'impacts' in locals() and impacts:
                export_data["impacts"] = impacts
            if 'impacts_multidates' in locals() and impacts_multidates:
                export_data["impacts_multidates"] = impacts_multidates
            if 'actifs' in locals() and actifs:
                export_data["actifs"] = actifs
            
            # Date au format jour/mois/année
            date_formatee = datetime.now().strftime("%d/%m/%Y")
            
            # Sauvegarder dans la BDD avec le commentaire formaté
            commentaire = f"{current_nom_scenario} - {date_formatee}"
            simulation_id = sauvegarder_simulation(export_data, commentaire)
            
            if simulation_id:
                st.sidebar.success(f"Nouvelle simulation '{current_nom_scenario}' sauvegardée avec succès")
                # Forcer un changement dans l'état de session pour déclencher un rechargement
                if 'refresh_counter' not in st.session_state:
                    st.session_state.refresh_counter = 0
                st.session_state.refresh_counter += 1
                print(f"DEBUG: Rechargement forcé, compteur={st.session_state.refresh_counter}")
                st.rerun()
            else:
                st.sidebar.error("Échec de la sauvegarde, veuillez réessayer")
    else:
        # Option pour mettre à jour une sauvegarde existante
        simulations = lister_simulations()
        
        if simulations:
            # Format d'affichage simplifié pour les simulations existantes
            options = {}
            for s in simulations:
                # Récupérer le nom du scénario
                if 'nom_scenario' in s and s['nom_scenario']:
                    scenario = s['nom_scenario']
                else:
                    # Essayer d'extraire du commentaire
                    commentaire = s['commentaire'] if s['commentaire'] else "Sans nom"
                    if ' - ' in commentaire:
                        scenario = commentaire.split(' - ')[0]
                    else:
                        scenario = commentaire
                
                # Format simplifié: Nom du fonds - Nom du scénario
                display_text = f"{s['nom_fonds']} - {scenario}"
                options[display_text] = s['id']
                
            sim_a_mettre_a_jour = st.selectbox(
                "Choisir la simulation à mettre à jour",
                options=list(options.keys())
            )
            
            if st.button("Mettre à jour la simulation"):
                simulation_id = options[sim_a_mettre_a_jour]
                
                # Préparer les données à sauvegarder avec une vérification sécurisée de toutes les variables
                export_data = {
                    "nom_fonds": current_nom_fonds,
                    "nom_scenario": current_nom_scenario,
                    "date_vl_connue": params.get('date_vl_connue', "31/12/2024"),
                    "date_fin_fonds": params.get('date_fin_fonds', "31/12/2028"),
                    "anr_derniere_vl": params.get('anr_derniere_vl', 10_000_000.0),
                    "nombre_parts": params.get('nombre_parts', 10_000.0),
                    "impacts": params.get('impacts', []),
                    "impacts_multidates": params.get('impacts_multidates', []),
                    "actifs": params.get('actifs', [])
                }
                
                # Utiliser les variables locales si elles existent
                if 'date_vl_connue_str' in locals() and date_vl_connue_str:
                    export_data["date_vl_connue"] = date_vl_connue_str
                if 'date_fin_fonds_str' in locals() and date_fin_fonds_str:
                    export_data["date_fin_fonds"] = date_fin_fonds_str
                if 'anr_derniere_vl' in locals() and anr_derniere_vl:
                    export_data["anr_derniere_vl"] = anr_derniere_vl
                if 'nombre_parts' in locals() and nombre_parts:
                    export_data["nombre_parts"] = nombre_parts
                if 'impacts' in locals() and impacts:
                    export_data["impacts"] = impacts
                if 'impacts_multidates' in locals() and impacts_multidates:
                    export_data["impacts_multidates"] = impacts_multidates
                if 'actifs' in locals() and actifs:
                    export_data["actifs"] = actifs
                
                # Supprimer l'ancienne simulation
                supprimer_simulation(simulation_id)
                
                # Créer une nouvelle avec les mêmes données
                date_formatee = datetime.now().strftime("%d/%m/%Y")
                commentaire = f"{current_nom_scenario} - {date_formatee} (Mise à jour)"
                new_id = sauvegarder_simulation(export_data, commentaire)
                
                if new_id:
                    st.sidebar.success(f"Simulation '{current_nom_scenario}' mise à jour avec succès")
                    # Forcer un changement dans l'état de session pour déclencher un rechargement
                    if 'refresh_counter' not in st.session_state:
                        st.session_state.refresh_counter = 0
                    st.session_state.refresh_counter += 1
                    st.rerun()
                else:
                    st.sidebar.error("Échec de la mise à jour, veuillez réessayer")
        else:
            st.info("Aucune simulation existante à mettre à jour")

# Option pour charger une simulation existante
with st.sidebar.expander("📂 Charger une simulation"):
    simulations = lister_simulations()
    
    if simulations:
        options = {f"{s['nom_fonds']} - {s['nom_scenario']}": s['id'] for s in simulations}
        sim_selectionnee = st.selectbox(
            "Choisir une simulation à charger",
            options=list(options.keys())
        )
        
        col1, col2 = st.columns(2)

    with col1:
    if st.button("Charger cette simulation"):
        simulation_id = options[sim_selectionnee]  # Cette ligne et les suivantes doivent être indentées
        params_charges = charger_simulation(simulation_id)
        st.session_state.params = params_charges
        st.rerun()
        
        with col2:
            if st.button("🗑️ Supprimer"):
            simulation_id = options[sim_selectionnee]
            supprimer_simulation(simulation_id)
            st.success("Simulation supprimée avec succès")
            st.rerun()
    else:
        st.info("Aucune simulation sauvegardée")

# === SAISIE UTILISATEUR ===
# Liste des fonds disponibles
fonds_disponibles = [
    "FPCI IE 1",
    "FPCI IE 2",
    "NIC 2",
    "FPCI Recovery",
    "NFS 1",
    "NIA",
    "NIO",
    "NIO 2",
    "NIO 3"
]

# Utiliser le selectbox pour le nom du fonds au lieu d'un text_input
nom_fonds = st.sidebar.selectbox(
    "Nom du fonds",
    options=fonds_disponibles,
    index=fonds_disponibles.index(params['nom_fonds']) if params['nom_fonds'] in fonds_disponibles else 0
)

nom_scenario = st.sidebar.text_input("Nom du scénario", params.get('nom_scenario', 'Base case'))
date_vl_connue_str = st.sidebar.text_input("Date dernière VL connue (jj/mm/aaaa)", params['date_vl_connue'])
date_fin_fonds_str = st.sidebar.text_input("Date fin de fonds (jj/mm/aaaa)", params['date_fin_fonds'])
anr_derniere_vl = champ_numerique("ANR dernière VL connue (€)", params['anr_derniere_vl'])
nombre_parts = champ_numerique("Nombre de parts", params['nombre_parts'])

# === DATES & PROJECTION ===
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

# Création de la liste des dates formatées pour le selectbox
dates_semestres_str = [d.strftime("%d/%m/%Y") for d in dates_semestres]

# Impacts personnalisés récurrents
st.sidebar.header("Impacts semestriels récurrents")
impacts = []
nb_impacts = st.sidebar.number_input("Nombre d'impacts récurrents", min_value=0, value=len(params['impacts']), step=1)
for i in range(nb_impacts):
    if i < len(params['impacts']):
        libelle_defaut, montant_defaut = params['impacts'][i]
    else:
        libelle_defaut, montant_defaut = f"Impact récurrent {i+1}", 0.0
    libelle = st.sidebar.text_input(f"Libellé impact récurrent {i+1}", libelle_defaut)
    montant = champ_numerique(f"Montant semestriel impact récurrent {i+1} (€)", montant_defaut)
    impacts.append((libelle, montant))

# Impacts multidates (à plusieurs dates différentes)
st.sidebar.header("Impacts multidates")
impacts_multidates = []
nb_impacts_multidates = st.sidebar.number_input("Nombre d'impacts multidates", min_value=0, 
                                             value=len(params.get('impacts_multidates', [])), step=1)

for i in range(nb_impacts_multidates):
    st.sidebar.subheader(f"Impact multidate {i+1}")
    
    # Récupération des paramètres par défaut s'ils existent
    if i < len(params.get('impacts_multidates', [])):
        impact_default = params['impacts_multidates'][i]
        libelle_defaut = impact_default['libelle']
        montants_defaut = impact_default['montants']
    else:
        libelle_defaut = f"Impact multidate {i+1}"
        montants_defaut = []
    
    # Libellé de l'impact multidate
    libelle = st.sidebar.text_input(f"Libellé impact multidate {i+1}", libelle_defaut)
    
    # Nombre d'occurrences pour cet impact
    nb_occurrences = st.sidebar.number_input(
        f"Nombre d'occurrences pour '{libelle}'", 
        min_value=0, 
        value=len(montants_defaut), 
        step=1
    )
    
    # Liste pour stocker les montants et dates de cet impact
    montants = []
    
    # Interface pour chaque occurrence
    for j in range(nb_occurrences):
        st.sidebar.markdown(f"##### Occurrence {j+1} de '{libelle}'")
        
        # Valeurs par défaut pour cette occurrence
        if j < len(montants_defaut):
            montant_defaut = montants_defaut[j]['montant']
            date_defaut = montants_defaut[j]['date']
            date_index = dates_semestres_str.index(date_defaut) if date_defaut in dates_semestres_str else 0
        else:
            montant_defaut = 0.0
            date_index = 0
        
        # Champ pour le montant
        montant = champ_numerique(f"Montant {j+1} de '{libelle}' (€)", montant_defaut)
        
        # Selectbox pour la date
        date_str = st.sidebar.selectbox(
            f"Date {j+1} de '{libelle}'",
            options=dates_semestres_str,
            index=date_index
        )
        
        montants.append({
            "date": date_str,
            "montant": montant
        })
    
    impacts_multidates.append({
        "libelle": libelle,
        "montants": montants
    })

# Actifs
st.sidebar.header("Ajouter des Actifs")
actifs = []
nb_actifs = st.sidebar.number_input("Nombre d'actifs à ajouter", min_value=1, value=max(1, len(params['actifs'])), step=1)
for i in range(nb_actifs):
    st.sidebar.subheader(f"Actif {i+1}")
    if i < len(params['actifs']):
        a = params['actifs'][i]
        nom_defaut = a['nom']
        pct_defaut = a['pct_detention'] * 100
        val_actuelle = a.get('valeur_actuelle', 1_000_000.0)
        val_proj = a.get('valeur_projetee', val_actuelle + 50_000)
    else:
        nom_defaut = f"Actif {i+1}"
        pct_defaut = 100.0
        val_actuelle = 1_000_000.0
        val_proj = 1_050_000.0

    nom_actif = st.sidebar.text_input(f"Nom Actif {i+1}", nom_defaut)
    # Remplace le slider par un champ de saisie directe du pourcentage au format français
    pct_text = st.sidebar.text_input(f"% Détention Actif {i+1}", value=f"{pct_defaut:.2f}".replace(".", ","))
    try:
        pct_detention = float(pct_text.replace(",", "."))
    except ValueError:
        pct_detention = pct_defaut  # En cas d'erreur, utiliser la valeur par défaut
    
    # Option pour provisionner l'IS
    is_a_provisionner = False
    if i < len(params['actifs']) and 'is_a_provisionner' in params['actifs'][i]:
        is_a_provisionner = params['actifs'][i]['is_a_provisionner']
    is_a_provisionner = st.sidebar.checkbox(f"IS à provisionner pour Actif {i+1} ?", value=is_a_provisionner)
    
    valeur_actuelle = champ_numerique(f"Valeur actuelle Actif {i+1} (€)", val_actuelle)
    valeur_projetee = champ_numerique(f"Valeur projetée en S+1 Actif {i+1} (€)", val_proj)
    
    # Calcul de la variation avec prise en compte de l'IS si nécessaire
    variation_brute = (valeur_projetee - valeur_actuelle) * (pct_detention / 100)
    
    # Appliquer la règle de l'IS (75% de l'impact en cas de plus-value)
    if is_a_provisionner and variation_brute > 0:
        variation = variation_brute * 0.75  # Réduction de 25% pour l'IS
    else:
        variation = variation_brute  # Pas de modification en cas de moins-value ou si IS non provisionné
    
    actifs.append({
        "nom": nom_actif,
        "pct_detention": pct_detention / 100,
        "valeur_actuelle": valeur_actuelle,
        "valeur_projetee": valeur_projetee,
        "variation": variation,
        "is_a_provisionner": is_a_provisionner,
        "variation_brute": variation_brute
    })

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
        var = a['variation'] if i == 1 else 0
        
        # Si IS à provisionner et il s'agit d'une plus-value, afficher uniquement le montant net
        if i == 1 and a['is_a_provisionner'] and a['variation_brute'] > 0:
            # Afficher seulement le montant net, sans texte supplémentaire
            row[f"Actif - {a['nom']}"] = format_fr_euro(var)
        else:
            row[f"Actif - {a['nom']}"] = format_fr_euro(var)
            
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
        libelle = impact['libelle']
        # Initialiser la valeur à 0 par défaut
        impact_valeur = 0
        
        # Parcourir toutes les occurrences de cet impact
        for occurrence in impact['montants']:
            if occurrence['date'] == date_str:
                impact_valeur += occurrence['montant']
        
        # Ajouter au total
        total_impacts_multidates += impact_valeur
        
        # Afficher dans le tableau
        row[f"Impact multidate - {libelle}"] = format_fr_euro(impact_valeur)

    if i > 0:
        anr_courant += total_var_actifs + total_impacts_recurrents
    # Appliquer les impacts multidates
    anr_courant += total_impacts_multidates

    vl = anr_courant / nombre_parts
    # Arrondir à deux décimales
    vl = round(vl, 2)
    vl_semestres.append(vl)
    row["VL prévisionnelle (€)"] = format_fr_euro(vl)

    projection_rows.append(row)

# === AFFICHAGE TABLEAU ===
projection = pd.DataFrame(projection_rows)
st.subheader("VL prévisionnelle")
st.dataframe(projection)

# Définir Helvetica comme police globale
plt.rcParams['font.family'] = 'Open Sans'

# === GRAPHIQUE BLEU STYLÉ ===
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
                
                # SECTION CORRIGÉE: Extraction de la valeur numérique avec gestion d'erreurs
                try:
                    if isinstance(cell_value, str):
                        if "€" in cell_value:
                            # Format standard avec euro
                            numeric_value = float(cell_value.replace(" ", "").replace("€", "").replace(",", "."))
                        elif "brut:" in cell_value and "net:" in cell_value:
                            # Format spécial pour IS: "brut: X XXX,XX € | net: X XXX,XX €"
                            # Extraire seulement la valeur nette
                            parts = cell_value.split("|")
                            if len(parts) >= 2:
                                net_part = parts[1].strip()
                                net_value = net_part.replace("net:", "").strip()
                                numeric_value = float(net_value.replace(" ", "").replace("€", "").replace(",", "."))
                            else:
                                # Si le format n'est pas exactement comme attendu
                                # Essayons de récupérer le montant net d'une autre façon
                                net_parts = cell_value.split("net:")
                                if len(net_parts) >= 2:
                                    net_value = net_parts[1].strip()
                                    numeric_value = float(net_value.replace(" ", "").replace("€", "").replace(",", "."))
                                else:
                                    st.warning(f"Format de cellule avec IS non reconnu: {cell_value}")
                                    numeric_value = 0
                        else:
                            # Autre texte, essayer de convertir directement
                            numeric_value = float(cell_value.replace(",", "."))
                    elif isinstance(cell_value, (int, float)):
                        # Déjà un nombre
                        numeric_value = float(cell_value)
                    else:
                        # Si c'est un autre type (None, etc.)
                        st.warning(f"Type non géré rencontré dans l'export Excel: {type(cell_value)}, valeur: {cell_value}")
                        numeric_value = 0
                except Exception as e:
                    # En cas d'erreur, afficher un message de débogage et mettre 0
                    st.warning(f"Erreur lors de la conversion de la valeur '{cell_value}' (type: {type(cell_value)}): {str(e)}")
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
                
                # SECTION CORRIGÉE: Extraction de la valeur numérique avec gestion d'erreurs
                try:
                    if isinstance(cell_value, str):
                        if "€" in cell_value:
                            # Format standard avec euro
                            numeric_value = float(cell_value.replace(" ", "").replace("€", "").replace(",", "."))
                        elif "brut:" in cell_value and "net:" in cell_value:
                            # Format spécial pour IS: "brut: X XXX,XX € | net: X XXX,XX €"
                            # Extraire seulement la valeur nette
                            parts = cell_value.split("|")
                            if len(parts) >= 2:
                                net_part = parts[1].strip()
                                net_value = net_part.replace("net:", "").strip()
                                numeric_value = float(net_value.replace(" ", "").replace("€", "").replace(",", "."))
                            else:
                                # Si le format n'est pas exactement comme attendu
                                # Essayons de récupérer le montant net d'une autre façon
                                net_parts = cell_value.split("net:")
                                if len(net_parts) >= 2:
                                    net_value = net_parts[1].strip()
                                    numeric_value = float(net_value.replace(" ", "").replace("€", "").replace(",", "."))
                                else:
                                    st.warning(f"Format de cellule avec IS non reconnu: {cell_value}")
                                    numeric_value = 0
                        else:
                            # Autre texte, essayer de convertir directement
                            numeric_value = float(cell_value.replace(",", "."))
                    elif isinstance(cell_value, (int, float)):
                        # Déjà un nombre
                        numeric_value = float(cell_value)
                    else:
                        # Si c'est un autre type (None, etc.)
                        st.warning(f"Type non géré rencontré dans l'export Excel: {type(cell_value)}, valeur: {cell_value}")
                        numeric_value = 0
                except Exception as e:
                    # En cas d'erreur, afficher un message de débogage et mettre 0
                    st.warning(f"Erreur lors de la conversion de la valeur '{cell_value}' (type: {type(cell_value)}): {str(e)}")
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
        'name':
        'VL (€)',
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

# Simplification des boutons d'exportation (retiré l'export PDF)
date_aujourd_hui = datetime.now().strftime("%Y%m%d")
nom_fichier_excel = f"{date_aujourd_hui} - Atterrissage VL - {nom_fonds}.xlsx"

st.download_button(
    label="📥 Exporter la projection avec graphique Excel",
    data=buffer,
    file_name=nom_fichier_excel,
    mime="application/vnd.ms-excel"
)

# === EXPORT PARAMÈTRES JSON ===
if st.sidebar.button("📤 Exporter les paramètres JSON"):
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
    st.sidebar.download_button("Télécharger paramètres JSON", json_export, file_name=nom_fichier_json)


# === BOUTON RÉINITIALISATION ===
if st.sidebar.button("♻️ Réinitialiser les paramètres"):
    st.session_state.params = default_params.copy()
    st.rerun()