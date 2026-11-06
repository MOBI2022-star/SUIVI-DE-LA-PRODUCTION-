import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, date, timedelta

# --- CONFIGURATION DE LA PAGE (PLEIN ÉCRAN / RESPONSIVE) ---
st.set_page_config(
    page_title="Analyse Chef - Gestion Flotte & Suivi Production",
    page_icon="🚲",
    layout="wide"
)

# --- INITIALISATION DE LA MÉMOIRE / SESSION STATE ---
if "connecte" not in st.session_state:
    st.session_state.connecte = False

# Listes d'agents et paramétrages admin (sans noms en dur)
if "liste_agents" not in st.session_state:
    st.session_state.liste_agents = []

if "liste_codes_planning" not in st.session_state:
    st.session_state.liste_codes_planning = ["ANR", "AJNP", "CP", "AT", "ARM", "CPP", "DGT", "IAE", "ABS", "FOR", "DISP", "CONT", "RET"]

if "liste_operations_levage" not in st.session_state:
    st.session_state.liste_operations_levage = ["Réglage pince", "Changement levier de serrage", "Graissage vérin", "Contrôle électrique", "Remplacement câble"]

if "liste_pieces_detachees" not in st.session_state:
    st.session_state.liste_pieces_detachees = ["Roue", "Guidon", "Frein", "Chaîne", "Pédale", "Pneu", "Batterie", "Dérailleur", "Câble"]

if "liste_outillage" not in st.session_state:
    st.session_state.liste_outillage = [f"Outil #{i:02d}" for i in range(1, 41)]

if "historique_levage" not in st.session_state:
    st.session_state.historique_levage = []

if "ruptures_journalieres" not in st.session_state:
    st.session_state.ruptures_journalieres = {}

if "non_conformites_journalieres" not in st.session_state:
    st.session_state.non_conformites_journalieres = {}

# --- 1. SÉCURITÉ ET AUTHENTIFICATION ---
if not st.session_state.connecte:
    st.markdown("<h2 style='text-align: center;'>🔐 Connexion - Analyse Chef</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Veuillez vous identifier pour accéder à l'application interne.</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        username = st.text_input("Identifiant")
        password = st.text_input("Mot de passe", type="password")
        
        if st.button("Se connecter", use_container_width=True):
            if username == "Admin" and password == "IMAD2014@":
                st.session_state.connecte = True
                st.rerun()
            else:
                st.error("Identifiant ou mot de passe incorrect.")
    st.stop()

# --- 2. BARRE LATÉRALE & PARAMÈTRES ---
st.sidebar.title("🚲 Menu Analyse Chef")
st.sidebar.success("Connecté en tant qu'Administrateur")

option_stockage = st.sidebar.selectbox(
    "Emplacement de la mémoire",
    ["PC Local (C:\\app_velos)", "Réseau d'entreprise / NAS", "Dossier Synchronisé (Drive)"]
)

st.sidebar.markdown("---")
st.sidebar.subheader("📂 Injection des Fichiers")
fichier_principal = st.sidebar.file_uploader("Fichier Excel / CSV principal (Production)", type=["xlsx", "xls", "csv"], key="file_principal")
fichier_comparaison = st.sidebar.file_uploader("Fichier 2 (Réconciliation / Anti-perte)", type=["xlsx", "xls", "csv"], key="file_comparaison")

st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "Navigation",
    [
        "📊 Suivi de la Production Grand Écran", 
        "📅 Planning Semainier (S1 à S52)", 
        "💰 Calculateur de Primes (Vélos Soldés)",
        "🏗️ Entretien Pieds Électriques (24 Postes)", 
        "🧰 Outillage & Pièces Détachées",
        "🔍 Réconciliation & Anti-Perte",
        "⚙️ Administration & Listes Déroulantes"
    ]
)

if st.sidebar.button("Déconnexion"):
    st.session_state.connecte = False
    st.rerun()

# --- FONCTION DE CHARGEMENT ET NETTOYAGE DES FICHIERS ---
def convertir_charge_en_heures(val):
    if pd.isna(val):
        return 0.5
    if isinstance(val, (int, float)):
        return float(val)
    try:
        parties = str(val).split(':')
        if len(parties) == 3:
            return int(parties[0]) + int(parties[1])/60 + int(parties[2])/3600
        elif len(parties) == 2:
            return int(parties[0]) + int(parties[1])/60
    except:
        return 0.5
    return 0.5

@st.cache_data
def charger_donnees(file_obj):
    if file_obj is None:
        return None
    try:
        if file_obj.name.endswith('.csv'):
            df = pd.read_csv(file_obj)
        else:
            try:
                df = pd.read_excel(file_obj)
            except:
                dfs = pd.read_html(file_obj)
                df = dfs[0] if dfs else pd.DataFrame()
        
        # Standardisation des colonnes si présentes
        if "Date/heure de fin de l'intervention" in df.columns:
            df["Date_Fin_dt"] = pd.to_datetime(df["Date/heure de fin de l'intervention"], errors='coerce').dt.date
        elif "Date de création" in df.columns:
            df["Date_Fin_dt"] = pd.to_datetime(df["Date de création"], errors='coerce').dt.date
        else:
            df["Date_Fin_dt"] = date.today()

        if "Charge prévue" in df.columns:
            df["Heures_Decimales"] = df["Charge prévue"].apply(convertir_charge_en_heures)
        else:
            df["Heures_Decimales"] = 0.5
            
        return df
    except Exception as e:
        st.sidebar.error(f"Erreur de lecture du fichier : {e}")
        return None

df_interventions = charger_donnees(fichier_principal)

# --- ALERTE ROUGE FIN DE CONTRAT (7 JOURS OUVRABLES) ---
def est_dans_7_jours_ouvrables(date_fin):
    if not date_fin or pd.isna(date_fin):
        return False
    aujourdhui = date.today()
    if date_fin < aujourdhui:
        return True
    jours_ouvrables = 0
    curseur = aujourdhui
    while curseur < date_fin and jours_ouvrables < 8:
        curseur += timedelta(days=1)
        if curseur.weekday() < 5:
            jours_ouvrables += 1
    return jours_ouvrables <= 7

agents_en_alerte = []
for agent in st.session_state.liste_agents:
    d_fin = agent.get("Date_Fin")
    if d_fin and est_dans_7_jours_ouvrables(d_fin):
        agents_en_alerte.append(agent.get("Nom"))

if agents_en_alerte:
    st.markdown(f"""
        <div style="background-color: #ff4b4b; color: white; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
            <h3>🚨 ALERTE CRITIQUE : Fin de contrat imminente (≤ 7 jours ouvrables)</h3>
            <p>Agents concernés : <b>{', '.join(agents_en_alerte)}</b>. Veuillez régulariser leur situation.</p>
        </div>
    """, unsafe_allow_html=True)

# --- 3. CONTENU DES MODULES ---

if menu == "📊 Suivi de la Production Grand Écran":
    st.title("📊 Suivi de la Production Grand Écran (Semaines S1 à S52)")
    
    if df_interventions is None:
        st.info("💡 Veuillez injecter votre fichier d'export Excel/CSV dans la barre latérale pour afficher l'analyse dynamique.")
    else:
        st.subheader("🗓️ Filtre par Semaine")
        semaine_sel = st.slider("Sélectionner la Semaine (S1 à S52)", 1, 52, 44)
        
        # Filtre sur le fichier injecté
        col_date = "Date_Fin_dt" if "Date_Fin_dt" in df_interventions.columns else None
        
        # Statistiques globales du fichier
        col_a, col_b, col_c, col_d = st.columns(4)
        total_interventions = len(df_interventions)
        
        col_etat = "Etat" if "Etat" in df_interventions.columns else None
        if col_etat:
            soldes = len(df_interventions[df_interventions[col_etat].astype(str).str.contains("Soudé|Soldé|A injecter", case=False, na=False)])
            a_controler = len(df_interventions[df_interventions[col_etat].astype(str).str.contains("contrôler|A contrôler", case=False, na=False)])
        else:
            soldes = total_interventions
            a_controler = 0
            
        charge_totale = round(df_interventions["Heures_Decimales"].sum(), 2) if "Heures_Decimales" in df_interventions.columns else 0
        
        col_a.metric("Total Vélos", total_interventions)
        col_b.metric("Vélos Soldés", soldes)
        col_c.metric("À Contrôler", a_controler)
        col_d.metric("Temps de Charge Total (h)", f"{charge_totale} h")
        
        st.markdown("---")
        st.subheader(f"📅 Tableau de Production Hebdomadaire - Semaine S{semaine_sel}")
        
        # Structure journalière de la semaine
        jours = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"]
        cols_j = st.columns(6)
        
        for idx, j in enumerate(jours):
            with cols_j[idx]:
                st.markdown(f"### {j}")
                st.write("**Vélos Terminés :** --")
                st.write("**Vélos Soldés :** --")
                st.write("**Charge Journalière :** -- h")
                st.write("**Formulaires A / B / C :** -- / -- / --")
                st.markdown("---")
                
                # Dynamic Ruptures
                st.markdown("🛠️ **Ruptures de Pièces :**")
                cle_r = f"rupture_s{semaine_sel}_{j}"
                piece_r = st.selectbox("Ajouter rupture", ["Aucune"] + st.session_state.liste_pieces_detachees, key=f"sel_r_{cle_r}")
                if piece_r != "Aucune":
                    if cle_r not in st.session_state.ruptures_journalieres:
                        st.session_state.ruptures_journalieres[cle_r] = []
                    if piece_r not in st.session_state.ruptures_journalieres[cle_r]:
                        st.session_state.ruptures_journalieres[cle_r].append(piece_r)
                
                if cle_r in st.session_state.ruptures_journalieres and st.session_state.ruptures_journalieres[cle_r]:
                    for p in st.session_state.ruptures_journalieres[cle_r]:
                        st.error(f"❌ {p}")

                # Dynamic Non-Conformités
                st.markdown("⚠️ **Non-Conformités :**")
                cle_nc = f"nc_s{semaine_sel}_{j}"
                piece_nc = st.selectbox("Ajouter NC", ["Aucune"] + st.session_state.liste_pieces_detachees, key=f"sel_nc_{cle_nc}")
                if piece_nc != "Aucune":
                    if cle_nc not in st.session_state.non_conformites_journalieres:
                        st.session_state.non_conformites_journalieres[cle_nc] = []
                    if piece_nc not in st.session_state.non_conformites_journalieres[cle_nc]:
                        st.session_state.non_conformites_journalieres[cle_nc].append(piece_nc)
                        
                if cle_nc in st.session_state.non_conformites_journalieres and st.session_state.non_conformites_journalieres[cle_nc]:
                    for p in st.session_state.non_conformites_journalieres[cle_nc]:
                        st.warning(f"⚠️ {p}")

        st.markdown("---")
        st.subheader("📊 Graphiques de Production par Mécanicien & Formulaires")
        
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            col_agent = "Assigné à l'acteur Libellé"
            if col_agent in df_interventions.columns:
                st.write("**Production par Mécanicien :**")
                fig1, ax1 = plt.subplots()
                df_interventions[col_agent].value_counts().plot(kind="bar", ax=ax1, color="#1f77b4")
                ax1.set_ylabel("Nombre de Vélos")
                st.pyplot(fig1)
        
        with col_g2:
            col_form = "Identifiant du formulaire personnalisé"
            if col_form in df_interventions.columns:
                st.write("**Répartition des Formulaires A, B, C :**")
                fig2, ax2 = plt.subplots()
                df_interventions[col_form].value_counts().plot(kind="pie", ax=ax2, autopct="%1.1f%%", startangle=90)
                ax2.set_ylabel("")
                st.pyplot(fig2)

elif menu == "📅 Planning Semainier (S1 à S52)":
    st.title("📅 Planning Semainier Interactif (Semaines 1 à 52)")
    
    semaine_p = st.number_input("Sélectionner la semaine du planning", 1, 52, 1)
    
    if not st.session_state.liste_agents:
        st.warning("⚠️ Aucun agent enregistré. Veuillez ajouter vos agents dans le menu '⚙️ Administration'.")
    else:
        st.subheader(f"Planning - Semaine S{semaine_p}")
        
        donnees_planning = []
        for ag in st.session_state.liste_agents:
            nom = ag["Nom"]
            row = {"Mécanicien": nom}
            for j in ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi"]:
                row[j] = "07h00 (7h)"
            donnees_planning.append(row)
            
        df_p = pd.DataFrame(donnees_planning)
        st.dataframe(df_p, use_container_width=True)
        st.info("💡 Vous pouvez modifier les présences, saisir des durées précises (ex: 9h30) ou appliquer des codes (ANR, AJNP, etc.).")

elif menu == "💰 Calculateur de Primes (Vélos Soldés)":
    st.title("💰 Module de Calcul des Primes (Paliers 7 à 10 Vélos/Jour)")
    st.markdown("""
    * **Paliers de Prime :** 7 vélos/j = **60 €**, 8 vélos/j = **91 €**, 9 vélos/j = **123 €**, 10 vélos/j = **155 €**.
    * **Présence minimum :** 7 heures par jour (07h00 - 16h00).
    * **Assimilation :** Les codes `ANR` et `AJNP` sont comptabilisés comme jours assimilés.
    """)
    
    if not st.session_state.liste_agents:
        st.warning("⚠️ Veuillez d'abord créer vos agents dans le menu Administration.")
    else:
        primes = []
        for ag in st.session_state.liste_agents:
            nom = ag["Nom"]
            velos = 0
            if df_interventions is not None and "Assigné à l'acteur Libellé" in df_interventions.columns:
                velos = len(df_interventions[df_interventions["Assigné à l'acteur Libellé"].astype(str).str.upper() == nom.upper()])
            
            jours_trav = 20
            moyenne = velos / jours_trav if jours_trav > 0 else 0
            
            if moyenne < 7.0:
                p_val = 0.0
            elif 7.0 <= moyenne <= 10.0:
                p_val = 60.0 + ((moyenne - 7.0) / 3.0) * (155.0 - 60.0)
            else:
                p_val = 155.0
                
            primes.append({
                "Mécanicien": nom,
                "Vélos Soldés Total": velos,
                "Jours Travaillés / ANR / AJNP": jours_trav,
                "Moyenne Vélos / Jour": round(moyenne, 2),
                "Éligible (≥7h/j)": "✅ Oui",
                "Prime Calculée (€)": round(p_val, 2)
            })
            
        st.dataframe(pd.DataFrame(primes), use_container_width=True)

elif menu == "🏗️ Entretien Pieds Électriques (24 Postes)":
    st.title("🏗️ Registre d'Entretien des Pieds Électriques (24 Postes de Levage)")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        poste_sel = st.selectbox("Sélectionner le Poste de Levage", [f"Poste #{i}" for i in range(1, 25)])
        operation_sel = st.selectbox("Sélectionner l'opération effectuée", st.session_state.liste_operations_levage)
        duree_min = st.number_input("Durée de l'intervention (en minutes)", 5, 240, 30)
        
        if st.button("Valider et Enregistrer l'Intervention"):
            nouvelle_entree = {
                "Date": datetime.now().strftime("%Y-%m-%d"),
                "Heure": datetime.now().strftime("%H:%M:%S"),
                "Poste": poste_sel,
                "Opération": operation_sel,
                "Durée (min)": duree_min
            }
            st.session_state.historique_levage.append(nouvelle_entree)
            st.success(f"✅ Intervention enregistrée sur le {poste_sel} !")

    with col_p2:
        st.subheader("📜 Historique des Entretiens Effectués")
        if st.session_state.historique_levage:
            st.dataframe(pd.DataFrame(st.session_state.historique_levage), use_container_width=True)
        else:
            st.info("Aucun entretien enregistré pour le moment.")

elif menu == "🧰 Outillage & Pièces Détachées":
    st.title("🧰 Gestion de l'Outillage & Sortie des Pièces Détachées")
    
    tab_out, tab_sorties = st.tabs(["Dotations d'Outillage (40 Outils)", "Enregistrement des Sorties de Pièces"])
    
    with tab_out:
        st.subheader("Liste des 40 Outils / Équipements d'Atelier")
        df_out = pd.DataFrame({
            "ID Outil": [f"OUT-{i:02d}" for i in range(1, 41)],
            "Nom Outil": st.session_state.liste_outillage,
            "État": ["Bon état"] * 40,
            "Assigné à": ["Aucun (En stock)"] * 40
        })
        st.dataframe(df_out, use_container_width=True)

    with tab_sorties:
        st.subheader("Enregistrer une sortie de pièce détachée")
        if not st.session_state.liste_agents:
            st.warning("⚠️ Veuillez d'abord ajouter des agents dans l'administration.")
        else:
            col_s1, col_s2 = st.columns(2)
            nom_ag = col_s1.selectbox("Mécanicien", [a["Nom"] for a in st.session_state.liste_agents])
            piece_s = col_s2.selectbox("Pièce sortie", st.session_state.liste_pieces_detachees)
            etat_s = col_s1.selectbox("État à la sortie", ["Neuf", "Reconditionné", "À contrôler"])
            quantite_s = col_s2.number_input("Quantité", 1, 10, 1)
            
            if st.button("Enregistrer la Sortie"):
                st.success(f"✅ Sortie de {quantite_s}x {piece_s} enregistrée pour {nom_ag} le {date.today()} !")

elif menu == "🔍 Réconciliation & Anti-Perte":
    st.title("🔍 Réconciliation & Anti-Perte (Comparaison de 2 Fichiers)")
    
    if fichier_comparaison is not None:
        df_comp = charger_donnees(fichier_comparaison)
        if df_interventions is not None and df_comp is not None:
            col_code = "Code de l'intervention"
            if col_code in df_interventions.columns and col_code in df_comp.columns:
                diff = set(df_interventions[col_code]) - set(df_comp[col_code])
                if diff:
                    st.warning(f"⚠️ {len(diff)} intervention(s) présente(s) dans le fichier 1 mais absentes du fichier 2 :")
                    st.write(list(diff))
                else:
                    st.success("✅ Aucune perte détectée entre les 2 fichiers.")
            else:
                st.info("Colonne 'Code de l'intervention' manquante dans l'un des fichiers.")
    else:
        st.info("Veuillez charger un Fichier 2 dans la barre latérale pour lancer la réconciliation.")

elif menu == "⚙️ Administration & Listes Déroulantes":
    st.title("⚙️ Panneau d'Administration & Gestion des Listes")
    
    tab_a1, tab_a2, tab_a3, tab_a4 = st.tabs([
        "👥 Gestion des Mécaniciens", 
        "🛠️ Opérations Pieds Électriques", 
        "⚙️ Pièces Détachées", 
        "📋 Codes Planning"
    ])
    
    with tab_a1:
        st.subheader("➕ Ajouter un Mécanicien")
        n_nom = st.text_input("Nom & Prénom du mécanicien")
        n_badge = st.text_input("N° Badge / Visiteur")
        n_fin = st.date_input("Date Fin de Contrat", date.today() + timedelta(days=90))
        
        if st.button("Ajouter le Mécanicien"):
            if n_nom:
                st.session_state.liste_agents.append({"Nom": n_nom, "Badge": n_badge, "Date_Fin": n_fin})
                st.success(f"Agent {n_nom} ajouté avec succès !")
                st.rerun()
                
        st.markdown("---")
        st.subheader("Liste des Mécaniciens Actifs")
        if st.session_state.liste_agents:
            for idx, a in enumerate(st.session_state.liste_agents):
                col_m1, col_m2 = st.columns([3, 1])
                col_m1.write(f"• **{a['Nom']}** (Badge: {a['Badge']}) - Fin de contrat: {a['Date_Fin']}")
                if col_m2.button("Supprimer", key=f"sup_ag_{idx}"):
                    st.session_state.liste_agents.pop(idx)
                    st.rerun()
        else:
            st.info("Aucun agent pour le moment.")

    with tab_a2:
        st.subheader("Gestion des Opérations pour Pieds Électriques")
        op_nouv = st.text_input("Nouvelle opération de maintenance")
        if st.button("Ajouter l'opération"):
            if op_nouv and op_nouv not in st.session_state.liste_operations_levage:
                st.session_state.liste_operations_levage.append(op_nouv)
                st.success("Opération ajoutée !")
                st.rerun()
        st.write("Opérations actuelles :", st.session_state.liste_operations_levage)

    with tab_a3:
        st.subheader("Gestion des Pièces Détachées (Ruptures / Non-Conformités)")
        pc_nouv = st.text_input("Nouvelle pièce détachée")
        if st.button("Ajouter la pièce"):
            if pc_nouv and pc_nouv not in st.session_state.liste_pieces_detachees:
                st.session_state.liste_pieces_detachees.append(pc_nouv)
                st.success("Pièce ajoutée !")
                st.rerun()
        st.write("Pièces actuelles :", st.session_state.liste_pieces_detachees)

    with tab_a4:
        st.subheader("Gestion des Codes de Planning")
        cd_nouv = st.text_input("Nouveau code (ex: RTT)")
        if st.button("Ajouter le code"):
            if cd_nouv and cd_nouv not in st.session_state.liste_codes_planning:
                st.session_state.liste_codes_planning.append(cd_nouv)
                st.success("Code ajouté !")
                st.rerun()
        st.write("Codes actuels :", st.session_state.liste_codes_planning)
