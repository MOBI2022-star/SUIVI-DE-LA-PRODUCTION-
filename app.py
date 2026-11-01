
import io, os, re, sqlite3, smtplib
from email.message import EmailMessage
from datetime import date, datetime, timedelta
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Suivi de production vélos", page_icon="🚲", layout="wide")
DB_PATH = os.getenv("DB_PATH", "production.db")

PIECES_DEFAUT = [
"Batteries","VBOX","Cache cadre","Cache cadre nouveau modèle","Panier","Roue arrière","Roue avant",
"Chargeurs","Capteurs de couples","Vis brid VBOX","Bouchon de câble anti vol",
"Fixation de bridage pour les VBOX","Tendeur de chaine","Vis de fixation pour les VBOX","Vis de carter",
"Manivelle droite VAE","Manivelle gauche VAE","Manivelle droite VLS","Manivelle gauche VLS",
"Couronne","Sticker VLS","Sticker VAE","Garde boue arrière VAE","Garde boue avant VAE",
"Garde boue arrière VLS","Garde boue avant VLS","Tige de selle","Selle","Béquille","Levier de frein",
"Carter","Catadiope","Roue avant VLS","Roue arrière VLS","Roue avant VAE","Roue arrière VAE",
"Poignée droite","Poignée gauche","Moteur fourche","Fourche VAE","Fourche VLS","Facade VAE","Facade VLS",
"Couronne VAE","Couronne VLS","Pédale gauche","Pédale droite","Chaine","Manette de vitesse","Guidon",
"Collier de selle","Collier de fourche","Rallonge cable moteur VAE","Cable phare garde-boue AR",
"Cable rallonge feux arrieres","Frein avant","Frein arrières","Levier de frein D","Levier de frein G",
"Feux arrieres","Sonette","Poigné Gauche","Pedalle Droite","Pedalle Gauche","Poingé Droite",
"Cable antivol","Carter Droite","Carter gauche","Gaine derailleur","Cable frein avant",
"Cable frein arriere","Gaine de frein avant","Gaine de frein arrière","Ecrou anti vole roue avant VAE",
"Ecrou anti vole roue avant VLS","Rondelle anti-couple roue avant VAE","Embout de gaine frein",
"Embout de gaine derailleur","Bouchon antivol"
]
EVENEMENTS = ["Zone de contrôle fermée (encombrement vélo)","Evacuation","Manque de magasiniers",
"Manque de personnel poste retour","Manque de personnel pour diag électrique","Batterie arrivant en quantité limitée",
"Ecart entre le diag et le contrôle","Vélo bloqué au diapason","Problème de peering",
"VBOX ne reconnaissant pas badge client","Manque de vélos","CARL Source dysfonctionnement",
"CARL Touch dysfonctionnement","Coupure d'électricité","Interruption de la connexion internet"]

def db():
    con=sqlite3.connect(DB_PATH, check_same_thread=False)
    con.execute("""CREATE TABLE IF NOT EXISTS daily(
      work_date TEXT PRIMARY KEY, week INTEGER, completed INTEGER DEFAULT 0, settled INTEGER DEFAULT 0, total_production INTEGER DEFAULT 0,
      staff INTEGER DEFAULT 0, returns INTEGER DEFAULT 0, charge_hours REAL DEFAULT 0,
      comment TEXT DEFAULT '', updated_at TEXT)""")
    con.execute("""CREATE TABLE IF NOT EXISTS breakdowns(
      id INTEGER PRIMARY KEY AUTOINCREMENT, work_date TEXT, piece TEXT, start_time TEXT,
      end_time TEXT, minutes REAL DEFAULT 0, note TEXT DEFAULT '')""")
    con.execute("""CREATE TABLE IF NOT EXISTS work_orders(
      intervention TEXT PRIMARY KEY, bike TEXT, status TEXT, end_date TEXT, mechanic TEXT,
      imported_at TEXT)""")
    con.execute("""CREATE TABLE IF NOT EXISTS imports(
      id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT, imported_at TEXT, rows_count INTEGER,
      notes TEXT)""")
    con.commit()
    return con
con=db()

def week_num(d):
    return date.fromisoformat(str(d)).isocalendar().week

def safe_int(v):
    try: return int(float(v)) if pd.notna(v) else 0
    except: return 0

def read_excel(upload):
    return pd.ExcelFile(upload)

def normalize(s):
    return re.sub(r"[^a-z0-9]+","",str(s).lower().strip())

def parse_history(upload):
    """Importe les feuilles hebdomadaires existantes et récupère les valeurs par date."""
    xls=pd.ExcelFile(upload)
    found=[]
    for sh in xls.sheet_names:
        try:
            raw=pd.read_excel(upload, sheet_name=sh, header=None)
            if raw.empty or raw.shape[0] < 9: continue
            # Les colonnes de jours sont en général C à G ; la ligne 0 contient les dates.
            for c in range(2, min(raw.shape[1],7)):
                dt=pd.to_datetime(raw.iat[0,c], errors="coerce")
                if pd.isna(dt): continue
                d=dt.date().isoformat()
                vals={}
                for r in range(min(len(raw),10)):
                    label=" ".join(str(v) for v in raw.iloc[r,:2].tolist() if pd.notna(v)).lower()
                    if "terminé" in label or "termine" in label: vals["completed"]=safe_int(raw.iat[r,c])
                    if "soldé" in label or "solde" in label: vals["settled"]=safe_int(raw.iat[r,c])
                    if re.search(r"\btotal\b", label) and "total j+1" not in label and "retour" not in label: vals["total_production"]=safe_int(raw.iat[r,c])
                    if "effectifs présents" in label: vals["staff"]=safe_int(raw.iat[r,c])
                    if "nombre de retours" in label: vals["returns"]=safe_int(raw.iat[r,c])
                    if "charge de travail quotidienne" in label:
                        v=raw.iat[r,c]
                        try:
                            td=pd.to_timedelta(str(v))
                            vals["charge_hours"]=td.total_seconds()/3600
                        except: pass
                if vals and (vals.get("completed",0) or vals.get("settled",0) or vals.get("staff",0)):
                    found.append({"work_date":d, **vals, "week":date.fromisoformat(d).isocalendar().week})
        except Exception:
            continue
    return pd.DataFrame(found).drop_duplicates("work_date", keep="last") if found else pd.DataFrame()

def import_workbook(file):
    hist=parse_history(file)
    if hist.empty: return 0
    for _,r in hist.iterrows():
        old=con.execute("SELECT work_date FROM daily WHERE work_date=?",(r.work_date,)).fetchone()
        cols=["work_date","week","completed","settled","total_production","staff","returns","charge_hours","updated_at"]
        vals=[r.work_date,int(r.week),int(r.get("completed",0)),int(r.get("settled",0)),
              int(r.get("total_production",r.get("completed",0)+r.get("settled",0))),
              int(r.get("staff",0)),int(r.get("returns",0)),float(r.get("charge_hours",0)),datetime.now().isoformat(timespec="seconds")]
        con.execute("""INSERT INTO daily(work_date,week,completed,settled,total_production,staff,returns,charge_hours,updated_at)
        VALUES(?,?,?,?,?,?,?,?,?) ON CONFLICT(work_date) DO UPDATE SET
        week=excluded.week, completed=excluded.completed, settled=excluded.settled,
        total_production=excluded.total_production, staff=excluded.staff, returns=excluded.returns,
        charge_hours=excluded.charge_hours, updated_at=excluded.updated_at""",vals)
    con.execute("INSERT INTO imports(filename,imported_at,rows_count,notes) VALUES(?,?,?,?)",
                (file.name,datetime.now().isoformat(timespec="seconds"),len(hist),"Historique hebdomadaire importé"))
    con.commit()
    return len(hist)

def pct(a,b): return (100*a/b) if b else 0

st.title("🚲 Suivi de production – Atelier vélos")
st.caption("Tableau de bord quotidien et hebdomadaire • Historique conservé dans production.db")

with st.sidebar:
    st.header("Importation Excel")
    hist_file=st.file_uploader("Classeur de suivi hebdomadaire (.xlsx)",type=["xlsx"],key="hist")
    if hist_file and st.button("Importer l'historique du classeur"):
        try:
            n=import_workbook(hist_file)
            st.success(f"{n} journées importées / mises à jour.")
        except Exception as e: st.error(f"Import impossible : {e}")
    st.divider()
    st.caption("Les données sont enregistrées dans la base locale production.db. Sur GitHub, prévoyez un stockage persistant pour l'hébergement.")

tabs=st.tabs(["📊 Tableau du jour","📅 Semaines S1 à S52","🧰 Ruptures de pièces","🚲 Vélos non soldés","📈 Graphiques","📥 Imports & exports"])

with tabs[0]:
    st.subheader("Saisie et synthèse du jour")
    d=st.date_input("Date de production", value=date.today())
    ds=d.isoformat()
    old=con.execute("SELECT completed,settled,total_production,staff,returns,charge_hours,comment FROM daily WHERE work_date=?",(ds,)).fetchone()
    old=old or (0,0,0,0,0,0.0,"")
    a,b,c,e,f=st.columns(5)
    completed=a.number_input("Vélos terminés",min_value=0,value=int(old[0]),step=1)
    settled=b.number_input("Vélos soldés",min_value=0,value=int(old[1]),step=1)
    total_prod=c.number_input("Production totale (ligne Total Excel)",min_value=0,value=int(old[2] or (old[0]+old[1])),step=1)
    staff=e.number_input("Mécaniciens présents",min_value=0,value=int(old[3]),step=1)
    returns=f.number_input("Retours",min_value=0,value=int(old[4]),step=1)
    charge=st.number_input("Charge totale (heures)",min_value=0.0,value=float(old[5]),step=0.25)
    comment=st.text_input("Commentaire / information du jour",value=str(old[6] or ""))
    if st.button("💾 Enregistrer la journée",type="primary"):
        con.execute("""INSERT INTO daily(work_date,week,completed,settled,total_production,staff,returns,charge_hours,comment,updated_at)
        VALUES(?,?,?,?,?,?,?,?,?,?) ON CONFLICT(work_date) DO UPDATE SET
        week=excluded.week,completed=excluded.completed,settled=excluded.settled,total_production=excluded.total_production,
        staff=excluded.staff,returns=excluded.returns,charge_hours=excluded.charge_hours,comment=excluded.comment,updated_at=excluded.updated_at""",
        (ds,d.isocalendar().week,completed,settled,total_prod,staff,returns,charge,comment,datetime.now().isoformat(timespec="seconds")))
        con.commit(); st.success("Journée enregistrée.")
    total=total_prod
    k=st.columns(5)
    k[0].metric("Mécaniciens présents",staff)
    k[1].metric("Vélos terminés",completed)
    k[2].metric("Vélos soldés",settled)
    k[3].metric("% soldés / production",f"{pct(settled,total):.1f}%")
    k[4].metric("À contrôler / terminés",f"{pct(completed,total):.1f}%")
    k2=st.columns(3)
    k2[0].metric("Production totale du jour",total)
    k2[1].metric("Vélos / mécanicien",f"{total/staff:.2f}" if staff else "—")
    k2[2].metric("Charge moyenne / vélo",f"{charge/total:.2f} h" if total else "—")
    st.caption("La production totale reprend la ligne « Total » du classeur (modifiable). Les pourcentages terminés / soldés sont calculés sur cette valeur.")
    st.info("Le total de production est saisi séparément car le classeur S43 présente des écarts entre la ligne Total et la somme terminés + soldés.")

with tabs[1]:
    st.subheader("Vue hebdomadaire")
    w=st.selectbox("Semaine",list(range(1,53)),index=date.today().isocalendar().week-1,format_func=lambda x:f"S{x:02d}")
    start=date.fromisocalendar(date.today().isocalendar().year,w,1)
    dates=[start+timedelta(days=i) for i in range(7)]
    rows=[]
    for day in dates:
        r=con.execute("SELECT completed,settled,total_production,staff,returns,charge_hours,comment,updated_at FROM daily WHERE work_date=?",(day.isoformat(),)).fetchone()
        rows.append({"Jour":day.strftime("%A %d/%m"),"Date":day.isoformat(),"Terminés":r[0] if r else 0,
        "Soldés":r[1] if r else 0,"Effectif":r[2] if r else 0,"Retours":r[3] if r else 0,
        "Charge (h)":r[4] if r else 0.0,"Vélos/ETP":((r[0]+r[1])/r[2]) if r and r[2] else 0,
        "Dernière mise à jour":r[6] if r else ""})
    df=pd.DataFrame(rows)
    st.dataframe(df,use_container_width=True,hide_index=True)
    week_total=int(df["Total production"].sum())
    q=st.columns(4)
    q[0].metric("Production semaine",week_total)
    q[1].metric("Terminés",int(df["Terminés"].sum()))
    q[2].metric("Soldés",int(df["Soldés"].sum()))
    q[3].metric("Effectif moyen",f"{df.loc[df.Effectif>0,'Effectif'].mean():.1f}" if (df.Effectif>0).any() else "—")
    st.download_button("Télécharger la semaine en CSV",df.to_csv(index=False).encode("utf-8-sig"),f"S{w:02d}_production.csv","text/csv")

with tabs[2]:
    st.subheader("Suivi des ruptures de pièces")
    with st.form("rupture"):
        dbr=st.date_input("Jour de rupture",value=date.today(),key="bd")
        piece=st.selectbox("Pièce en rupture",PIECES_DEFAUT)
        cc1,cc2=st.columns(2)
        debut=cc1.time_input("Heure de début",value=datetime.strptime("07:00","%H:%M").time())
        fin=cc2.time_input("Heure de fin",value=datetime.strptime("10:00","%H:%M").time())
        note=st.text_input("Commentaire (facultatif)")
        ok=st.form_submit_button("Ajouter la rupture")
        if ok:
            dt1=datetime.combine(dbr,debut); dt2=datetime.combine(dbr,fin)
            if dt2<dt1: dt2+=timedelta(days=1)
            mins=(dt2-dt1).total_seconds()/60
            con.execute("INSERT INTO breakdowns(work_date,piece,start_time,end_time,minutes,note) VALUES(?,?,?,?,?,?)",
                        (dbr.isoformat(),piece,debut.strftime("%H:%M"),fin.strftime("%H:%M"),mins,note))
            con.commit(); st.success(f"Durée calculée : {mins/60:.2f} heure(s)")
    br=pd.read_sql_query("SELECT work_date AS Date,piece AS Pièce,start_time AS Début,end_time AS Fin,minutes AS Minutes,note AS Commentaire FROM breakdowns ORDER BY work_date DESC",con)
    if not br.empty:
        br["Durée"]=(br["Minutes"]/60).round(2).astype(str)+" h"
        st.dataframe(br,use_container_width=True,hide_index=True)
        st.download_button("Exporter les ruptures",br.to_csv(index=False).encode("utf-8-sig"),"ruptures_pieces.csv","text/csv")
    else: st.info("Aucune rupture enregistrée.")

with tabs[3]:
    st.subheader("Liste des vélos non soldés")
    st.write("Importe l'export des ordres de travail, puis associe les colonnes de ton fichier aux champs ci-dessous.")
    wo=st.file_uploader("Export ordres de travail (.xlsx ou .csv)",type=["xlsx","xls","csv"],key="wo")
    if wo:
        try:
            raw=pd.read_csv(wo) if wo.name.lower().endswith(".csv") else pd.read_excel(wo)
            st.write("Aperçu de l'export",raw.head(8))
            cols=list(raw.columns)
            c1,c2,c3=st.columns(3)
            int_col=c1.selectbox("Colonne numéro d'intervention",cols)
            status_col=c2.selectbox("Colonne statut",cols)
            end_col=c3.selectbox("Colonne date de fin",["(aucune)"]+cols)
            c4,c5=st.columns(2)
            bike_col=c4.selectbox("Colonne matériel / vélo",["(aucune)"]+cols)
            mech_col=c5.selectbox("Colonne mécanicien",["(aucune)"]+cols)
            if st.button("Importer / actualiser les ordres de travail"):
                imported=0
                for _,r in raw.iterrows():
                    iid=str(r[int_col]).strip()
                    if not iid or iid.lower()=="nan": continue
                    status=str(r[status_col]).strip()
                    end=str(r[end_col])[:10] if end_col!="(aucune)" and pd.notna(r[end_col]) else ""
                    bike=str(r[bike_col]) if bike_col!="(aucune)" else ""
                    mech=str(r[mech_col]) if mech_col!="(aucune)" else ""
                    con.execute("""INSERT INTO work_orders(intervention,bike,status,end_date,mechanic,imported_at)
                    VALUES(?,?,?,?,?,?) ON CONFLICT(intervention) DO UPDATE SET bike=excluded.bike,status=excluded.status,
                    end_date=excluded.end_date,mechanic=excluded.mechanic,imported_at=excluded.imported_at""",
                    (iid,bike,status,end,mech,datetime.now().isoformat(timespec="seconds")))
                    imported+=1
                con.execute("INSERT INTO imports(filename,imported_at,rows_count,notes) VALUES(?,?,?,?)",
                (wo.name,datetime.now().isoformat(timespec="seconds"),imported,"Ordres de travail actualisés"))
                con.commit(); st.success(f"{imported} ordres importés / actualisés.")
        except Exception as ex: st.error(f"Lecture du fichier impossible : {ex}")
    allwo=pd.read_sql_query("SELECT intervention AS Intervention,bike AS Matériel,status AS Statut,end_date AS 'Date de fin',mechanic AS Mécanicien FROM work_orders",con)
    if not allwo.empty:
        stat=st.text_input("Filtrer les statuts considérés comme soldés (mots séparés par virgule)",value="soldé,solde,clos,terminé,termine,closed")
        done=[s.strip().lower() for s in stat.split(",")]
        mask=allwo["Statut"].fillna("").str.lower().apply(lambda x:any(z in x for z in done))
        unsold=allwo[~mask]
        st.metric("Vélos / interventions non soldés",len(unsold))
        st.dataframe(unsold,use_container_width=True,hide_index=True)
        st.download_button("Télécharger la liste des non soldés (Excel)",unsold.to_csv(index=False).encode("utf-8-sig"),"velos_non_soldes.csv","text/csv")
        with st.expander("Préparer un e-mail avec la liste"):
            to=st.text_input("Adresse e-mail destinataire")
            if st.button("Créer le brouillon e-mail"):
                body="Bonjour,\n\nVeuillez trouver ci-dessous la liste des vélos non soldés :\n\n"+unsold.to_string(index=False)+"\n\nCordialement"
                st.code(f"mailto:{to}?subject=Liste%20des%20v%C3%A9los%20non%20sold%C3%A9s")
                st.caption("Le bouton prépare un lien mailto. L'envoi automatique nécessite une configuration SMTP sécurisée sur le serveur.")
    else: st.info("Aucun ordre de travail importé.")

with tabs[4]:
    st.subheader("Analyses de production")
    data=pd.read_sql_query("SELECT * FROM daily ORDER BY work_date",con)
    if not data.empty:
        data["Date"]=pd.to_datetime(data["work_date"])
        data["Production"]=data["total_production"]
        data["Mois"]=data["Date"].dt.to_period("M").astype(str)
        st.plotly_chart(px.line(data,x="Date",y="Production",markers=True,title="Production totale par jour"),use_container_width=True)
        monthly=data.groupby("Mois",as_index=False).agg(Production=("Production","sum"),Charge=("charge_hours","sum"),Effectif_moyen=("staff","mean"))
        st.plotly_chart(px.bar(monthly,x="Mois",y="Production",title="Production par mois"),use_container_width=True)
        st.plotly_chart(px.bar(data,x="Date",y="charge_hours",title="Charge totale par jour (heures)"),use_container_width=True)
        data["Vélos par mécanicien"]=data.apply(lambda r:r["Production"]/r["staff"] if r["staff"] else 0,axis=1)
        st.plotly_chart(px.line(data,x="Date",y="Vélos par mécanicien",markers=True,title="Moyenne de vélos par mécanicien"),use_container_width=True)
    else: st.info("Importe ton historique ou saisis des journées pour afficher les graphiques.")

with tabs[5]:
    st.subheader("Historique des imports")
    imp=pd.read_sql_query("SELECT filename AS Fichier, imported_at AS 'Date import', rows_count AS 'Lignes', notes AS Détail FROM imports ORDER BY id DESC",con)
    st.dataframe(imp,use_container_width=True,hide_index=True)
    st.subheader("Exporter la base de suivi")
    for table in ["daily","breakdowns","work_orders"]:
        data=pd.read_sql_query(f"SELECT * FROM {table}",con)
        st.download_button(f"Exporter {table}.csv",data.to_csv(index=False).encode("utf-8-sig"),f"{table}.csv","text/csv",key=f"exp_{table}")
    st.warning("La base SQLite conserve les saisies entre les redémarrages sur un disque persistant. Sur Streamlit Community Cloud, ajoute un stockage persistant externe pour éviter la perte de la base lors de certaines opérations de déploiement.")
