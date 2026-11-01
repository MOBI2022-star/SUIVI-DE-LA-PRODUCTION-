# Suivi de production vélos

Application Streamlit pour saisir et analyser la production quotidienne, conserver l'historique, suivre les ruptures de pièces et importer les ordres de travail.

## Lancer en local

```bash
python -m venv .venv
# Windows : .venv\\Scripts\\activate
# macOS/Linux : source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Puis ouvrir l'adresse locale affichée par Streamlit.

## Déployer sur GitHub + Streamlit Community Cloud

1. Créer un dépôt GitHub et y déposer `app.py`, `requirements.txt` et `README.md`.
2. Ouvrir https://share.streamlit.io/ et connecter GitHub.
3. Choisir le dépôt, la branche et `app.py`.
4. Déployer.

## Importer les classeurs fournis

- Dans la barre latérale, importer le classeur de suivi `.xlsx`. L'application recherche les dates et les lignes « terminé », « soldé », effectif, retours et charge dans les feuilles hebdomadaires.
- Dans l'onglet « Vélos non soldés », importer l'export des ordres de travail et associer les colonnes intervention, statut, date de fin, matériel et mécanicien.

## Données et limites à connaître

- Les saisies sont conservées dans `production.db` (SQLite).
- Pour un hébergement cloud, prévoir un volume persistant ou une base externe (PostgreSQL/Supabase) : le disque local d'un hébergement gratuit peut être éphémère.
- Le classeur historique est importé dans les données journalières ; l'application ne réécrit pas le fichier Excel source.
- Les champs de l'export des ordres de travail sont configurables, car les en-têtes exacts varient selon les exports.
- L'envoi direct par SMTP n'est pas activé par défaut. L'application permet de télécharger le CSV et de préparer un courriel ; l'envoi automatique nécessite un serveur SMTP et des secrets configurés.
- La logique de pourcentage suppose que « terminés » et « soldés » sont deux groupes distincts. Confirmer cette définition avec le responsable du fichier source avant d'utiliser les indicateurs comme reporting officiel.
