# Atterrissage VL

Application de simulation d'évolution de Valeur Liquidative (VL) pour les fonds d'investissement.

## Fonctionnalités

- Projection de VL sur plusieurs semestres
- Gestion d'impacts récurrents (frais semestriels)
- Gestion d'impacts ponctuels à dates spécifiques
- Modélisation de l'évolution des actifs du portefeuille
- Prise en compte de l'IS sur les plus-values
- Visualisation graphique de l'évolution de la VL
- Export Excel et JSON
- Sauvegarde des simulations en base de données

## Installation

```bash
pip install -r requirements.txt
```

## Lancement de l'application

```bash
streamlit run app.py
```

## Structure du projet

- `app.py` : Application principale Streamlit
- `requirements.txt` : Dépendances Python
- `data/` : Répertoire de stockage des données (simulations sauvegardées en SQLite)

## Utilisation

1. Renseignez les paramètres généraux du fonds
2. Ajoutez des impacts récurrents et multidates
3. Configurez les actifs du portefeuille
4. Consultez la projection VL et les graphiques
5. Exportez les résultats ou sauvegardez la simulation

## Licence

MIT