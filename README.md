
# Rapport de Ventes HelloAsso

Ce projet est un script Python permettant de générer des rapports de ventes détaillés à partir de l'API HelloAsso.

## Fonctionnalités

- Récupération des commandes via l'API HelloAsso.
- Calcul du chiffre d'affaires, des bénéfices et des statistiques par produit.
- Génération de fichiers CSV pour un suivi détaillé.
- Envoi d'un email avec un résumé HTML et des fichiers attachés.
- Graphique du chiffre d'affaires par jour inclus dans l'email.

## Installation

1. **Cloner le dépôt :**

```bash
git clone https://github.com/shigaepouyen/helloasso-report.git
cd helloasso-report
```

2. **Créer un environnement virtuel :**

Il est recommandé d'utiliser un environnement virtuel pour gérer les dépendances du projet.

```bash
python3 -m venv venv
source venv/bin/activate  # Sous Windows : venv\Scripts\activate
```

3. **Installer les dépendances :**

```bash
pip install -r requirements.txt
```

## Configuration

Avant d'exécuter le script, vous devez configurer le fichier `config.ini` :

```ini
[helloasso]
client_id = VOTRE_CLIENT_ID
client_secret = VOTRE_CLIENT_SECRET
organization_slug = VOTRE_ORGANIZATION_SLUG
operation = NOM_DE_L_OPERATION

[smtp]
server = smtp.example.com
port = 465
user = votre_email@example.com
password = votre_mot_de_passe

[email]
recipient = destinataire@example.com

[products]
Produit1 = prix_de_vente,cout_de_revient
Produit2 = prix_de_vente,cout_de_revient
```

## Utilisation

1. **Exécuter le script :**

```bash
python main.py
```

2. **Résultats :**

- Les rapports sont enregistrés au format CSV.
- Un email est envoyé avec les statistiques détaillées.

## Contribution

Les contributions sont les bienvenues ! Veuillez soumettre une Pull Request avec une description claire des modifications.

## Licence

Ce projet est sous licence MIT. Consultez le fichier `LICENSE` pour plus de détails.
