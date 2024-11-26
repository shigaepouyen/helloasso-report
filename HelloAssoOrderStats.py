import requests
from collections import defaultdict
from datetime import datetime, timezone, timedelta
from decimal import Decimal, InvalidOperation
import logging
import configparser
import os
import json
from tabulate import tabulate
import matplotlib.pyplot as plt
import pandas as pd
from dateutil import parser
import unicodedata
import colorlog
import csv
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage

# Effacer l'écran du terminal au lancement
os.system('cls' if os.name == 'nt' else 'clear')

# Configuration du journal (logging) avec couleurs
log_colors = {
    'DEBUG': 'cyan',
    'INFO': 'green',
    'WARNING': 'yellow',
    'ERROR': 'red',
    'CRITICAL': 'red,bg_white',
}

formatter = colorlog.ColoredFormatter(
    "%(log_color)s%(levelname)s: %(message)s",
    log_colors=log_colors
)

handler = logging.StreamHandler()
handler.setFormatter(formatter)

logger = logging.getLogger()
logger.addHandler(handler)
logger.setLevel(logging.ERROR)  # Changez à DEBUG pour plus de détails

# Lecture du fichier de configuration
config = configparser.ConfigParser()
script_dir = os.path.dirname(os.path.abspath(__file__))
config_file_path = os.path.join(script_dir, 'config.ini')
config.read(config_file_path)

# Configuration API HelloAsso
config_helloasso = {
    'client_id': config.get('helloasso', 'client_id'),
    'client_secret': config.get('helloasso', 'client_secret'),
    'organization_slug': config.get('helloasso', 'organization_slug'),
    'operation': config.get('helloasso', 'operation')
}

API_BASE_URL = "https://api.helloasso.com/v5"
AUTH_URL = "https://api.helloasso.com/oauth2/token"

# Configuration SMTP
config_smtp = {
    'server': config.get('smtp', 'server'),
    'port': config.getint('smtp', 'port'),
    'user': config.get('smtp', 'user'),
    'password': config.get('smtp', 'password'),
}

# Adresse e-mail du destinataire
recipient_email = config.get('email', 'recipient')

# Produits avec prix de vente et coût
PRODUCTS_PRICES = {}
PRODUCT_COSTS = {}

for product_name in config['products']:
    prices = config.get('products', product_name).split(',')
    if len(prices) == 2:
        try:
            sale_price = Decimal(prices[0].strip())
            cost_price = Decimal(prices[1].strip())
            PRODUCTS_PRICES[product_name] = sale_price
            PRODUCT_COSTS[product_name] = cost_price
        except (InvalidOperation, ValueError) as e:
            logger.error(f"Erreur de conversion des prix pour le produit {product_name} : {e}")
    else:
        logger.warning(f"Format incorrect pour le produit {product_name} dans le fichier de configuration.")

def normalize_product_name(product_name):
    """Normalise les noms de produits pour éviter les divergences."""
    # Supprimer les accents et mettre en minuscules
    normalized_name = ''.join(
        c for c in unicodedata.normalize('NFD', product_name)
        if unicodedata.category(c) != 'Mn'
    ).lower().strip()
    return normalized_name

def get_access_token():
    """Récupère un jeton d'accès OAuth2 pour l'API HelloAsso."""
    token_file = os.path.join(script_dir, 'token.json')
    token_data = {}

    if os.path.exists(token_file):
        with open(token_file, 'r') as f:
            token_data = json.load(f)
        access_token = token_data.get('access_token')
        refresh_token = token_data.get('refresh_token')
        expires_at = token_data.get('expires_at')

        if expires_at and datetime.utcnow().timestamp() < expires_at:
            logger.info("Utilisation de l'access token existant.")
            return access_token
        elif refresh_token:
            logger.info("Rafraîchissement de l'access token...")
            data = {
                "grant_type": "refresh_token",
                "client_id": config_helloasso['client_id'],
                "refresh_token": refresh_token
            }
            response = requests.post(AUTH_URL, data=data)
            response.raise_for_status()
            new_token_data = response.json()
            return save_access_token(new_token_data)

    logger.info("Obtention d'un nouvel access token...")
    data = {
        "grant_type": "client_credentials",
        "client_id": config_helloasso['client_id'],
        "client_secret": config_helloasso['client_secret']
    }
    response = requests.post(AUTH_URL, data=data)
    response.raise_for_status()
    return save_access_token(response.json())

def save_access_token(token_data):
    """Sauvegarde le jeton d'accès et retourne l'access token."""
    access_token = token_data["access_token"]
    token_data["expires_at"] = datetime.utcnow().timestamp() + token_data["expires_in"]
    token_file = os.path.join(script_dir, 'token.json')
    with open(token_file, 'w') as f:
        json.dump(token_data, f)
    return access_token

def get_orders(access_token):
    """Récupère toutes les commandes de l'organisation avec pagination."""
    url = f"{API_BASE_URL}/organizations/{config_helloasso['organization_slug']}/orders"
    headers = {"Authorization": f"Bearer {access_token}"}
    params = {"pageIndex": 1, "pageSize": 20, "withDetails": True}
    all_orders = []

    while True:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        data = response.json()

        orders = data.get("data", [])
        logger.debug(f"Récupération des commandes de la page {params['pageIndex']} : {len(orders)} commandes")
        all_orders.extend(orders)

        if params["pageIndex"] >= data["pagination"]["totalPages"]:
            break
        params["pageIndex"] += 1

    # Vérifier la structure des premières commandes
    if all_orders:
        first_order = all_orders[0]
        if isinstance(first_order, dict):
            logger.debug("Les commandes sont des dictionnaires comme attendu.")
        else:
            logger.error(f"Les commandes ne sont pas des dictionnaires : {first_order}")

    return all_orders

def calculate_sales_summary(orders):
    """Calcule le résumé des ventes pour une liste de commandes, y compris les acheteurs distincts."""
    sales_summary = defaultdict(lambda: {
        'quantity': 0,
        'revenue': Decimal('0.00'),
        'profit': Decimal('0.00'),
        'buyers': set()  # Pour stocker les emails des acheteurs distincts
    })
    total_revenue = Decimal('0.00')
    total_profit = Decimal('0.00')

    for index, order in enumerate(orders, start=1):
        if not isinstance(order, dict):
            logger.error(f"L'ordre à l'index {index} n'est pas un dictionnaire : {order}")
            continue  # Passer à l'ordre suivant

        payer = order.get("payer")
        if not isinstance(payer, dict):
            logger.error(f"L'attribut 'payer' de l'ordre à l'index {index} n'est pas un dictionnaire : {payer}")
            continue  # Passer à l'ordre suivant

        payer_email = payer.get("email")
        if not isinstance(payer_email, str):
            logger.error(f"L'email du payeur dans l'ordre à l'index {index} est invalide : {payer_email}")
            continue  # Passer à l'ordre suivant

        for item_index, item in enumerate(order.get("items", []), start=1):
            if not isinstance(item, dict):
                logger.error(f"L'article à l'index {item_index} dans l'ordre {order.get('id', 'N/A')} n'est pas un dictionnaire : {item}")
                continue  # Passer à l'article suivant

            product_name = normalize_product_name(item.get("name", ""))
            quantity = item.get("quantity", 1)

            # Récupérer le montant de l'article
            amount_info = item.get('amount', {})
            unit_price_cents = 0

            if isinstance(amount_info, dict):
                unit_price_cents = amount_info.get('total', 0)
            elif isinstance(amount_info, int):
                unit_price_cents = amount_info
                logger.warning(f"L'attribut 'amount' de l'article à l'index {item_index} dans l'ordre {order.get('id', 'N/A')} est un entier : {unit_price_cents}. Traitement comme 'total' en centimes.")
            else:
                logger.error(f"L'attribut 'amount' de l'article à l'index {item_index} dans l'ordre {order.get('id', 'N/A')} est d'un type inattendu : {type(amount_info).__name__}. Ignoré.")
                continue  # Passer à l'article suivant

            try:
                unit_price = Decimal(str(unit_price_cents)) / 100
            except (InvalidOperation, ValueError, TypeError) as e:
                logger.error(f"Erreur lors de la conversion du montant de l'article '{unit_price_cents}' en Decimal : {e}")
                continue  # Passer à l'article suivant

            total_price = unit_price * quantity

            # Normaliser le nom du produit
            normalized_product_name = product_name

            if normalized_product_name in PRODUCTS_PRICES:
                sales_summary[normalized_product_name]['quantity'] += quantity
                sales_summary[normalized_product_name]['revenue'] += total_price
                sales_summary[normalized_product_name]['buyers'].add(payer_email)  # Ajouter l'email du client

                # Calcul du bénéfice
                profit = (PRODUCTS_PRICES[normalized_product_name] - PRODUCT_COSTS[normalized_product_name]) * quantity
                sales_summary[normalized_product_name]['profit'] += profit

                # Mise à jour des totaux globaux
                total_revenue += total_price
                total_profit += profit

    # Convertir les ensembles de buyers en leur nombre
    for product, data in sales_summary.items():
        data['buyers'] = len(data['buyers'])

    # Loguer le résumé des ventes pour vérification
    logger.debug(f"Résumé des ventes : {dict(sales_summary)}")
    return sales_summary, total_revenue, total_profit

def save_summary_to_csv(summary, total_revenue, total_profit):
    """Sauvegarde le résumé des ventes dans un fichier CSV, trié par quantité vendue."""
    # Trier les produits par quantité (ordre décroissant)
    sorted_summary = sorted(summary.items(), key=lambda x: x[1]['quantity'], reverse=True)

    csv_file = os.path.join(script_dir, 'sales_summary.csv')
    with open(csv_file, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow([
            "Produit", "Quantité", "Chiffre d'affaires (€)", "Bénéfice (€)", 
            "Nombre d'acheteurs", "Moyenne produits/acheteur"
        ])
        for product, data in sorted_summary:
            avg_per_buyer = round(data['quantity'] / data['buyers'], 2) if data['buyers'] > 0 else 0
            writer.writerow([
                product,
                data['quantity'],
                round(data['revenue'], 2),
                round(data['profit'], 2),
                data['buyers'],
                avg_per_buyer
            ])
        writer.writerow(["Total", "", round(total_revenue, 2), round(total_profit, 2), "", ""])
    logger.info(f"Le résumé des ventes a été enregistré dans {csv_file}.")

def get_order_details(order_id, access_token):
    """Récupère les détails d'une commande via l'API HelloAsso."""
    url = f"{API_BASE_URL}/orders/{order_id}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "accept": "application/json"
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

def get_best_seller(orders, access_token):
    """Détermine le meilleur vendeur basé sur le nombre total de produits et chiffre d'affaires par code parrain."""
    parrain_sales = defaultdict(lambda: {'quantity': 0, 'revenue': Decimal('0.00')})  # Initialisation avec quantité et revenu

    for order in orders:
        order_id = order.get("id")
        if not order_id:
            logger.error("Commande sans identifiant. Ignorée.")
            continue

        try:
            order_details = get_order_details(order_id, access_token)
        except Exception as e:
            logger.error(f"Erreur lors de la récupération des détails de la commande {order_id} : {e}")
            continue

        # Identifier le code parrain dans la commande
        parrain_code = None
        for item in order_details.get("items", []):
            if normalize_product_name(item.get("name", "")) == normalize_product_name("J’ai un parrain – soutenez un élève !"):
                custom_fields = item.get("customFields", [])
                for field in custom_fields:
                    if field.get("name", "").startswith("Vous avez été parrainé"):
                        parrain_code = field.get("answer", "").strip()
                        break

        # Si un code parrain est trouvé, additionner les autres produits
        if parrain_code:
            for item in order_details.get("items", []):
                if normalize_product_name(item.get("name", "")) != normalize_product_name("J’ai un parrain – soutenez un élève !"):
                    amount_info = item.get("amount", 0)
                    unit_price_cents = 0

                    if isinstance(amount_info, dict):
                        unit_price_cents = amount_info.get('total', 0)
                    elif isinstance(amount_info, int):
                        unit_price_cents = amount_info
                        logger.warning(f"L'attribut 'amount' de l'article dans la commande {order_id} est un entier : {unit_price_cents}. Traitement comme 'total' en centimes.")
                    else:
                        logger.error(f"L'attribut 'amount' de l'article dans la commande {order_id} est d'un type inattendu : {type(amount_info).__name__}. Ignoré.")
                        continue  # Passer à l'article suivant

                    try:
                        unit_price = Decimal(str(unit_price_cents)) / 100
                    except (InvalidOperation, ValueError, TypeError) as e:
                        logger.error(f"Erreur lors de la conversion du montant de l'article '{unit_price_cents}' en Decimal : {e}")
                        continue  # Passer à l'article suivant

                    quantity = item.get("quantity", 1)
                    total_price = unit_price * quantity

                    # Ajouter les quantités et le revenu pour le code parrain
                    parrain_sales[parrain_code]['quantity'] += quantity
                    parrain_sales[parrain_code]['revenue'] += total_price

    # Identifier le meilleur vendeur
    if parrain_sales:
        best_seller = max(parrain_sales.items(), key=lambda x: x[1]['quantity'])
        logger.info(f"Le meilleur vendeur est {best_seller[0]} avec {best_seller[1]['quantity']} produits vendus.")
    else:
        logger.info("Aucun parrainage trouvé.")

    return parrain_sales

def log_sales_summary(summary, total_revenue, total_profit, num_orders):
    """Affiche un résumé des ventes sous forme de tableau dans les logs, avec la date de génération."""
    # Ajouter la date et l'heure actuelles
    current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"\nRésumé des ventes au {current_date}, {num_orders} commandes avec les produits suivants :\n")

    # Trier les produits par quantité (ordre décroissant)
    sorted_summary = sorted(summary.items(), key=lambda x: x[1]['quantity'], reverse=True)

    table = []
    for product, data in sorted_summary:
        avg_per_buyer = round(data['quantity'] / data['buyers'], 2) if data['buyers'] > 0 else 0
        table.append([
            product,
            data['quantity'],
            f"{data['revenue']:.2f} €",
            f"{data['profit']:.2f} €",
            data['buyers'],
            avg_per_buyer  # Ajouter l'indicateur
        ])

    # Résumés totaux
    table.append([
        "Total",
        "",
        f"{total_revenue:.2f} €",
        f"{total_profit:.2f} €",
        "",
        ""  # Pas d'indicateur global
    ])

    # Affichage du tableau dans les logs
    headers = ["Produit", "Quantité", "Chiffre d'affaires", "Bénéfice", "Nombre d'acheteurs", "Moyenne produits/acheteur"]
    table_output = tabulate(table, headers=headers, tablefmt="grid")
    print(f"{table_output}\n")

def log_parrain_sales(parrain_sales):
    """Affiche un tableau des ventes par code parrain, incluant le chiffre d'affaires."""
    if parrain_sales:
        # Trier par quantité de produits vendus (ordre décroissant)
        sorted_parrain_sales = sorted(parrain_sales.items(), key=lambda x: x[1]['quantity'], reverse=True)

        table = []
        for parrain, data in sorted_parrain_sales:
            table.append([
                parrain,
                data['quantity'],
                f"{data['revenue']:.2f} €"  # Ajouter le chiffre d'affaires
            ])

        headers = ["Code Parrain", "Nombre de produits", "Chiffre d'affaires (€)"]
        # Spécifier l'alignement des colonnes : gauche, droite, droite
        table_output = tabulate(
            table,
            headers=headers,
            tablefmt="grid",
            colalign=("left", "right", "right")
        )
        print("\nRésumé des ventes par code parrain :\n" + table_output)
    else:
        logger.info("Aucun parrainage trouvé.")

def aggregate_sales_by_date(orders):
    """Agrège les ventes par date."""
    sales_per_day = defaultdict(Decimal)
    for order in orders:
        order_date_str = order.get("date", "")
        try:
            # Utiliser parser.parse pour gérer différents formats de date
            order_date = parser.parse(order_date_str).date()
        except (ValueError, TypeError) as e:
            logger.error(f"Erreur lors du parsing de la date {order_date_str}: {e}")
            continue  # Passer à la commande suivante en cas d'erreur

        total_order_amount = Decimal('0.00')
        for item in order.get("items", []):
            amount_info = item.get("amount", 0)
            unit_price_cents = 0

            if isinstance(amount_info, dict):
                unit_price_cents = amount_info.get('total', 0)
            elif isinstance(amount_info, int):
                unit_price_cents = amount_info
                logger.warning(f"L'attribut 'amount' de l'article dans l'ordre {order.get('id', 'N/A')} est un entier : {unit_price_cents}. Traitement comme 'total' en centimes.")
            else:
                logger.error(f"L'attribut 'amount' de l'article dans l'ordre {order.get('id', 'N/A')} est d'un type inattendu : {type(amount_info).__name__}. Ignoré.")
                continue  # Passer à l'article suivant

            try:
                unit_price = Decimal(str(unit_price_cents)) / 100
            except (InvalidOperation, ValueError, TypeError) as e:
                logger.error(f"Erreur lors de la conversion du montant de l'article '{unit_price_cents}' en Decimal : {e}")
                continue  # Passer à l'article suivant

            quantity = item.get("quantity", 1)
            total_price = unit_price * quantity
            total_order_amount += total_price

        sales_per_day[order_date] += total_order_amount
    return sales_per_day

def plot_sales_over_time(sales_per_day):
    """Génère un graphique du chiffre d'affaires par jour."""
    # Convertir les données en DataFrame pandas pour faciliter le tri
    df = pd.DataFrame({
        'Date': list(sales_per_day.keys()),
        'Chiffre d\'affaires': [float(sales_per_day[date]) for date in sales_per_day.keys()]
    })
    df = df.sort_values('Date')

    # Configurer le graphique
    plt.figure(figsize=(12, 6))
    plt.plot(df['Date'], df['Chiffre d\'affaires'], marker='o', linestyle='-', color='blue')
    plt.title('Chiffre d\'affaires par jour')
    plt.xlabel('Date')
    plt.ylabel('Chiffre d\'affaires (€)')
    plt.grid(True)

    # Rotation des étiquettes de l'axe X pour une meilleure lisibilité
    plt.xticks(rotation=45)

    # Ajuster les marges pour éviter que les étiquettes soient coupées
    plt.tight_layout()

    # Enregistrer le graphique avec une résolution adaptée
    plot_file = os.path.join(script_dir, 'sales_over_time.png')
    plt.savefig(plot_file, dpi=150)  # DPI ajusté pour un bon équilibre qualité/taille
    logger.info(f"Le graphique du chiffre d'affaires a été enregistré dans {plot_file}.")

    plt.close()  # Fermer la figure pour libérer la mémoire

def save_orders_to_csv(orders):
    """Sauvegarde les détails des commandes dans un fichier CSV pour la distribution."""
    # Liste des produits à exclure
    excluded_products = [
        normalize_product_name("j’ai un parrain – soutenez un élève !"),
        # Ajoutez d'autres produits à exclure si nécessaire
    ]

    # Obtenir la liste complète des produits pour créer les colonnes, en excluant les produits
    product_set = set()
    for order in orders:
        for item in order.get("items", []):
            product_name = normalize_product_name(item.get("name", ""))
            if product_name not in excluded_products:
                product_set.add(product_name)
    product_list = sorted(product_set)  # Trier les produits par ordre alphabétique

    # Initialiser une liste pour collecter les lignes du CSV
    rows = []

    for order in orders:
        # Récupérer la date de la commande
        order_date_str = order.get("date", "")
        try:
            order_date = parser.parse(order_date_str).strftime('%Y-%m-%d %H:%M:%S')
        except (ValueError, TypeError) as e:
            logger.error(f"Erreur lors du parsing de la date {order_date_str}: {e}")
            order_date = order_date_str  # Utiliser la chaîne originale en cas d'erreur

        # Récupérer les informations du client
        payer = order.get('payer', {})
        first_name = payer.get('firstName', '')
        last_name = payer.get('lastName', '')
        client_name = f"{first_name} {last_name}".strip()
        if not client_name:
            client_name = "Client Inconnu"

        # Récupérer le numéro de la commande
        order_number = order.get('id', 'N/A')

        # Récupérer le montant de la commande
        amount_info = order.get('amount', {})
        amount_cents = 0

        if isinstance(amount_info, dict):
            amount_cents = amount_info.get('total', 0)
        elif isinstance(amount_info, int):
            amount_cents = amount_info
            logger.warning(f"L'attribut 'amount' de la commande {order_number} est un entier : {amount_cents}. Traitement comme 'total' en centimes.")
        else:
            logger.error(f"L'attribut 'amount' de la commande {order_number} est d'un type inattendu : {type(amount_info).__name__}. Ignoré.")

        try:
            amount = Decimal(str(amount_cents)) / 100
        except (InvalidOperation, ValueError, TypeError) as e:
            logger.error(f"Erreur lors de la conversion du montant '{amount_cents}' en Decimal : {e}")
            amount = Decimal('0.00')

        # Initialiser les quantités de produits à zéro
        product_quantities = {product: 0 for product in product_list}

        for item in order.get("items", []):
            product_name = normalize_product_name(item.get("name", ""))
            if product_name in excluded_products:
                continue  # Exclure ce produit
            quantity = item.get("quantity", 1)
            if product_name in product_quantities:
                product_quantities[product_name] += quantity

        row = {
            'Date': order_date,
            'Client': client_name,
            'Numéro de la commande': order_number,
            'Montant (€)': f"{amount:.2f}",
            **product_quantities
        }

        rows.append(row)

    # Trier les lignes par le nom du client
    rows.sort(key=lambda x: x['Client'])

    # Écrire les lignes triées dans le fichier CSV
    csv_file = os.path.join(script_dir, 'orders.csv')
    with open(csv_file, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Date', 'Client', 'Numéro de la commande', 'Montant (€)'] + product_list
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)

    logger.info(f"Le fichier orders.csv a été enregistré dans {csv_file}.")

def generate_email_body(summary, parrain_sales, num_orders):
    """Génère le corps de l'e-mail avec les résumés des ventes."""
    lines = []
    lines.append(f"Nombre total de commandes : {num_orders}\n")
    lines.append("Résumé des ventes par produit :")
    lines.append("-" * 30)
    
    # Trier les produits par quantité vendue (ordre décroissant)
    sorted_summary = sorted(summary.items(), key=lambda x: x[1]['quantity'], reverse=True)
    
    for product, data in sorted_summary:
        avg_per_buyer = round(data['quantity'] / data['buyers'], 2) if data['buyers'] > 0 else 0
        lines.append(
            f"Produit : {product}\n"
            f"  Quantité vendue : {data['quantity']}\n"
            f"  Chiffre d'affaires : {data['revenue']:.2f} €\n"
            f"  Bénéfice : {data['profit']:.2f} €\n"
            f"  Nombre d'acheteurs : {data['buyers']}\n"
            f"  Moyenne produits/acheteur : {avg_per_buyer}\n"
            + "-" * 30
        )
    
    lines.append("\nRésumé des ventes par code parrain :")
    lines.append("-" * 30)
    
    if parrain_sales:
        # Trier par quantité de produits vendus (ordre décroissant)
        sorted_parrain_sales = sorted(parrain_sales.items(), key=lambda x: x[1]['quantity'], reverse=True)
        for parrain, data in sorted_parrain_sales:
            lines.append(
                f"Code Parrain : {parrain}\n"
                f"  Nombre de produits vendus : {data['quantity']}\n"
                f"  Chiffre d'affaires : {data['revenue']:.2f} €\n"
                + "-" * 30
            )
    else:
        lines.append("Aucun parrainage trouvé.")
    
    return "\n".join(lines)

def attach_file_to_email(msg, file_path, filename):
    """Attache un fichier au message e-mail."""
    try:
        with open(file_path, 'rb') as f:
            part = MIMEApplication(f.read(), Name=filename)
        part['Content-Disposition'] = f'attachment; filename="{filename}"'
        msg.attach(part)
    except Exception as e:
        logger.error(f"Erreur lors de l'attachement du fichier {filename} : {e}")

def send_email(summary, parrain_sales, recipient_email, num_orders, total_revenue, total_profit, sales_per_day):
    """Envoie le rapport par e-mail avec les pièces jointes et l'image intégrée dans le corps de l'email."""
    # Récupérer le nom de l'opération depuis la configuration
    operation_name = config_helloasso['operation']
    
    # Générer la date actuelle au format souhaité
    current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Générer l'objet de l'e-mail
    subject = f"[{operation_name}] Résumé des Ventes au {current_date}"
    
    # Générer le corps de l'e-mail en HTML et en texte brut
    email_body_html = generate_html_table(summary, parrain_sales, num_orders, total_revenue, total_profit)
    email_body_plain = generate_plain_text_body(summary, parrain_sales, num_orders, total_revenue, total_profit)
    
    # Créer le message e-mail avec une structure multipart
    msg = MIMEMultipart('related')  # Utiliser 'related' pour lier les images au HTML
    msg["From"] = config_smtp['user']
    msg["To"] = recipient_email
    msg["Subject"] = subject
    
    # Créer la partie alternative (texte brut et HTML)
    alternative_part = MIMEMultipart('alternative')
    msg.attach(alternative_part)
    
    # Attacher le corps de l'email en texte brut
    part1 = MIMEText(email_body_plain, "plain", "utf-8")
    alternative_part.attach(part1)
    
    # Attacher le corps de l'email en HTML
    part2 = MIMEText(email_body_html, "html", "utf-8")
    alternative_part.attach(part2)
    
    # Attacher les fichiers CSV
    attach_file_to_email(msg, os.path.join(script_dir, 'orders.csv'), 'orders.csv')
    attach_file_to_email(msg, os.path.join(script_dir, 'sales_summary.csv'), 'sales_summary.csv')
    
    # Intégrer le graphique dans le corps de l'email
    plot_file = os.path.join(script_dir, 'sales_over_time.png')
    
    try:
        with open(plot_file, 'rb') as img:
            mime = MIMEImage(img.read())
            mime.add_header('Content-ID', '<sales_plot>')  # Content-ID unique
            mime.add_header('Content-Disposition', 'inline', filename='sales_over_time.png')
            msg.attach(mime)
    except Exception as e:
        logger.error(f"Erreur lors de l'intégration du graphique dans l'email : {e}")
    
    try:
        # Créer le contexte SSL
        context = ssl.create_default_context()
        
        # Se connecter au serveur SMTP avec SSL
        with smtplib.SMTP_SSL(config_smtp['server'], config_smtp['port'], context=context) as server:
            server.login(config_smtp['user'], config_smtp['password'])
            server.send_message(msg)
            logger.info("E-mail envoyé avec succès.")
    except Exception as e:
        logger.error(f"Erreur lors de l'envoi de l'e-mail : {e}")

def generate_html_table(summary, parrain_sales, num_orders, total_revenue, total_profit):
    """Génère un tableau HTML pour le résumé des ventes avec les totaux et intègre le graphique."""
    current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Créer le tableau des ventes par produit
    sorted_summary = sorted(summary.items(), key=lambda x: x[1]['quantity'], reverse=True)
    table_summary = []
    for product, data in sorted_summary:
        avg_per_buyer = round(data['quantity'] / data['buyers'], 2) if data['buyers'] > 0 else 0
        table_summary.append([
            product,
            data['quantity'],
            f"{data['revenue']:.2f} €",
            f"{data['profit']:.2f} €",
            data['buyers'],
            avg_per_buyer
        ])
    
    # Ajouter la ligne des totaux
    table_summary.append([
        "Total",
        "",
        f"{total_revenue:.2f} €",
        f"{total_profit:.2f} €",
        "",
        ""
    ])
    
    headers_summary = ["Produit", "Quantité", "Chiffre d'affaires (€)", "Bénéfice (€)", "Nombre d'acheteurs", "Moyenne produits/acheteur"]
    # Aligner les colonnes : gauche, droite, droite, droite, droite, droite
    html_table_summary = tabulate(
        table_summary,
        headers=headers_summary,
        tablefmt="html",
        colalign=("left", "right", "right", "right", "right", "right")
    )
    
    # Créer le tableau des ventes par code parrain
    table_parrain = []
    if parrain_sales:
        sorted_parrain_sales = sorted(parrain_sales.items(), key=lambda x: x[1]['quantity'], reverse=True)
        for parrain, data in sorted_parrain_sales:
            table_parrain.append([
                parrain,
                data['quantity'],
                f"{data['revenue']:.2f} €"
            ])
    else:
        table_parrain.append(["Aucun parrainage trouvé.", "", ""])
    
    headers_parrain = ["Code Parrain", "Nombre de produits", "Chiffre d'affaires (€)"]
    # Aligner les colonnes : gauche, droite, droite
    html_table_parrain = tabulate(
        table_parrain,
        headers=headers_parrain,
        tablefmt="html",
        colalign=("left", "right", "right")
    )
    
    # Ajouter l'image intégrée dans le corps de l'e-mail
    html_image = """<img src="cid:sales_plot" alt="Chiffre d'affaires au fil du temps" style="width: 100%; height: auto;">"""
    
    # Générer le corps de l'email en HTML avec un conteneur
    html_body = f"""
    <html>
        <body>
            <div style="max-width: 800px; margin: auto;">
                <p>Bonjour,</p>
                <p>Voici le résumé de vos ventes au {current_date} avec {num_orders} commandes :</p>
                
                <h2>Résumé des Ventes par Produit</h2>
                {html_table_summary}
                
                <h2>Évolution du Chiffre d'Affaires</h2>
                {html_image}

                <h2>Résumé des Ventes par Code Parrain</h2>
                {html_table_parrain}
                
                <p>Cordialement,<br>
                   Votre Équipe</p>
            </div>
        </body>
    </html>
    """
    
    return html_body

def generate_plain_text_body(summary, parrain_sales, num_orders, total_revenue, total_profit):
    """Génère le corps de l'e-mail en texte brut avec des tableaux et les totaux."""
    current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    body = f"Bonjour,\n\n"
    body += f"Voici le résumé de vos ventes au {current_date} avec {num_orders} commandes :\n\n"
    
    # Résumé des ventes par produit
    body += "Résumé des Ventes par Produit :\n"
    sorted_summary = sorted(summary.items(), key=lambda x: x[1]['quantity'], reverse=True)
    table_summary = []
    for product, data in sorted_summary:
        avg_per_buyer = round(data['quantity'] / data['buyers'], 2) if data['buyers'] > 0 else 0
        table_summary.append([
            product,
            data['quantity'],
            f"{data['revenue']:.2f} €",
            f"{data['profit']:.2f} €",
            data['buyers'],
            avg_per_buyer
        ])
    
    # Ajouter la ligne des totaux
    table_summary.append([
        "Total",
        "",
        f"{total_revenue:.2f} €",
        f"{total_profit:.2f} €",
        "",
        ""
    ])
    
    headers_summary = ["Produit", "Quantité", "Chiffre d'affaires (€)", "Bénéfice (€)", "Nombre d'acheteurs", "Moyenne produits/acheteur"]
    # Spécifier l'alignement des colonnes
    plain_table_summary = tabulate(
        table_summary,
        headers=headers_summary,
        tablefmt="grid",
        colalign=("left", "right", "right", "right", "right", "right")
    )
    body += plain_table_summary + "\n\n"
    
    # Résumé des ventes par code parrain
    body += "Résumé des Ventes par Code Parrain :\n"
    table_parrain = []
    if parrain_sales:
        sorted_parrain_sales = sorted(parrain_sales.items(), key=lambda x: x[1]['quantity'], reverse=True)
        for parrain, data in sorted_parrain_sales:
            table_parrain.append([
                parrain,
                data['quantity'],
                f"{data['revenue']:.2f} €"
            ])
    else:
        table_parrain.append(["Aucun parrainage trouvé.", "", ""])
    headers_parrain = ["Code Parrain", "Nombre de produits", "Chiffre d'affaires (€)"]
    # Spécifier l'alignement des colonnes
    plain_table_parrain = tabulate(
        table_parrain,
        headers=headers_parrain,
        tablefmt="grid",
        colalign=("left", "right", "right")
    )
    body += plain_table_parrain + "\n\n"
    
    body += "Cordialement,\nVotre Équipe"
    
    return body

def main():
    try:
        logger.info("Récupération du jeton d'accès...")
        access_token = get_access_token()

        logger.info("Récupération des commandes...")
        orders = get_orders(access_token)
        num_orders = len(orders)  # Calcul du nombre total de commandes

        # Enregistrer les commandes dans un fichier CSV pour la distribution
        logger.info("Enregistrement des commandes dans un fichier CSV pour la distribution...")
        save_orders_to_csv(orders)

        # Calcul des ventes
        logger.info("Calcul des ventes...")
        summary, total_revenue, total_profit = calculate_sales_summary(orders)

        # Agréger les ventes par date
        sales_per_day = aggregate_sales_by_date(orders)

        # Générer le graphique des ventes
        logger.info("Génération du graphique des ventes...")
        plot_sales_over_time(sales_per_day)

        # Détermination du meilleur vendeur
        logger.info("Détermination du meilleur vendeur...")
        parrain_sales = get_best_seller(orders, access_token)

        # Affichage du résumé des ventes
        logger.info("Affichage du résumé des ventes...")
        log_sales_summary(summary, total_revenue, total_profit, num_orders)

        # Enregistrer le résumé des ventes dans un fichier CSV
        logger.info("Enregistrement du résumé des ventes dans un fichier CSV...")
        save_summary_to_csv(summary, total_revenue, total_profit)

        # Affichage des ventes par code parrain
        logger.info("Affichage des ventes par code parrain...")
        log_parrain_sales(parrain_sales)

        # Envoi du rapport par e-mail
        logger.info("Envoi du rapport par e-mail...")
        send_email(summary, parrain_sales, recipient_email, num_orders, total_revenue, total_profit, sales_per_day)

    except Exception as e:
        logger.error(f"Une erreur est survenue : {e}")

if __name__ == "__main__":
    main()
