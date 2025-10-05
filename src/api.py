import requests
import json
import os
from datetime import datetime, timedelta
import logging
from rich.progress import Progress

from src.config import app_config

logger = logging.getLogger("rich")

def get_access_token():
    """Récupère un jeton d'accès OAuth2 pour l'API HelloAsso."""
    token_data = {}

    if os.path.exists(app_config.token_file):
        with open(app_config.token_file, 'r') as f:
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
                "client_id": app_config.helloasso['client_id'],
                "refresh_token": refresh_token
            }
            response = requests.post(app_config.auth_url, data=data)
            response.raise_for_status()
            new_token_data = response.json()
            return save_access_token(new_token_data)

    logger.info("Obtention d'un nouvel access token...")
    data = {
        "grant_type": "client_credentials",
        "client_id": app_config.helloasso['client_id'],
        "client_secret": app_config.helloasso['client_secret']
    }
    response = requests.post(app_config.auth_url, data=data)
    response.raise_for_status()
    return save_access_token(response.json())

def save_access_token(token_data):
    """Sauvegarde le jeton d'accès et retourne l'access token."""
    access_token = token_data["access_token"]
    token_data["expires_at"] = datetime.utcnow().timestamp() + token_data["expires_in"]
    with open(app_config.token_file, 'w') as f:
        json.dump(token_data, f)
    return access_token

def get_orders(access_token):
    """Récupère toutes les commandes, en utilisant un cache si disponible."""
    # Vérifier si le cache est activé et valide
    if app_config.cache['enabled'] and os.path.exists(app_config.cache_file):
        file_mod_time = datetime.fromtimestamp(os.path.getmtime(app_config.cache_file))
        if datetime.now() - file_mod_time < timedelta(hours=app_config.cache['max_age_hours']):
            logger.info("Utilisation du cache pour les commandes.")
            with open(app_config.cache_file, 'r', encoding='utf-8') as f:
                return json.load(f)

    # Si le cache n'est pas utilisé, récupérer depuis l'API
    logger.info("Récupération des commandes depuis l'API HelloAsso...")
    url = f"{app_config.api_base_url}/organizations/{app_config.helloasso['organization_slug']}/orders"
    headers = {"Authorization": f"Bearer {access_token}"}
    params = {"pageIndex": 1, "pageSize": 20, "withDetails": True}
    all_orders = []

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    data = response.json()
    total_pages = data["pagination"]["totalPages"]
    all_orders.extend(data.get("data", []))

    if total_pages > 1:
        with Progress() as progress:
            task = progress.add_task("[cyan]Téléchargement des commandes...", total=total_pages)
            progress.update(task, advance=1)

            for page_index in range(2, total_pages + 1):
                params["pageIndex"] = page_index
                response = requests.get(url, headers=headers, params=params)
                response.raise_for_status()
                data = response.json()
                all_orders.extend(data.get("data", []))
                progress.update(task, advance=1)

    # Sauvegarder les données dans le cache si activé
    if app_config.cache['enabled']:
        logger.info(f"Sauvegarde de {len(all_orders)} commandes dans le cache...")
        with open(app_config.cache_file, 'w', encoding='utf-8') as f:
            json.dump(all_orders, f, indent=4)

    return all_orders
