import configparser
import os
from decimal import Decimal, InvalidOperation
import logging

# Configuration du journal (logging)
logger = logging.getLogger("rich")

def load_config():
    """Charge la configuration depuis le fichier config.ini."""
    config = configparser.ConfigParser()
    script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    config_file_path = os.path.join(script_dir, 'config.ini')
    config.read(config_file_path)
    return config

def get_helloasso_config(config):
    """Récupère la configuration de l'API HelloAsso."""
    return {
        'client_id': config.get('helloasso', 'client_id'),
        'client_secret': config.get('helloasso', 'client_secret'),
        'organization_slug': config.get('helloasso', 'organization_slug'),
        'operation': config.get('helloasso', 'operation')
    }

def get_smtp_config(config):
    """Récupère la configuration SMTP."""
    return {
        'server': config.get('smtp', 'server'),
        'port': config.getint('smtp', 'port'),
        'user': config.get('smtp', 'user'),
        'password': config.get('smtp', 'password'),
    }

def get_email_config(config):
    """Récupère la configuration de l'e-mail."""
    return {
        'recipient': config.get('email', 'recipient')
    }

def get_product_config(config):
    """Récupère la configuration des produits."""
    products_prices = {}
    product_costs = {}
    for product_name in config['products']:
        prices = config.get('products', product_name).split(',')
        if len(prices) == 2:
            try:
                sale_price = Decimal(prices[0].strip())
                cost_price = Decimal(prices[1].strip())
                products_prices[product_name] = sale_price
                product_costs[product_name] = cost_price
            except (InvalidOperation, ValueError) as e:
                logger.error(f"Erreur de conversion des prix pour le produit {product_name} : {e}")
        else:
            logger.warning(f"Format incorrect pour le produit {product_name} dans le fichier de configuration.")
    return products_prices, product_costs

def get_parrain_config(config):
    """Récupère la configuration du produit parrain."""
    return config.get('parameters', 'parrain_product_name')

def validate_config(config):
    """Valide que toutes les clés de configuration nécessaires sont présentes."""
    required = {
        'helloasso': ['client_id', 'client_secret', 'organization_slug', 'operation'],
        'smtp': ['server', 'port', 'user', 'password'],
        'email': ['recipient'],
        'parameters': ['parrain_product_name']
    }
    for section, keys in required.items():
        if not config.has_section(section):
            raise ValueError(f"La section [{section}] est manquante dans config.ini.")
        for key in keys:
            if not config.has_option(section, key) or not config.get(section, key):
                raise ValueError(f"L'option '{key}' est manquante ou vide dans la section [{section}] de config.ini.")
    if not config.has_section('products'):
        raise ValueError("La section [products] est manquante dans config.ini.")

def get_cache_config(config):
    """Récupère la configuration du cache."""
    if not config.has_section('cache'):
        return {'enabled': False, 'max_age_hours': 1}
    return {
        'enabled': config.getboolean('cache', 'enabled', fallback=False),
        'max_age_hours': config.getint('cache', 'max_age_hours', fallback=1)
    }

class AppConfig:
    """Classe de configuration pour l'application."""
    def __init__(self):
        config = load_config()
        validate_config(config)  # Valider la configuration au démarrage
        self.helloasso = get_helloasso_config(config)
        self.smtp = get_smtp_config(config)
        self.email = get_email_config(config)
        self.products_prices, self.product_costs = get_product_config(config)
        self.parrain_product_name = get_parrain_config(config)
        self.cache = get_cache_config(config)

        self.api_base_url = "https://api.helloasso.com/v5"
        self.auth_url = "https://api.helloasso.com/oauth2/token"
        self.script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.token_file = os.path.join(self.script_dir, 'token.json')
        self.cache_file = os.path.join(self.script_dir, 'orders_cache.json')

# Instance globale de la configuration
app_config = AppConfig()