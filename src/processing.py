from collections import defaultdict
from decimal import Decimal, InvalidOperation
import logging
import unicodedata
import re
from rich.progress import Progress

from src.config import app_config

logger = logging.getLogger("rich")

def normalize_parrain_code(code):
    """Normalise le code parrain pour une meilleure correspondance."""
    # Enlever les accents
    code = ''.join(
        c for c in unicodedata.normalize('NFD', code)
        if unicodedata.category(c) != 'Mn'
    )
    # Tout en majuscules
    code = code.upper()
    # Remplacer multiples espaces par un espace unique
    code = re.sub(r'\s+', ' ', code).strip()
    # Correction du pattern "4E B" → "4B"
    match = re.search(r'([0-9])E\s+([A-Z])$', code)
    if match:
        code = code[:match.start()] + match.group(1) + match.group(2)
    # Correction du pattern "5 J" → "5J"
    match = re.search(r'([0-9])\s+([A-Z])$', code)
    if match:
        code = code[:match.start()] + match.group(1) + match.group(2)
    # Extraire la classe
    match = re.search(r'([0-9][A-Z])$', code)
    if not match:
        return code
    classe = match.group(1)
    code = code[:match.start()].strip()
    tokens = code.split(' ')
    nom = tokens[0]
    return f"{nom} {classe}"

def normalize_product_name(product_name):
    """Normalise les noms de produits pour éviter les divergences."""
    return ''.join(
        c for c in unicodedata.normalize('NFD', product_name)
        if unicodedata.category(c) != 'Mn'
    ).lower().strip()

def calculate_sales_summary(orders):
    """Calcule le résumé des ventes pour une liste de commandes."""
    sales_summary = defaultdict(lambda: {
        'quantity': 0,
        'revenue': Decimal('0.00'),
        'profit': Decimal('0.00'),
        'buyers': set()
    })
    total_revenue = Decimal('0.00')
    total_profit = Decimal('0.00')

    with Progress() as progress:
        task = progress.add_task("[cyan]Calcul du résumé des ventes...", total=len(orders))
        for order in orders:
            payer_email = order.get("payer", {}).get("email")
            if not payer_email:
                progress.update(task, advance=1)
                continue

            for item in order.get("items", []):
                product_name = normalize_product_name(item.get("name", ""))
                quantity = item.get("quantity", 1)

                amount_info = item.get('amount', {})
                unit_price_cents = amount_info.get('total', 0) if isinstance(amount_info, dict) else amount_info

                try:
                    unit_price = Decimal(str(unit_price_cents)) / 100
                except (InvalidOperation, ValueError, TypeError) as e:
                    logger.error(f"Erreur de conversion du montant '{unit_price_cents}': {e}")
                    continue

                total_price = unit_price * quantity

                if product_name in app_config.products_prices:
                    sales_summary[product_name]['quantity'] += quantity
                    sales_summary[product_name]['revenue'] += total_price
                    sales_summary[product_name]['buyers'].add(payer_email)

                    profit = (app_config.products_prices[product_name] - app_config.product_costs[product_name]) * quantity
                    sales_summary[product_name]['profit'] += profit

                    total_revenue += total_price
                    total_profit += profit
            progress.update(task, advance=1)

    for product, data in sales_summary.items():
        data['buyers'] = len(data['buyers'])

    return sales_summary, total_revenue, total_profit

def get_best_seller(orders):
    """Détermine le meilleur vendeur basé sur les codes parrains."""
    parrain_sales = defaultdict(lambda: {'quantity': 0, 'revenue': Decimal('0.00')})
    parrain_product_name_normalized = normalize_product_name(app_config.parrain_product_name)

    with Progress() as progress:
        task = progress.add_task("[cyan]Calcul du meilleur vendeur...", total=len(orders))
        for order in orders:
            parrain_code = None
            for item in order.get("items", []):
                if normalize_product_name(item.get("name", "")) == parrain_product_name_normalized:
                    custom_fields = item.get("customFields", [])
                    if custom_fields:
                        parrain_code_raw = custom_fields[0].get("answer", "").strip()
                        parrain_code = normalize_parrain_code(parrain_code_raw)
                    break

            if parrain_code:
                for item in order.get("items", []):
                    if normalize_product_name(item.get("name", "")) == parrain_product_name_normalized:
                        continue

                    amount_info = item.get("amount", {})
                    unit_price_cents = amount_info.get('total', 0) if isinstance(amount_info, dict) else amount_info

                    try:
                        unit_price = Decimal(str(unit_price_cents)) / 100
                    except (InvalidOperation, ValueError, TypeError) as e:
                        logger.error(f"Erreur de conversion du montant '{unit_price_cents}': {e}")
                        continue

                    quantity = item.get("quantity", 1)
                    total_price = unit_price * quantity

                    parrain_sales[parrain_code]['quantity'] += quantity
                    parrain_sales[parrain_code]['revenue'] += total_price
            progress.update(task, advance=1)

    return parrain_sales

def aggregate_sales_by_date(orders):
    """Aggrège les ventes par date."""
    sales_per_day = defaultdict(lambda: {'revenue': 0, 'order_count': 0})
    for order in orders:
        date_str = order['date'][:10]
        total_order_amount = order['amount']['total']
        sales_per_day[date_str]['revenue'] += total_order_amount
        sales_per_day[date_str]['order_count'] += 1
    return sales_per_day