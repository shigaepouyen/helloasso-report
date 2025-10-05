import logging
import requests
import smtplib
from rich.logging import RichHandler
from rich.console import Console

# Importations depuis les nouveaux modules
from src.config import app_config
from src.api import get_access_token, get_orders
from src.processing import (
    calculate_sales_summary,
    get_best_seller,
    aggregate_sales_by_date
)
from src.reporting import (
    save_orders_to_csv,
    save_summary_to_csv,
    plot_sales_over_time,
    log_sales_summary,
    log_daily_sales,
    log_parrain_sales,
    send_email
)

# Initialisation de la console Rich
console = Console()
console.clear()

# Configuration du journal (logging) avec RichHandler
logging.basicConfig(
    level="INFO",  # Niveau de log par défaut
    format="%(message)s",
    datefmt="[%X]",
    handlers=[RichHandler(rich_tracebacks=True, console=console)]
)

logger = logging.getLogger("rich")

def main():
    """Point d'entrée principal du script."""
    try:
        # 1. Authentification et récupération des données
        logger.info("Récupération du jeton d'accès...")
        access_token = get_access_token()

        logger.info("Récupération des commandes...")
        orders = get_orders(access_token)
        num_orders = len(orders)

        # 2. Traitement des données
        logger.info("Calcul des ventes...")
        summary, total_revenue, total_profit = calculate_sales_summary(orders)

        logger.info("Agrégation des ventes par date...")
        sales_per_day = aggregate_sales_by_date(orders)

        logger.info("Détermination du meilleur vendeur...")
        parrain_sales = get_best_seller(orders)

        # 3. Génération des rapports
        logger.info("Enregistrement des commandes dans un fichier CSV...")
        save_orders_to_csv(orders)

        logger.info("Enregistrement du résumé des ventes dans un fichier CSV...")
        save_summary_to_csv(summary, total_revenue, total_profit)

        logger.info("Génération du graphique des ventes...")
        plot_sales_over_time(sales_per_day)

        # 4. Affichage des résultats dans la console
        logger.info("Affichage du résumé des ventes...")
        log_sales_summary(summary, total_revenue, total_profit, num_orders)

        logger.info("Affichage des ventes quotidiennes...")
        log_daily_sales(sales_per_day)

        logger.info("Affichage des ventes par code parrain...")
        log_parrain_sales(parrain_sales)

        # 5. Envoi de l'e-mail
        logger.info("Envoi du rapport par e-mail...")
        send_email(
            summary,
            parrain_sales,
            app_config.email['recipient'],
            num_orders,
            total_revenue,
            total_profit,
            sales_per_day
        )

        logger.info("Le script s'est terminé avec succès.")

    except ValueError as e:
        logger.error(f"Erreur de configuration : {e}", exc_info=False)
    except requests.exceptions.RequestException as e:
        logger.error(f"Erreur de communication avec l'API HelloAsso : {e}", exc_info=False)
    except smtplib.SMTPException as e:
        logger.error(f"Erreur lors de l'envoi de l'e-mail : {e}", exc_info=False)
    except Exception as e:
        logger.error(f"Une erreur critique et inattendue est survenue : {e}", exc_info=True)

if __name__ == "__main__":
    main()