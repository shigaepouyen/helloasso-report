import csv
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from datetime import datetime
import logging
import os
import matplotlib.pyplot as plt
from rich.console import Console
from rich.table import Table
from rich import box
from dateutil import parser

from src.config import app_config
from src.processing import normalize_product_name

logger = logging.getLogger("rich")
console = Console()

def save_summary_to_csv(summary, total_revenue, total_profit):
    """Sauvegarde le résumé des ventes dans un fichier CSV."""
    sorted_summary = sorted(summary.items(), key=lambda x: x[1]['quantity'], reverse=True)
    csv_file = os.path.join(app_config.script_dir, 'sales_summary.csv')
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

def save_orders_to_csv(orders):
    """Sauvegarde les détails des commandes dans un fichier CSV."""
    excluded_products = [normalize_product_name(app_config.parrain_product_name)]

    product_set = set()
    for order in orders:
        for item in order.get("items", []):
            product_name = normalize_product_name(item.get("name", ""))
            if product_name not in excluded_products:
                product_set.add(product_name)
    product_list = sorted(product_set)

    rows = []
    for order in orders:
        order_date_str = order.get("date", "")
        try:
            order_date = parser.parse(order_date_str).strftime('%Y-%m-%d %H:%M:%S')
        except (ValueError, TypeError):
            order_date = order_date_str

        payer = order.get('payer', {})
        first_name = payer.get('firstName', 'Prénom Inconnu').strip()
        last_name = payer.get('lastName', 'Nom Inconnu').strip()
        client_email = payer.get('email', 'Email Inconnu')
        order_number = order.get('id', 'N/A')

        amount_info = order.get('amount', {})
        amount_cents = amount_info.get('total', 0) if isinstance(amount_info, dict) else amount_info
        amount = amount_cents / 100

        product_quantities = {product: 0 for product in product_list}
        for item in order.get("items", []):
            product_name = normalize_product_name(item.get("name", ""))
            if product_name in product_quantities:
                product_quantities[product_name] += item.get("quantity", 1)

        row = {
            'Date': order_date, 'Nom': last_name, 'Prénom': first_name, 'Email': client_email,
            'Numéro de la commande': order_number, 'Montant (€)': f"{amount:.2f}",
            **product_quantities
        }
        rows.append(row)

    rows.sort(key=lambda x: x['Nom'])

    csv_file = os.path.join(app_config.script_dir, 'orders.csv')
    with open(csv_file, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Date', 'Nom', 'Prénom', 'Email', 'Numéro de la commande', 'Montant (€)'] + product_list
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    logger.info(f"Le fichier orders.csv a été enregistré dans {csv_file}.")

def plot_sales_over_time(sales_per_day):
    """Génère un graphique du chiffre d'affaires et du nombre de commandes par jour."""
    import pandas as pd
    dates = sorted(sales_per_day.keys(), key=lambda x: datetime.strptime(x, '%Y-%m-%d'))
    dates_datetime = [datetime.strptime(date_str, '%Y-%m-%d') for date_str in dates]

    revenues = [sales_per_day[date]['revenue'] / 100 for date in dates]
    order_counts = [sales_per_day[date].get('order_count', 0) for date in dates]

    df = pd.DataFrame({'Date': dates_datetime, 'Chiffre d\'affaires': revenues, 'Nombre de commandes': order_counts})

    fig, ax1 = plt.subplots(figsize=(12, 6))

    color = 'tab:blue'
    ax1.set_xlabel('Date')
    ax1.set_ylabel('Chiffre d\'affaires (€)', color=color)
    ax1.plot(df['Date'], df['Chiffre d\'affaires'], marker='o', linestyle='-', color=color, label='Chiffre d\'affaires')
    ax1.tick_params(axis='y', labelcolor=color)

    ax2 = ax1.twinx()
    color = 'tab:green'
    ax2.set_ylabel('Nombre de commandes', color=color)
    ax2.bar(df['Date'], df['Nombre de commandes'], color=color, alpha=0.3, label='Nombre de commandes')
    ax2.tick_params(axis='y', labelcolor=color)

    plt.title('Chiffre d\'affaires et nombre de commandes par jour')
    plt.xticks(rotation=45)
    fig.tight_layout()

    lines_labels = [ax.get_legend_handles_labels() for ax in [ax1, ax2]]
    lines, labels = [sum(lol, []) for lol in zip(*lines_labels)]
    fig.legend(lines, labels, loc='upper left')

    plot_file = os.path.join(app_config.script_dir, 'sales_over_time.png')
    plt.savefig(plot_file, dpi=150)
    plt.close()
    logger.info(f"Le graphique a été enregistré dans {plot_file}.")

def send_email(summary, parrain_sales, recipient_email, num_orders, total_revenue, total_profit, sales_per_day):
    """Envoie le rapport par e-mail."""
    operation_name = app_config.helloasso['operation']
    current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    subject = f"[{operation_name}] Résumé des Ventes au {current_date}"

    summary_html = generate_summary_html_table(summary, total_revenue, total_profit)
    daily_sales_table_html = generate_daily_sales_table_html(sales_per_day)
    parrain_sales_html = generate_parrain_sales_table_html(parrain_sales)

    email_body_html = f"""
    <html><body>
        <p>Bonjour,</p>
        <p>Résumé des ventes au {current_date}, {num_orders} commandes :</p>
        {summary_html}
        <p><img src="cid:sales_plot" alt="Graphique des ventes" style="max-width: 100%;"/></p>
        {daily_sales_table_html}
        {parrain_sales_html}
        <p>Cordialement,<br/>Votre équipe</p>
    </body></html>
    """

    msg = MIMEMultipart('related')
    msg["From"] = app_config.smtp['user']
    msg["To"] = recipient_email
    msg["Subject"] = subject

    alternative_part = MIMEMultipart('alternative')
    msg.attach(alternative_part)
    alternative_part.attach(MIMEText("Veuillez activer l'affichage HTML pour voir ce rapport.", "plain", "utf-8"))
    alternative_part.attach(MIMEText(email_body_html, "html", "utf-8"))

    for file in ['orders.csv', 'sales_summary.csv', 'sales_over_time.png']:
        attach_file_to_email(msg, os.path.join(app_config.script_dir, file), file)

    plot_file = os.path.join(app_config.script_dir, 'sales_over_time.png')
    try:
        with open(plot_file, 'rb') as img:
            mime = MIMEImage(img.read())
            mime.add_header('Content-ID', '<sales_plot>')
            msg.attach(mime)
    except Exception as e:
        logger.error(f"Erreur d'intégration du graphique : {e}")

    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(app_config.smtp['server'], app_config.smtp['port'], context=context) as server:
            server.login(app_config.smtp['user'], app_config.smtp['password'])
            server.send_message(msg)
            logger.info("E-mail envoyé avec succès.")
    except Exception as e:
        logger.error(f"Erreur d'envoi de l'e-mail : {e}")

def attach_file_to_email(msg, file_path, filename):
    """Attache un fichier à l'e-mail."""
    try:
        with open(file_path, 'rb') as f:
            part = MIMEApplication(f.read(), Name=filename)
        part['Content-Disposition'] = f'attachment; filename="{filename}"'
        msg.attach(part)
    except Exception as e:
        logger.error(f"Erreur d'attachement du fichier {filename} : {e}")

def generate_summary_html_table(summary, total_revenue, total_profit):
    """Génère un tableau HTML pour le résumé des ventes."""
    sorted_summary = sorted(summary.items(), key=lambda x: x[1]['quantity'], reverse=True)
    html = """<table border="1" cellpadding="5" style="border-collapse: collapse; width: 100%;">
    <thead><tr><th>Produit</th><th>Quantité</th><th>Chiffre d'affaires (€)</th><th>Bénéfice (€)</th><th>Acheteurs</th><th>Moyenne/acheteur</th></tr></thead>
    <tbody>"""
    for product, data in sorted_summary:
        avg_per_buyer = round(data['quantity'] / data['buyers'], 2) if data['buyers'] > 0 else 0
        html += f"""<tr><td>{product}</td><td align="right">{data['quantity']}</td><td align="right">{data['revenue']:.2f}</td>
        <td align="right">{data['profit']:.2f}</td><td align="right">{data['buyers']}</td><td align="right">{avg_per_buyer}</td></tr>"""
    html += f"""<tr style="font-weight: bold;"><td>Total</td><td></td><td align="right">{total_revenue:.2f}</td>
    <td align="right">{total_profit:.2f}</td><td></td><td></td></tr></tbody></table>"""
    return html

def generate_daily_sales_table_html(sales_per_day):
    """Génère un tableau HTML des ventes quotidiennes."""
    sorted_dates = sorted(sales_per_day.keys(), key=lambda x: datetime.strptime(x, '%Y-%m-%d'))
    html = """<h2>Ventes quotidiennes</h2><table border="1" cellpadding="5" style="border-collapse: collapse; width: 100%;">
    <thead><tr><th>Date</th><th>Commandes</th><th>Chiffre d'affaires (€)</th></tr></thead><tbody>"""
    for date in sorted_dates:
        revenue = sales_per_day[date]['revenue'] / 100
        html += f"""<tr><td>{date}</td><td align="right">{sales_per_day[date]['order_count']}</td><td align="right">{revenue:.2f}</td></tr>"""
    html += "</tbody></table>"
    return html

def generate_parrain_sales_table_html(parrain_sales):
    """Génère un tableau HTML des ventes par code parrain."""
    if not parrain_sales:
        return "<p>Aucun parrainage trouvé.</p>"
    sorted_sales = sorted(parrain_sales.items(), key=lambda x: x[1]['quantity'], reverse=True)
    html = """<h2>Ventes par code parrain</h2><table border="1" cellpadding="5" style="border-collapse: collapse; width: 100%;">
    <thead><tr><th>Code Parrain</th><th>Produits vendus</th><th>Chiffre d'affaires (€)</th></tr></thead><tbody>"""
    for parrain, data in sorted_sales:
        html += f"""<tr><td>{parrain}</td><td align="right">{data['quantity']}</td><td align="right">{data['revenue']:.2f}</td></tr>"""
    html += "</tbody></table>"
    return html

def log_sales_summary(summary, total_revenue, total_profit, num_orders):
    """Affiche un résumé des ventes dans la console."""
    console.print(f"\n[bold]Résumé des ventes au {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}, {num_orders} commandes :[/bold]\n")
    table = Table(show_header=True, header_style="bold magenta", box=box.SIMPLE)
    table.add_column("Produit")
    table.add_column("Quantité", justify="right")
    table.add_column("Chiffre d'affaires", justify="right")
    table.add_column("Bénéfice", justify="right")
    table.add_column("Acheteurs", justify="right")
    table.add_column("Moyenne/acheteur", justify="right")

    sorted_summary = sorted(summary.items(), key=lambda x: x[1]['quantity'], reverse=True)
    for product, data in sorted_summary:
        avg = round(data['quantity'] / data['buyers'], 2) if data['buyers'] > 0 else 0
        table.add_row(product, str(data['quantity']), f"{data['revenue']:.2f} €", f"{data['profit']:.2f} €", str(data['buyers']), str(avg))

    table.add_row("[bold]Total[/bold]", "", f"[bold]{total_revenue:.2f} €[/bold]", f"[bold]{total_profit:.2f} €[/bold]", "", "")
    console.print(table)

def log_daily_sales(sales_per_day):
    """Affiche les ventes quotidiennes dans la console."""
    table = Table(title="Ventes quotidiennes", show_header=True, header_style="bold magenta")
    table.add_column("Date")
    table.add_column("Commandes", justify="right")
    table.add_column("Chiffre d'affaires (€)", justify="right")

    sorted_dates = sorted(sales_per_day.keys(), key=lambda x: datetime.strptime(x, '%Y-%m-%d'))
    for date in sorted_dates:
        revenue = sales_per_day[date]['revenue'] / 100
        table.add_row(date, str(sales_per_day[date]['order_count']), f"{revenue:.2f}")
    console.print(table)

def log_parrain_sales(parrain_sales):
    """Affiche les ventes par code parrain dans la console."""
    if not parrain_sales:
        logger.info("Aucun parrainage trouvé.")
        return

    table = Table(title="Ventes par code parrain", show_header=True, header_style="bold magenta")
    table.add_column("Code Parrain")
    table.add_column("Produits vendus", justify="right")
    table.add_column("Chiffre d'affaires (€)", justify="right")

    sorted_sales = sorted(parrain_sales.items(), key=lambda x: x[1]['quantity'], reverse=True)
    for parrain, data in sorted_sales:
        table.add_row(parrain, str(data['quantity']), f"{data['revenue']:.2f}")
    console.print(table)