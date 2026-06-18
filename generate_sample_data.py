#!/usr/bin/env python3
"""
generate_sample_data.py
Generates a realistic 500-row Italian sales dataset for testing Sales Analyzer.
Run: python generate_sample_data.py
Output: sample_data.xlsx
"""

import random
from datetime import date, timedelta
import pandas as pd
import numpy as np

random.seed(42)
np.random.seed(42)

PRODUCTS = {
    "Elettronica": [
        "Smartphone Samsung", "iPhone 15", "Laptop Dell XPS", "Tablet iPad",
        "Cuffie Sony WH-1000XM5", "Smart TV LG 55\"", "Fotocamera Canon EOS",
        "Smartwatch Apple Watch", "Console PlayStation 5",
    ],
    "Abbigliamento": [
        "Giacca Invernale", "Jeans Levi's", "Scarpe Nike Air Max",
        "Borsa Gucci", "Cappotto Zara", "T-Shirt Polo Ralph Lauren",
        "Stivali UGG", "Occhiali Ray-Ban",
    ],
    "Casa e Cucina": [
        "Robot da Cucina Bimby", "Macchina del Caffè De'Longhi",
        "Set Pentole Lagostina", "Frigorifero Samsung", "Aspirapolvere Dyson",
        "Forno a Microonde Whirlpool", "Tostapane Philips",
    ],
    "Sport e Fitness": [
        "Tapis Roulant Technogym", "Bicicletta da Corsa Trek",
        "Casco da Sci Salomon", "Racchette da Tennis Wilson",
        "Zaino da Trekking Osprey", "Tappetino Yoga", "Manubri Regolabili",
    ],
    "Alimentari": [
        "Olio EVO Bio 5L", "Pasta Barilla 5kg", "Vino Chianti DOCG",
        "Parmigiano Reggiano 1kg", "Prosciutto Crudo San Daniele",
        "Caffè Illy 1kg", "Aceto Balsamico di Modena",
    ],
}

CITIES = [
    "Roma", "Milano", "Napoli", "Torino", "Palermo",
    "Genova", "Bologna", "Firenze", "Bari", "Catania",
    "Venezia", "Verona", "Messina", "Padova", "Trieste",
    "Brescia", "Taranto", "Modena", "Reggio Calabria", "Reggio Emilia",
    "Perugia", "Ravenna", "Livorno", "Cagliari", "Foggia",
    "Rimini", "Salerno", "Ferrara", "Sassari", "Bergamo",
]

CHANNELS = ["Online", "Negozio", "Marketplace"]

CITY_WEIGHTS = [
    20, 18, 10, 8, 5,
    5, 5, 5, 4, 4,
    3, 3, 2, 2, 2,
    2, 2, 2, 1, 1,
    1, 1, 1, 1, 1,
    1, 1, 1, 1, 1,
]

CHANNEL_WEIGHTS = [50, 30, 20]


def random_date(start: date, end: date) -> date:
    delta = (end - start).days
    return start + timedelta(days=random.randint(0, delta))


def seasonal_multiplier(d: date) -> float:
    """Higher sales in Nov-Dec (Black Friday, Christmas) and July-Aug."""
    month = d.month
    if month in (11, 12):
        return random.uniform(1.4, 2.0)
    if month in (6, 7, 8):
        return random.uniform(1.1, 1.5)
    if month in (1, 2):
        return random.uniform(0.6, 0.9)
    return random.uniform(0.9, 1.2)


def main():
    start_date = date(2022, 1, 1)
    end_date = date(2024, 12, 31)

    rows = []
    for _ in range(500):
        category = random.choice(list(PRODUCTS.keys()))
        product = random.choice(PRODUCTS[category])
        d = random_date(start_date, end_date)
        channel = random.choices(CHANNELS, weights=CHANNEL_WEIGHTS)[0]
        city = random.choices(CITIES, weights=CITY_WEIGHTS)[0]

        base_price = {
            "Elettronica": random.uniform(80, 1500),
            "Abbigliamento": random.uniform(20, 400),
            "Casa e Cucina": random.uniform(30, 600),
            "Sport e Fitness": random.uniform(15, 800),
            "Alimentari": random.uniform(5, 60),
        }[category]

        season = seasonal_multiplier(d)
        channel_multiplier = {"Online": 1.0, "Negozio": 1.1, "Marketplace": 0.95}[channel]

        quantity = max(1, int(np.random.poisson(3) * season))
        unit_price = round(base_price * channel_multiplier * random.uniform(0.85, 1.15), 2)
        revenue = round(unit_price * quantity, 2)

        rows.append({
            "Data": pd.Timestamp(d),
            "Prodotto": product,
            "Categoria": category,
            "Ricavo": revenue,
            "Quantità": quantity,
            "Canale": channel,
            "Città": city,
        })

    df = pd.DataFrame(rows)
    df = df.sort_values("Data").reset_index(drop=True)

    out_path = "sample_data.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Vendite")

    print(f"✓ Generato {len(df)} righe → {out_path}")
    print(f"  Periodo: {df['Data'].min().date()} — {df['Data'].max().date()}")
    print(f"  Ricavo totale: €{df['Ricavo'].sum():,.2f}")
    print(f"  Prodotti unici: {df['Prodotto'].nunique()}")
    print(f"  Canali: {', '.join(df['Canale'].unique())}")
    print(f"  Città: {df['Città'].nunique()}")


if __name__ == "__main__":
    main()
