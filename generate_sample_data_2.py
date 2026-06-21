#!/usr/bin/env python3
"""
generate_sample_data_2.py
Generates a second 300-row Italian sales dataset (2025) for testing multi-file load.
Includes blank Canale cells (filled with "Banco" by the app) and different products.
Run: python3 generate_sample_data_2.py
Output: sample_data_2.xlsx
"""

import random
from datetime import date, timedelta
import pandas as pd
import numpy as np

random.seed(99)
np.random.seed(99)

PRODUCTS = {
    "Gioielleria": [
        "Orologio Rolex Datejust", "Anello Diamante 0.5ct",
        "Bracciale Cartier", "Collana Perle", "Orecchini Tiffany",
        "Pendente Oro 18k", "Spilla Art Déco",
    ],
    "Libreria": [
        "Enciclopedia Treccani", "Romanzo Storico Italiano",
        "Manuale di Fotografia", "Atlas Geografico", "Bibbia Illustrata",
        "Corso di Cucina Vol.1", "Dizionario Garzanti",
    ],
    "Giocattoli": [
        "LEGO Technic 42150", "Barbie Deluxe", "Hot Wheels Pista",
        "Puzzle 1000pz Ravensburger", "Drone DJI Mini",
        "Gioco da Tavolo Scarabeo", "Triciclo Chicco",
    ],
    "Automotive": [
        "Navigatore TomTom", "Dashcam Full HD", "Coprisedili in Pelle",
        "Catene da Neve", "Kit Riparazione Gomme",
        "Aspirapolvere Auto", "Profumatore Abitacolo",
    ],
    "Cosmetica": [
        "Profumo Chanel No.5", "Crema Antirughe Lancôme",
        "Set Make-Up MAC", "Fondotinta Giorgio Armani",
        "Palette Ombretti Urban Decay", "Rossetto Dior",
        "Siero Vitamina C",
    ],
}

CITIES = [
    "Roma", "Milano", "Napoli", "Torino", "Bologna",
    "Firenze", "Venezia", "Palermo", "Genova", "Bari",
    "Verona", "Padova", "Trieste", "Cagliari", "Rimini",
]

# Channels — some rows will have a blank value to test Banco-fill
CHANNELS = ["Online", "Negozio", "Marketplace", "Banco", ""]

CHANNEL_WEIGHTS = [35, 25, 15, 15, 10]   # 10% blank

CITY_WEIGHTS = [20, 18, 10, 8, 7, 6, 5, 4, 4, 3, 3, 3, 3, 3, 3]


def random_date(start: date, end: date) -> date:
    return start + timedelta(days=random.randint(0, (end - start).days))


def seasonal_multiplier(d: date) -> float:
    month = d.month
    if month in (11, 12):
        return random.uniform(1.5, 2.2)
    if month in (6, 7, 8):
        return random.uniform(1.0, 1.4)
    if month in (1, 2):
        return random.uniform(0.5, 0.8)
    return random.uniform(0.85, 1.15)


def main():
    start_date = date(2025, 1, 1)
    end_date   = date(2025, 12, 31)

    rows = []
    for _ in range(300):
        category = random.choice(list(PRODUCTS.keys()))
        product  = random.choice(PRODUCTS[category])
        d        = random_date(start_date, end_date)
        channel  = random.choices(CHANNELS, weights=CHANNEL_WEIGHTS)[0]
        city     = random.choices(CITIES,   weights=CITY_WEIGHTS)[0]

        base_price = {
            "Gioielleria": random.uniform(150, 4000),
            "Libreria":    random.uniform(8, 120),
            "Giocattoli":  random.uniform(15, 350),
            "Automotive":  random.uniform(20, 280),
            "Cosmetica":   random.uniform(25, 300),
        }[category]

        season   = seasonal_multiplier(d)
        quantity = max(1, int(np.random.poisson(2) * season))
        price    = round(base_price * random.uniform(0.9, 1.1), 2)
        revenue  = round(price * quantity, 2)

        rows.append({
            "Data":      pd.Timestamp(d),
            "Prodotto":  product,
            "Categoria": category,
            "Ricavo":    revenue,
            "Quantità":  quantity,
            "Canale":    channel,      # some are blank → filled with "Banco" by the app
            "Città":     city,
        })

    df = pd.DataFrame(rows).sort_values("Data").reset_index(drop=True)

    out = "sample_data_2.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Vendite 2025")

    blank_ch = (df["Canale"] == "").sum()
    print(f"Generato {len(df)} righe → {out}")
    print(f"  Periodo : {df['Data'].min().date()} — {df['Data'].max().date()}")
    print(f"  Ricavo  : €{df['Ricavo'].sum():,.2f}")
    print(f"  Canali  : {sorted(df['Canale'].unique())}")
    print(f"  Celle vuote in Canale: {blank_ch} (verranno impostate a 'Banco' dall'app)")


if __name__ == "__main__":
    main()
