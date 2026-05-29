import pandas as pd
import numpy as np
from datetime import date, timedelta
import random

random.seed(42)
np.random.seed(42)

START = date(2025, 1, 1)
END   = date(2025, 12, 31)

CITIES = ["Lyon", "Bordeaux", "Nantes", "Toulouse", "Lille",
          "Marseille", "Strasbourg", "Rennes", "Montpellier", "Nice"]

SOURCES = {
    "LeBonCoin":   {"base": 55, "inbound": False, "cpl": 8,  "conv_q": 0.30, "conv_c": 0.14},
    "LaCentrale":  {"base": 30, "inbound": False, "cpl": 8,  "conv_q": 0.28, "conv_c": 0.13},
    "GBP":         {"base":  2, "inbound": True,  "cpl": 0,  "conv_q": 0.58, "conv_c": 0.27},
    "Facebook":    {"base":  3, "inbound": True,  "cpl": 0,  "conv_q": 0.46, "conv_c": 0.20},
    "Formulaire":  {"base":  1, "inbound": True,  "cpl": 5,  "conv_q": 0.52, "conv_c": 0.25},
    "Partenariat": {"base":  2, "inbound": True,  "cpl": 0,  "conv_q": 0.62, "conv_c": 0.30},
}

STATUSES = ["NRP", "En cours", "RDV planifié", "RDV effectué", "Signé", "Perdu"]
STATUS_W_LOW  = [0.22, 0.28, 0.18, 0.14, 0.10, 0.08]
STATUS_W_HIGH = [0.12, 0.22, 0.20, 0.18, 0.20, 0.08]

BRANDS = ["Renault","Peugeot","Citroën","Volkswagen","BMW","Toyota","Ford","Dacia","Opel","Mercedes"]

def inbound_ramp(month):
    if month <= 2:  return 0.10
    if month <= 4:  return 0.30
    if month <= 6:  return 0.65
    if month <= 9:  return 1.20
    return 1.80

def outbound_slight_decline(month):
    return 1.0 - 0.015 * (month - 1)

rows = []
lead_id = 1
day_range = (END - START).days + 1

for day_offset in range(day_range):
    current_date = START + timedelta(days=day_offset)
    month = current_date.month

    for source, cfg in SOURCES.items():
        base_daily = cfg["base"] / 30.0
        multiplier = inbound_ramp(month) if cfg["inbound"] else outbound_slight_decline(month)
        n_leads = np.random.poisson(base_daily * multiplier)

        for _ in range(n_leads):
            w = STATUS_W_HIGH if cfg["inbound"] else STATUS_W_LOW
            status = random.choices(STATUSES, weights=w)[0]

            if status == "NRP":         delay = random.randint(1, 3)
            elif status == "En cours":  delay = random.randint(1, 5)
            elif status in ("RDV planifié",): delay = random.randint(1, 7)
            elif status == "RDV effectué":    delay = random.randint(2, 10)
            elif status == "Signé":     delay = random.randint(1, 8) if cfg["inbound"] else random.randint(2, 14)
            else:                       delay = random.randint(3, 15)

            val = round(random.triangular(3500, 8500, 12000) / 100) * 100 if status == "Signé" else None
            margin_pct = round(random.uniform(0.06, 0.14), 3) if val else None
            margin = round(val * margin_pct) if val else None

            rows.append({
                "lead_id":       f"VM{lead_id:05d}",
                "date":          current_date.isoformat(),
                "month":         current_date.strftime("%Y-%m"),
                "month_num":     month,
                "source":        source,
                "inbound":       cfg["inbound"],
                "city":          random.choice(CITIES),
                "brand":         random.choice(BRANDS),
                "status":        status,
                "delay_days":    delay,
                "cpl_eur":       cfg["cpl"],
                "vehicle_value": val,
                "margin_eur":    margin,
                "margin_pct":    margin_pct,
            })
            lead_id += 1

df = pd.DataFrame(rows)
df["date"] = pd.to_datetime(df["date"])
df.to_csv("/home/claude/vroom-portfolio/data/leads_raw.csv", index=False)
print(f"Generated {len(df):,} leads")
print(df.groupby("source")["lead_id"].count().to_string())
