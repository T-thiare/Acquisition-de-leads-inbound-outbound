"""
Vroom Market — Visualisations analytiques Python
Génère 4 graphiques prêts pour portfolio GitHub
"""
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.patheffects as pe
import numpy as np
import warnings
warnings.filterwarnings("ignore")

df = pd.read_csv("/home/claude/vroom-portfolio/data/leads_raw.csv", parse_dates=["date"])
df["qualified"] = df["status"].isin(["RDV planifié","RDV effectué","Signé"])
df["closed"]    = df["status"] == "Signé"
df["rdv"]       = df["status"].isin(["RDV planifié","RDV effectué","Signé"])
df["in_out"]    = df["inbound"].map({True:"Inbound",False:"Outbound"})

NAVY   = "#11182C"
GOLD   = "#E0BC00"
GREEN  = "#1D9E75"
MUTED  = "#94A3B8"
RED    = "#EF4444"
AMBER  = "#F59E0B"
INBND  = "#10B981"
OUTBND = "#64748B"
BG     = "#F8FAFC"
MONTH_LABELS = ["Jan","Fév","Mar","Avr","Mai","Jun","Jul","Aoû","Sep","Oct","Nov","Déc"]
SOURCE_ORDER = ["LeBonCoin","LaCentrale","GBP","Facebook","Formulaire","Partenariat"]

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.grid": True,
    "grid.color": "#E5E7EB",
    "grid.linewidth": 0.6,
    "figure.facecolor": BG,
    "axes.facecolor": BG,
    "text.color": NAVY,
    "axes.labelcolor": NAVY,
    "xtick.color": "#6B7280",
    "ytick.color": "#6B7280",
})

# ═══════════════════════════════════════════════════════════════════════════
# FIGURE 1 — Évolution mensuelle inbound vs outbound
# ═══════════════════════════════════════════════════════════════════════════
monthly = df.groupby(["month_num","in_out"])["lead_id"].count().unstack(fill_value=0).reset_index()

fig, axes = plt.subplots(2, 1, figsize=(13, 9), facecolor=BG,
                          gridspec_kw={"height_ratios":[2,1],"hspace":0.08})

ax1 = axes[0]
months = list(range(1, 13))
out_vals = [monthly[monthly["month_num"]==m]["Outbound"].sum() for m in months]
in_vals  = [monthly[monthly["month_num"]==m]["Inbound"].sum()  for m in months]

x = np.arange(12)
b1 = ax1.bar(x, out_vals, color=OUTBND, alpha=0.85, label="Outbound (LBC + La Centrale)", zorder=3, width=0.6)
b2 = ax1.bar(x, in_vals,  bottom=out_vals, color=INBND, alpha=0.92, label="Inbound (GBP + FB + Form. + Part.)", zorder=3, width=0.6)

for i, (out, inb) in enumerate(zip(out_vals, in_vals)):
    total = out + inb
    if inb > 0:
        ax1.text(i, total + 0.5, str(inb), ha="center", va="bottom",
                fontsize=7.5, color=INBND, fontweight="bold")

ax1.set_xticks(x); ax1.set_xticklabels([])
ax1.set_ylabel("Nombre de leads", fontsize=10, labelpad=8)
ax1.set_ylim(0, max(out_vals) * 1.2)
ax1.legend(loc="upper left", frameon=True, fancybox=True, framealpha=0.9, fontsize=9)
ax1.set_title("VROOM MARKET — Évolution des leads par canal · Jan–Déc 2025",
              fontsize=13, fontweight="bold", color=NAVY, pad=14, loc="left")

# % inbound sparkline
ax2 = axes[1]
pct_in = [inb/(out+inb)*100 if (out+inb)>0 else 0 for out,inb in zip(out_vals,in_vals)]
ax2.fill_between(x, pct_in, alpha=0.15, color=INBND)
ax2.plot(x, pct_in, color=INBND, linewidth=2.5, marker="o", markersize=5, zorder=4)
ax2.axhline(15, color=GOLD, linewidth=1, linestyle="--", alpha=0.7, label="Seuil cible 15%")
for i, p in enumerate(pct_in):
    if i in [0, 5, 11]:
        ax2.annotate(f"{p:.1f}%", (i, p), textcoords="offset points", xytext=(0,8),
                    ha="center", fontsize=8, color=INBND, fontweight="bold")

ax2.set_xticks(x); ax2.set_xticklabels(MONTH_LABELS, fontsize=9)
ax2.set_ylabel("% Inbound", fontsize=9, labelpad=8)
ax2.set_ylim(0, max(pct_in)*1.4)
ax2.legend(loc="upper left", fontsize=8, frameon=False)
ax2.grid(axis="x", alpha=0)

fig.text(0.01, 0.01, "Source : Dataset simulé Vroom Market 2025 | Benchmarks: HubSpot Lead Gen Report, secteur automobile France",
         fontsize=7.5, color="#9CA3AF", style="italic")

plt.savefig("/home/claude/vroom-portfolio/dashboards/fig1_evolution_inbound_outbound.png",
            dpi=160, bbox_inches="tight", facecolor=BG)
plt.close()
print("Fig 1 saved")

# ═══════════════════════════════════════════════════════════════════════════
# FIGURE 2 — Performance comparative par source (bubble chart)
# ═══════════════════════════════════════════════════════════════════════════
src_perf = df.groupby("source").agg(
    total=("lead_id","count"),
    closing=("closed","mean"),
    qual=("qualified","mean"),
    delay=("delay_days","mean"),
    inbound=("inbound","first"),
    cpl=("cpl_eur","first"),
    margin=("margin_eur","sum")
).reset_index()

fig, ax = plt.subplots(figsize=(12, 7.5), facecolor=BG)

COLORS_SRC = {
    "LeBonCoin":   OUTBND,
    "LaCentrale":  "#475569",
    "GBP":         INBND,
    "Facebook":    "#8B5CF6",
    "Formulaire":  "#3B82F6",
    "Partenariat": GOLD,
}

for _, row in src_perf.iterrows():
    size = row["total"] * 4.5
    color = COLORS_SRC.get(row["source"], MUTED)
    edge = NAVY if row["inbound"] else "#CBD5E1"
    alpha = 0.88 if row["inbound"] else 0.60

    scatter = ax.scatter(row["qual"]*100, row["closing"]*100,
                         s=size, c=color, alpha=alpha,
                         edgecolors=edge, linewidths=1.8, zorder=5)

    offset_x = 1.5
    offset_y = 1.2
    name_short = row["source"]
    ax.annotate(f"{name_short}\n({row['total']} leads)",
                (row["qual"]*100 + offset_x, row["closing"]*100 + offset_y),
                fontsize=8.5, fontweight="bold" if row["inbound"] else "normal",
                color=NAVY, zorder=6)

ax.axvline(40, color="#E5E7EB", linewidth=1, linestyle="--")
ax.axhline(18, color="#E5E7EB", linewidth=1, linestyle="--")
ax.text(42, 18.5, "Zone performance", fontsize=7.5, color=MUTED, style="italic")

ax.set_xlabel("Taux de qualification (%)", fontsize=11, labelpad=8)
ax.set_ylabel("Taux de closing (%)", fontsize=11, labelpad=8)
ax.set_title("Performance par canal d'acquisition — Taille = volume de leads · 2025",
             fontsize=12, fontweight="bold", color=NAVY, pad=12, loc="left")

in_patch  = mpatches.Patch(color=INBND,  alpha=0.88, label="Canaux Inbound")
out_patch = mpatches.Patch(color=OUTBND, alpha=0.60, label="Canaux Outbound")
ax.legend(handles=[in_patch, out_patch], loc="lower right", fontsize=9, frameon=True)

fig.text(0.01, 0.01, "Taux de qualification = leads ayant atteint RDV+ | Taux de closing = deals signés / total leads",
         fontsize=7.5, color="#9CA3AF", style="italic")

plt.tight_layout()
plt.savefig("/home/claude/vroom-portfolio/dashboards/fig2_performance_par_source.png",
            dpi=160, bbox_inches="tight", facecolor=BG)
plt.close()
print("Fig 2 saved")

# ═══════════════════════════════════════════════════════════════════════════
# FIGURE 3 — Funnel de conversion comparatif inbound vs outbound
# ═══════════════════════════════════════════════════════════════════════════
def build_funnel(subset):
    n = len(subset)
    return {
        "Leads reçus":    n,
        "Contactés":      subset[subset["status"]!="NRP"].shape[0],
        "Qualifiés":      subset["qualified"].sum(),
        "RDV":            subset["rdv"].sum(),
        "Signés":         subset["closed"].sum(),
    }

in_df  = df[df["inbound"]==True]
out_df = df[df["inbound"]==False]
funnel_in  = build_funnel(in_df)
funnel_out = build_funnel(out_df)

stages = list(funnel_in.keys())
vals_in  = [funnel_in[s]  / funnel_in["Leads reçus"]  * 100 for s in stages]
vals_out = [funnel_out[s] / funnel_out["Leads reçus"] * 100 for s in stages]

fig, ax = plt.subplots(figsize=(13, 7), facecolor=BG)

y = np.arange(len(stages))
h = 0.35

bars_out = ax.barh(y + h/2, vals_out, h, color=OUTBND, alpha=0.75, label="Outbound")
bars_in  = ax.barh(y - h/2, vals_in,  h, color=INBND, alpha=0.88, label="Inbound")

for bar, val, raw_in, raw_out in zip(bars_in, vals_in,
                                      [funnel_in[s] for s in stages],
                                      [funnel_out[s] for s in stages]):
    ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2,
            f"{val:.1f}%  ({raw_in})", va="center", fontsize=8.5,
            color=INBND, fontweight="bold")

for bar, val, raw in zip(bars_out, vals_out, [funnel_out[s] for s in stages]):
    ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2,
            f"{val:.1f}%  ({raw})", va="center", fontsize=8.5, color=OUTBND)

ax.set_yticks(y); ax.set_yticklabels(stages, fontsize=10)
ax.set_xlabel("% des leads entrants", fontsize=10)
ax.set_xlim(0, 130)
ax.set_title("Funnel de conversion : Inbound vs Outbound · 2025",
             fontsize=12, fontweight="bold", color=NAVY, pad=12, loc="left")
ax.legend(fontsize=10, frameon=True, loc="lower right")
ax.invert_yaxis()
ax.grid(axis="y", alpha=0)

fig.text(0.01, 0.01,
         f"Inbound : {len(in_df)} leads total | Outbound : {len(out_df)} leads total",
         fontsize=8, color="#9CA3AF", style="italic")

plt.tight_layout()
plt.savefig("/home/claude/vroom-portfolio/dashboards/fig3_funnel_comparatif.png",
            dpi=160, bbox_inches="tight", facecolor=BG)
plt.close()
print("Fig 3 saved")

# ═══════════════════════════════════════════════════════════════════════════
# FIGURE 4 — ROI & Coût par lead
# ═══════════════════════════════════════════════════════════════════════════
roi_data = []
for src in SOURCE_ORDER:
    s = df[df["source"]==src]
    n = len(s)
    closed_n = s["closed"].sum()
    cpl = s["cpl_eur"].iloc[0]
    total_cost = n * cpl
    marge = s["margin_eur"].sum() if not s["margin_eur"].isna().all() else 0
    cpa = total_cost / closed_n if closed_n > 0 else 0
    roi_data.append({
        "source": src,
        "inbound": s["inbound"].iloc[0],
        "n": n,
        "cpa": cpa,
        "cpl": cpl,
        "marge": marge if not pd.isna(marge) else 0,
        "cost": total_cost,
    })
roi_df = pd.DataFrame(roi_data)

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6.5), facecolor=BG)

# ── Bar chart CPA ──
colors_bar = [INBND if r["inbound"] else OUTBND for _, r in roi_df.iterrows()]
bars = ax1.bar(roi_df["source"], roi_df["cpa"], color=colors_bar, alpha=0.85,
               width=0.6, zorder=3, edgecolor="white", linewidth=0.8)

for bar, (_, row) in zip(bars, roi_df.iterrows()):
    h_b = bar.get_height()
    label = f"{h_b:.0f} €" if h_b > 0 else "0 € *"
    color_txt = INBND if row["inbound"] else OUTBND
    ax1.text(bar.get_x() + bar.get_width()/2, h_b + 1,
             label, ha="center", va="bottom", fontsize=9, fontweight="bold", color=color_txt)

ax1.set_ylabel("Coût par acquisition (€)", fontsize=10)
ax1.set_title("Coût par deal signé (CPA)\npar canal d'acquisition", fontsize=11,
              fontweight="bold", color=NAVY, loc="left")
ax1.tick_params(axis="x", labelsize=9)
ax1.set_ylim(0, max(roi_df["cpa"]) * 1.35)

note = ax1.text(0.5, -0.16, "* Canaux gratuits (CPL=0€). CPA outbound = coût temps humain estimé.",
                transform=ax1.transAxes, ha="center", fontsize=7.5, color=MUTED, style="italic")

# ── Marge vs Coût scatter + ROI ──
non_zero = roi_df[roi_df["cost"] > 0].copy()
zero = roi_df[roi_df["cost"] == 0].copy()

ax2.scatter(non_zero["cost"], non_zero["marge"],
            s=[r["n"]*8 for _, r in non_zero.iterrows()],
            c=[OUTBND]*len(non_zero), alpha=0.8, edgecolors=NAVY, linewidths=1.5, zorder=5)

for _, row in non_zero.iterrows():
    ax2.annotate(row["source"], (row["cost"], row["marge"]),
                textcoords="offset points", xytext=(5,5),
                fontsize=8, color=NAVY)

max_cost = non_zero["cost"].max() * 1.1
ax2.plot([0, max_cost], [0, max_cost], color="#E5E7EB", linewidth=1.5, linestyle="--", label="Seuil ROI = 1×")
ax2.plot([0, max_cost], [0, max_cost*3], color=GOLD, linewidth=1, linestyle=":", alpha=0.7, label="ROI = 3×")
ax2.plot([0, max_cost], [0, max_cost*6], color=GREEN, linewidth=1, linestyle=":", alpha=0.7, label="ROI = 6×")

ax2.set_xlabel("Coût total acquisition (€)", fontsize=10)
ax2.set_ylabel("Marge brute générée (€)", fontsize=10)
ax2.set_title("Marge brute vs Coût d'acquisition\n(canaux payants uniquement)", fontsize=11,
              fontweight="bold", color=NAVY, loc="left")
ax2.legend(fontsize=8, frameon=True, loc="upper left")

fig.suptitle("VROOM MARKET · Analyse ROI par Canal · 2025",
             fontsize=13, fontweight="bold", color=NAVY, y=1.01)

plt.tight_layout()
plt.savefig("/home/claude/vroom-portfolio/dashboards/fig4_roi_cpa_analyse.png",
            dpi=160, bbox_inches="tight", facecolor=BG)
plt.close()
print("Fig 4 saved")
print("\nAll charts saved to dashboards/")
