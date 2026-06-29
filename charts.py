"""
Genera los dos graficos del reporte como PNG (para incrustar en el mail via CID).
  charts/comparativo.png -> barras 2026 vs 2025 por sucursal (acumulado del mes)
  charts/progreso.png    -> barra horizontal total mes en curso vs anio anterior

Imagenes estaticas de alta resolucion (nitidas en pantallas retina).
"""
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

# Paleta corporativa: azul marino (en curso) + celeste (anio anterior)
AZUL = "#1f3a6e"        # azul marino - 2026
CELESTE = "#7fb3e0"     # celeste - 2025
GRIS_TXT = "#8a8a8a"
VERDE = "#1a7d2e"
ROJO = "#c62828"

DPI = 200  # alta resolucion para que no se vea pixelado


def _money_k(x, _):
    if x >= 1000:
        return f"${x/1000:.0f}k"
    return f"${x:.0f}"


def chart_comparativo(rows, a26_lbl, a25_lbl, out_path):
    branches = [r["branch"] for r in rows]
    v26 = [r["a26"] for r in rows]
    v25 = [r["a25"] for r in rows]

    n = len(branches)
    x = range(n)
    w = 0.38

    fig, ax = plt.subplots(figsize=(8.4, 3.8), dpi=DPI)
    ax.bar([i - w/2 for i in x], v26, width=w, label=a26_lbl, color=AZUL, zorder=3)
    ax.bar([i + w/2 for i in x], v25, width=w, label=a25_lbl, color=CELESTE, zorder=3)

    ax.set_xticks(list(x))
    ax.set_xticklabels(branches, fontsize=9, color="#333333")
    ax.yaxis.set_major_formatter(FuncFormatter(_money_k))
    ax.tick_params(axis="y", labelsize=8, colors=GRIS_TXT)
    ax.tick_params(axis="x", length=0)

    for spine in ["top", "right", "left"]:
        ax.spines[spine].set_visible(False)
    ax.spines["bottom"].set_color("#dddddd")
    ax.grid(axis="y", color="#eeeeee", zorder=0)

    ax.legend(loc="upper right", frameon=False, fontsize=9)
    fig.tight_layout(pad=0.6)
    fig.savefig(out_path, transparent=True, bbox_inches="tight")
    plt.close(fig)


def chart_progreso(a26, a25, pct, out_path):
    fig, ax = plt.subplots(figsize=(8.4, 1.25), dpi=DPI)

    base = max(a26, a25)
    ax.barh([0], [a25], height=0.5, color="#e9e9e9", zorder=2)
    color = VERDE if a26 >= a25 else ROJO
    ax.barh([0], [a26], height=0.5, color=color, zorder=3)
    ax.axvline(a25, color=GRIS_TXT, linestyle=":", linewidth=1.2, zorder=4)

    ax.set_xlim(0, base * 1.12)
    ax.set_ylim(-0.6, 0.6)
    ax.axis("off")

    ax.text(a26, 0, f"  ${a26:,.0f}", va="center", ha="left", fontsize=12,
            fontweight="bold", color="#111111", zorder=5)
    ax.text(a25, 0.42, f"Año ant. ${a25:,.0f}", va="bottom", ha="center",
            fontsize=8, color=GRIS_TXT)

    fig.tight_layout(pad=0.3)
    fig.savefig(out_path, transparent=True, bbox_inches="tight")
    plt.close(fig)


def build_charts(rows, totals, mes, a26_lbl, a25_lbl, out_dir="charts"):
    d = Path(out_dir)
    d.mkdir(exist_ok=True)
    p1 = d / "comparativo.png"
    p2 = d / "progreso.png"
    chart_comparativo(rows, a26_lbl, a25_lbl, p1)
    chart_progreso(totals["a26"], totals["a25"], totals["pct"], p2)
    return str(p1), str(p2)
