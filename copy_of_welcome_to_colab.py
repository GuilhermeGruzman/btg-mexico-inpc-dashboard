import requests
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime

# =========================
# 1. Baixar dados do Focus
# =========================

url = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoAnuais?$format=json"
r = requests.get(url)
data = r.json()["value"]

df = pd.DataFrame(data)

# =========================
# 2. Ajustar tipos
# =========================

df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
df["DataReferencia"] = pd.to_numeric(df["DataReferencia"], errors="coerce")
df["Mediana"] = pd.to_numeric(df["Mediana"], errors="coerce")
df["baseCalculo"] = pd.to_numeric(df["baseCalculo"], errors="coerce")

# =========================
# 3. Filtrar IPCA, base=0, desde 2024
# =========================

df = df[
    (df["Indicador"] == "IPCA") &
    (df["baseCalculo"] == 0) &
    (df["Data"] >= "2024-01-01")
].copy()

# Remover repetição artificial
df = df[df["Mediana"].ne(df["Mediana"].shift())].copy()

# =========================
# 4. Separar por ano (com corte)
# =========================

df_2024 = df[
    (df["DataReferencia"] == 2024) &
    (df["Data"] <= "2024-12-31")
].copy()

df_2025 = df[
    (df["DataReferencia"] == 2025) &
    (df["Data"] <= "2025-12-31")
].copy()

df_2026 = df[
    (df["DataReferencia"] == 2026)
].copy()

df_2024 = df_2024.sort_values("Data").set_index("Data")
df_2025 = df_2025.sort_values("Data").set_index("Data")
df_2026 = df_2026.sort_values("Data").set_index("Data")

# =========================
# 5. Valores finais
# =========================

ipca_2024 = 4.83
ipca_2025 = 4.26

last_2026_date = df_2026.index.max()
last_2026_value = df_2026.loc[last_2026_date, "Mediana"]

date_2024 = pd.Timestamp("2024-12-31")
date_2025 = pd.Timestamp("2025-12-31")
date_2026 = last_2026_date

# =========================
# 6. Data para salvar
# =========================

hoje = datetime.today().strftime("%Y%m%d")

arquivo = f"/content/focus_ipca_2024_2026_{hoje}.png"

# =========================
# 7. Plot
# =========================

plt.figure(figsize=(12,6))

line_2024, = plt.plot(
    df_2024.index, df_2024["Mediana"],
    label="IPCA 2024 (Real: 4,83%)",
    linewidth=2.5
)

line_2025, = plt.plot(
    df_2025.index, df_2025["Mediana"],
    label="IPCA 2025 (Real: 4,26%)",
    linewidth=2.5
)

line_2026, = plt.plot(
    df_2026.index, df_2026["Mediana"],
    label="IPCA 2026 (Última exp.)",
    linewidth=2.5
)

# =========================
# 8. Balões
# =========================

plt.text(
    date_2024, ipca_2024,
    "4,83%",
    color=line_2024.get_color(),
    fontsize=9,
    ha="left",
    va="bottom",
    bbox=dict(boxstyle="round,pad=0.2", fc="white", ec=line_2024.get_color(), alpha=0.9)
)

plt.text(
    date_2025, ipca_2025,
    "4,26%",
    color=line_2025.get_color(),
    fontsize=9,
    ha="left",
    va="top",
    bbox=dict(boxstyle="round,pad=0.2", fc="white", ec=line_2025.get_color(), alpha=0.9)
)

plt.text(
    date_2026, last_2026_value,
    f"{last_2026_value:.2f}%",
    color=line_2026.get_color(),
    fontsize=9,
    ha="left",
    va="bottom",
    bbox=dict(boxstyle="round,pad=0.2", fc="white", ec=line_2026.get_color(), alpha=0.9)
)

# =========================
# 9. Layout
# =========================

plt.xlabel("Data")
plt.ylabel("Inflação (%)")
plt.title("Expectativas Focus IPCA (2024–2026)")
plt.legend()
plt.grid(alpha=0.3)

# Eixo X: 2 meses
ax = plt.gca()
ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
ax.xaxis.set_major_formatter(mdates.DateFormatter("%b/%Y"))

plt.xticks(rotation=45)

# =========================
# 10. Rodapé
# =========================

plt.figtext(
    0.01, -0.06,
    f"Última atualização: {datetime.today().strftime('%d/%m/%Y')} | Elaborado por Guilherme Gruzman",
    ha="left",
    fontsize=9,
    alpha=0.8
)

plt.tight_layout()

# =========================
# 11. Salvar
# =========================

plt.savefig(arquivo, dpi=300, bbox_inches="tight")

plt.show()

print("Arquivo salvo em:", arquivo)
