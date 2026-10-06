# @file     geradorGraficos.py
# @author   Frederico S. Gonçalves

# Importacao das bibliotecas.
import sys
import matplotlib.pyplot as plt
import pandas as pd

# Constantes
GRAFICO_EXPERIMENTAL = "0"
GRAFICO_AJUSTADO = "1"
GRAFICO_GERAL = "2"

# Variáveis.
nome_arquivo = sys.argv[1]
tipo_grafico = sys.argv[2]

# Dataframe.
df = pd.read_excel(nome_arquivo, decimal=',', sheet_name="CO2_in_out")

# Lê o resumo.
linhas_texto = [str(x) for x in df["Minuto"].dropna() if isinstance(x, str)]
resumo_extraido = {}
for linha in linhas_texto:
    if " = " in linha:
        chave, valor = linha.split(" = ", 1)
        chave = chave.replace("-", "").strip()
        resumo_extraido[chave] = valor.strip()

# Converte todas as linhas que serão usadas nos gráficos para números.
df["Minuto_out"] = pd.to_numeric(df["Minuto_out"], errors='coerce')
df["CO2_out_%"] = pd.to_numeric(df["CO2_out_%"], errors='coerce')
df["CO2_out_suav_%"] = pd.to_numeric(df["CO2_out_suav_%"], errors='coerce')

# Plotagem dos gráficos.
if df is not None and tipo_grafico == GRAFICO_EXPERIMENTAL:   # Curvas do Gráfico Experimental.
  plt.figure("Gráfico Experimental", figsize=(12, 6))
  plt.plot(df["Minuto_out"], df["CO2_out_%"], color="red", linestyle="--", marker="o", markersize=4, label="CO2 de Saída")

elif df is not None and tipo_grafico == GRAFICO_AJUSTADO:   # Curvas do Gráfico Ajustado.
  plt.figure("Gráfico Ajustado", figsize=(12, 6))
  plt.plot(df["Minuto_out"], df["CO2_out_suav_%"], color="red", linewidth=2, label="CO2 de Saída (Suavizado)")

elif df is not None and tipo_grafico == GRAFICO_GERAL:   # Curvas do Gráfico Geral.
  plt.figure("Gráfico Geral", figsize=(12, 6))
  plt.plot(df["Minuto_out"], df["CO2_out_%"], color="red", linestyle="--", marker="o", markersize=3, linewidth=0.5, label="CO2 de Saída")
  plt.plot(df["Minuto_out"], df["CO2_out_suav_%"], color="red", linewidth=2, label="CO2 de Saída (Suavizado)")

# Informações comuns a todos os gráficos.
plt.title("Teor de CO2 - Saída")
plt.xlabel("Tempo (m)")
plt.ylabel("CO2 (%)")
plt.grid(True)
plt.legend()
plt.subplots_adjust(bottom=0.20)
texto_info = (
    f"\nminutos OUT={resumo_extraido.get('Minutos_out', '0')}  "
    f"(descartados no aquecimento: {resumo_extraido.get('Descartados', '0')})\n"
    f"media CO2_out={resumo_extraido.get('Media_out', '0')}%\n"
)
plt.figtext(0.5, 0.05, texto_info, ha="center", fontsize=10, color="black")
plt.show()