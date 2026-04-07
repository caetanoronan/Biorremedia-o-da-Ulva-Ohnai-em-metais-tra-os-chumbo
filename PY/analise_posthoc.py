import itertools
import re

import pandas as pd
from scipy import stats

ODS_FILE = "DADOS DO EXPERIMENTO BIORREMEDIAÇÃO - DORLAM'S.ods"


def extract_group(value):
    match = re.search(r"(C[0-3])", str(value))
    return match.group(1) if match else None


def bonferroni_adjust(p_values):
    m = len(p_values)
    return [min(p * m, 1.0) for p in p_values]


def pairwise_posthoc(df, group_col, value_col, context_label):
    groups = sorted(df[group_col].dropna().unique())
    pairs = list(itertools.combinations(groups, 2))

    raw_t = []
    raw_u = []
    rows = []

    for g1, g2 in pairs:
        x = df[df[group_col] == g1][value_col].dropna().values
        y = df[df[group_col] == g2][value_col].dropna().values

        t_stat, t_p = stats.ttest_ind(x, y, equal_var=False)
        u_stat, u_p = stats.mannwhitneyu(x, y, alternative="two-sided")

        raw_t.append(t_p)
        raw_u.append(u_p)
        rows.append({
            "Contexto": context_label,
            "Variavel": value_col,
            "Grupo_A": g1,
            "Grupo_B": g2,
            "n_A": len(x),
            "n_B": len(y),
            "Media_A": x.mean(),
            "Media_B": y.mean(),
            "Mediana_A": pd.Series(x).median(),
            "Mediana_B": pd.Series(y).median(),
            "t_stat": t_stat,
            "t_p_raw": t_p,
            "u_stat": u_stat,
            "u_p_raw": u_p,
        })

    t_adj = bonferroni_adjust(raw_t)
    u_adj = bonferroni_adjust(raw_u)

    for i, row in enumerate(rows):
        row["t_p_bonf"] = t_adj[i]
        row["u_p_bonf"] = u_adj[i]
        row["Signif_t(0.05)"] = "Sim" if t_adj[i] < 0.05 else "Nao"
        row["Signif_u(0.05)"] = "Sim" if u_adj[i] < 0.05 else "Nao"

    return pd.DataFrame(rows)


# === Biomassa COM NaNO3 ===
bi_ini_com = pd.read_excel(ODS_FILE, sheet_name="Biomassa inicial  com NaNO3", engine="odf")
bi_ini_com.columns = bi_ini_com.columns.str.strip()

biomassa_com = pd.DataFrame()
biomassa_com["grupo"] = bi_ini_com.iloc[:, 0].apply(extract_group)
biomassa_com["inicial"] = pd.to_numeric(bi_ini_com.iloc[:, 1], errors="coerce")
biomassa_com = biomassa_com.dropna()

# === PAM SEM e COM NaNO3 ===
pam_ini_sem = pd.read_excel(ODS_FILE, sheet_name="PAM INICIAL-sem NaNO3", engine="odf")
pam_fim_sem = pd.read_excel(ODS_FILE, sheet_name="PAM sem NaNO3", engine="odf")
pam_ini_com = pd.read_excel(ODS_FILE, sheet_name="PAM 1- NaNO3", engine="odf")
pam_fim_com = pd.read_excel(ODS_FILE, sheet_name="PAM 2- NaNO3", engine="odf")

for df in [pam_ini_sem, pam_fim_sem, pam_ini_com, pam_fim_com]:
    df.columns = df.columns.str.strip()

pam_sem = pd.DataFrame()
pam_sem["grupo"] = pam_ini_sem["FRASCOS"].apply(extract_group)
pam_sem["inicial"] = pam_ini_sem.iloc[:, 1:4].mean(axis=1)
pam_sem["final"] = pam_fim_sem.iloc[:, 1:4].mean(axis=1)
pam_sem["mudanca"] = pam_sem["final"] - pam_sem["inicial"]
pam_sem = pam_sem.dropna()

pam_com = pd.DataFrame()
pam_com["grupo"] = pam_ini_com["FRASCOS"].apply(extract_group)
pam_com["inicial"] = pam_ini_com.iloc[:, 1:4].mean(axis=1)
pam_com["final"] = pam_fim_com.iloc[:, 1:4].mean(axis=1)
pam_com["mudanca"] = pam_com["final"] - pam_com["inicial"]
pam_com = pam_com.dropna()

# Apenas variaveis com efeito global significativo nas analises anteriores
frames = [
    pairwise_posthoc(biomassa_com, "grupo", "inicial", "Biomassa COM NaNO3"),
    pairwise_posthoc(pam_sem, "grupo", "final", "PAM SEM NaNO3"),
    pairwise_posthoc(pam_sem, "grupo", "mudanca", "PAM SEM NaNO3"),
    pairwise_posthoc(pam_com, "grupo", "inicial", "PAM COM NaNO3"),
    pairwise_posthoc(pam_com, "grupo", "mudanca", "PAM COM NaNO3"),
]

result = pd.concat(frames, ignore_index=True)
result = result.round(6)

# Resumo objetivo
summary = (
    result[result["Signif_u(0.05)"] == "Sim"]
    .groupby(["Contexto", "Variavel"]) 
    .size()
    .reset_index(name="Comparacoes_significativas_U")
)

result.to_csv("PostHoc_Comparacoes.csv", index=False)
summary.to_csv("PostHoc_Resumo.csv", index=False)

with pd.ExcelWriter("PostHoc_Comparacoes.xlsx", engine="openpyxl") as writer:
    result.to_excel(writer, sheet_name="Comparacoes", index=False)
    summary.to_excel(writer, sheet_name="Resumo", index=False)

# Relatorio HTML simples
rows_html = "\n".join(
    f"<tr><td>{r.Contexto}</td><td>{r.Variavel}</td><td>{r.Grupo_A} vs {r.Grupo_B}</td><td>{r.u_p_bonf:.4f}</td><td>{r['Signif_u(0.05)']}</td></tr>"
    for _, r in result.iterrows()
)

html = f"""<!DOCTYPE html>
<html lang='pt-BR'>
<head>
  <meta charset='UTF-8'>
  <meta name='viewport' content='width=device-width, initial-scale=1.0'>
  <title>Relatorio Post-Hoc</title>
  <style>
    body {{ font-family: Segoe UI, Arial, sans-serif; margin: 24px; color: #1f2a44; }}
    h1 {{ color: #0d3b66; }}
    table {{ border-collapse: collapse; width: 100%; margin-top: 12px; }}
    th, td {{ border-bottom: 1px solid #dde4ee; padding: 8px; text-align: left; font-size: 14px; }}
    th {{ background: #f5f8fc; }}
  </style>
</head>
<body>
  <h1>Relatorio Final - Testes Post-Hoc</h1>
  <p>Metodo: comparacoes pareadas com correcao de Bonferroni (teste U de Mann-Whitney e t de Welch).</p>
  <p>Arquivos exportados: PostHoc_Comparacoes.csv, PostHoc_Resumo.csv, PostHoc_Comparacoes.xlsx.</p>
  <table>
    <thead>
      <tr>
        <th>Contexto</th><th>Variavel</th><th>Comparacao</th><th>p-ajustado (U)</th><th>Significativo</th>
      </tr>
    </thead>
    <tbody>
      {rows_html}
    </tbody>
  </table>
</body>
</html>
"""

with open("Relatorio_PostHoc.html", "w", encoding="utf-8") as f:
    f.write(html)

print("OK - PostHoc_Comparacoes.csv")
print("OK - PostHoc_Resumo.csv")
print("OK - PostHoc_Comparacoes.xlsx")
print("OK - Relatorio_PostHoc.html")
