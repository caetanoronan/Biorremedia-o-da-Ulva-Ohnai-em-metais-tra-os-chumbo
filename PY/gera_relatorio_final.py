import pandas as pd
from datetime import datetime

comparacao = pd.read_csv("Comparacao_COM_vs_SEM_NaNO3.csv")
posthoc_resumo = pd.read_csv("PostHoc_Resumo.csv")
posthoc_comp = pd.read_csv("PostHoc_Comparacoes.csv")

# Helpers HTML

def df_to_html_table(df):
    header = "".join(f"<th>{c}</th>" for c in df.columns)
    rows = []
    for _, r in df.iterrows():
        cells = "".join(f"<td>{r[c]}</td>" for c in df.columns)
        rows.append(f"<tr>{cells}</tr>")
    body = "\n".join(rows)
    return f"<table><thead><tr>{header}</tr></thead><tbody>{body}</tbody></table>"

# Tabelas para o relatorio
comp_fmt = comparacao.copy()
for c in ["Sem_NaNO3", "Com_NaNO3", "Efeito"]:
    comp_fmt[c] = comp_fmt[c].map(lambda x: f"{x:.4f}")
comp_fmt["Efeito_Pct"] = comp_fmt["Efeito_Pct"].map(lambda x: f"{x:.2f}%")

posthoc_resumo_fmt = posthoc_resumo.copy()
posthoc_resumo_fmt.columns = ["Contexto", "Variável", "Comparações significativas (U)"]

posthoc_signif = posthoc_comp[posthoc_comp["Signif_u(0.05)"] == "Sim"].copy()
posthoc_signif = posthoc_signif[["Contexto", "Variavel", "Grupo_A", "Grupo_B", "u_p_bonf"]]
posthoc_signif.columns = ["Contexto", "Variável", "Grupo A", "Grupo B", "p-ajustado (U)"]
posthoc_signif["p-ajustado (U)"] = posthoc_signif["p-ajustado (U)"].map(lambda x: f"{x:.4f}")

comp_html = df_to_html_table(comp_fmt)
posthoc_resumo_html = df_to_html_table(posthoc_resumo_fmt)
if len(posthoc_signif):
  posthoc_signif_html = df_to_html_table(posthoc_signif)
else:
  posthoc_signif_html = df_to_html_table(pd.DataFrame([{
    "Contexto": "-",
    "Variável": "-",
    "Grupo A": "-",
    "Grupo B": "-",
    "p-ajustado (U)": "-"
  }]))

hoje = datetime.now().strftime("%d/%m/%Y")

html = f"""<!DOCTYPE html>
<html lang='pt-BR'>
<head>
  <meta charset='UTF-8'>
  <meta name='viewport' content='width=device-width, initial-scale=1.0'>
  <title>Relatório Final - Experimento de Biorremediação</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      color: #16352d;
      background: #e5f5f9;
      line-height: 1.65;
    }}
    .wrap {{ width: min(98vw, 1920px); margin: 24px auto; padding: 0 10px 24px; }}
    .card {{
      background: #fff;
      border-radius: 12px;
      box-shadow: 0 8px 24px rgba(23, 74, 57, 0.12);
      padding: 26px;
      margin-bottom: 18px;
      border: 1px solid #d9efe8;
    }}
    h1, h2, h3 {{ color: #1f6a4d; margin-top: 0; }}
    h1 {{ font-size: 2rem; margin-bottom: 8px; }}
    h2 {{ font-size: 1.35rem; border-bottom: 2px solid #d9efe8; padding-bottom: 8px; }}
    .meta {{ color: #3f5d55; }}
    .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 12px; }}
    .kpi {{
      border-left: 5px solid #2ca25f;
      background: #f2fbf8;
      border-radius: 8px;
      padding: 12px;
    }}
    .kpi strong {{ display: block; font-size: 1.15rem; color: #1f6a4d; }}
    ul {{ margin: 8px 0 0 18px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 0.92rem; }}
    th, td {{ border-bottom: 1px solid #bcded4; padding: 9px; text-align: left; }}
    th {{ background: #e5f5f9; color: #1e5f47; }}
    .small {{ color: #3f5d55; font-size: 0.9rem; }}
    .ok {{ color: #1f6a4d; font-weight: 700; }}
    .warn {{ color: #2d7b59; font-weight: 700; }}
    a {{ color: #238b45; text-decoration: underline; text-underline-offset: 2px; }}
    tbody tr:nth-child(even) {{ background: #fbfefc; }}
    tbody tr:hover {{ background: #f3fbf8; }}

    .tabs {{
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      margin-bottom: 14px;
    }}
    .tab-btn {{
      background: #f2fbf8;
      color: #1f6a4d;
      border: 1px solid #bcded4;
      border-radius: 8px;
      padding: 8px 12px;
      cursor: pointer;
      font-weight: 600;
    }}
    .tab-btn.active {{
      background: #2ca25f;
      color: #fff;
      border-color: #238b45;
      box-shadow: inset 0 -3px 0 #99d8c9;
    }}
    .tab-pane {{ display: none; }}
    .tab-pane.active {{ display: block; }}
    .table-expl {{
      background: #f2fbf8;
      border-left: 4px solid #2ca25f;
      padding: 10px 12px;
      border-radius: 6px;
      margin-bottom: 10px;
      color: #2a554a;
      font-size: 0.93rem;
      line-height: 1.6;
    }}
    .table-scroll {{
      overflow-x: auto;
      border: 1px solid #d9efe8;
      border-radius: 8px;
      width: 100%;
    }}

    .theme-toggle {{
      position: fixed;
      top: 14px;
      right: 14px;
      z-index: 9999;
      background: #fff;
      color: #2ca25f;
      border: 1px solid #bcded4;
      border-radius: 999px;
      padding: 8px 12px;
      font-weight: 700;
      cursor: pointer;
      box-shadow: 0 8px 20px rgba(23, 74, 57, 0.12);
    }}

    @media (max-width: 900px) {{
      .tab-btn {{ width: 100%; text-align: left; }}
    }}

    html[data-theme='dark'] body {{
      color: #e8f5ef;
      background: #0f1f1a;
    }}

    html[data-theme='dark'] .card {{
      background: #132923;
      border-color: #2b5143;
      box-shadow: 0 10px 24px rgba(0, 0, 0, 0.35);
    }}

    html[data-theme='dark'] h1,
    html[data-theme='dark'] h2,
    html[data-theme='dark'] h3,
    html[data-theme='dark'] .kpi strong,
    html[data-theme='dark'] .ok,
    html[data-theme='dark'] .warn {{
      color: #ccece6;
    }}

    html[data-theme='dark'] .meta,
    html[data-theme='dark'] .small,
    html[data-theme='dark'] .table-expl {{
      color: #c0ddd1;
    }}

    html[data-theme='dark'] .kpi,
    html[data-theme='dark'] .table-expl,
    html[data-theme='dark'] .tab-btn,
    html[data-theme='dark'] .table-scroll {{
      background: #173228;
      border-color: #2b5143;
    }}

    html[data-theme='dark'] .tab-btn {{
      color: #d8efe6;
    }}

    html[data-theme='dark'] .tab-btn.active {{
      background: #2ca25f;
      border-color: #66c2a4;
      box-shadow: inset 0 -3px 0 #99d8c9;
    }}

    html[data-theme='dark'] th,
    html[data-theme='dark'] td {{
      border-bottom-color: #2b5143;
    }}

    html[data-theme='dark'] th {{
      background: #1a3a30;
      color: #d8efe6;
    }}

    html[data-theme='dark'] tbody tr:nth-child(even) {{
      background: #173228;
    }}

    html[data-theme='dark'] tbody tr:hover {{
      background: #1d3d32;
    }}

    html[data-theme='dark'] a {{
      color: #99d8c9;
    }}
  </style>
</head>
<body>
  <button class='theme-toggle' id='themeToggle' type='button'>🌙 Escuro</button>
  <div class='wrap'>
    <section class='card'>
      <h1>Relatório Final do Experimento de Biorremediação</h1>
      <p class='meta'>Data de geração: {hoje}</p>
      <p>Este relatório consolida as análises de Biomassa e PAM (fluorescência), a comparação entre tratamentos (SEM e COM NaNO3), os testes pós-hoc e os painéis interativos para apresentação.</p>
    </section>

    <section class='card'>
      <h2>1. Objetivo e Escopo</h2>
      <p>Avaliar o efeito do tratamento com NaNO3 sobre crescimento (biomassa) e eficiência fotossintética (PAM/Fv/Fm), comparando grupos C0, C1, C2 e C3.</p>
      <ul>
        <li>Fluxo completo de limpeza e padronização dos dados.</li>
        <li>Estatística descritiva e testes globais (ANOVA e Kruskal-Wallis).</li>
        <li>Comparação COM vs SEM NaNO3.</li>
        <li>Testes pós-hoc pareados com correção de Bonferroni.</li>
        <li>Dashboards com modo didático e glossário.</li>
      </ul>
    </section>

    <section class='card'>
      <h2>2. Resumo Executivo</h2>
      <div class='kpi-grid'>
        <div class='kpi'><strong>Biomassa COM NaNO3 (inicial)</strong>Diferença global significativa entre grupos.</div>
        <div class='kpi'><strong>PAM SEM NaNO3 (final)</strong>Diferença global significativa entre grupos.</div>
        <div class='kpi'><strong>PAM COM NaNO3 (final)</strong>Sem diferença global significativa.</div>
        <div class='kpi'><strong>Comparação COM vs SEM</strong>Ganhos expressivos em C1 para Biomassa e PAM.</div>
      </div>
      <p class='small'>Nota: valor extremo de biomassa final em C3 (1713.0) foi tratado analiticamente com visualizações complementares COM e SEM outlier.</p>
    </section>

    <section class='card'>
      <h2>2.1 Interpretação de Correlações</h2>
      <div class='table-expl'>
        Este guia ajuda a interpretar relações entre variáveis (ex.: Biomassa e PAM) de forma correta e comunicável.
      </div>
      <ul>
        <li><strong>O que é correlação:</strong> mede como duas variáveis variam juntas (direção e intensidade).</li>
        <li><strong>Leitura do coeficiente:</strong> próximo de +1 (positiva forte), próximo de -1 (negativa forte), próximo de 0 (fraca).</li>
        <li><strong>Cautela:</strong> correlação não prova causalidade.</li>
        <li><strong>Confiabilidade:</strong> interpretar sempre junto com p-valor e tamanho amostral (n).</li>
        <li><strong>Contexto biológico:</strong> traduzir o achado em linguagem prática, indicando se maior biomassa acompanha melhor desempenho fotossintético em cada grupo.</li>
      </ul>
    </section>

    <section class='card'>
      <h2>3. Resultados em Abas</h2>
      <div class='tabs' role='tablist' aria-label='Resultados estatísticos'>
        <button class='tab-btn active' role='tab' aria-selected='true' data-tab='tab-comparacao'>Comparação COM vs SEM</button>
        <button class='tab-btn' role='tab' aria-selected='false' data-tab='tab-posthoc-resumo'>Pós-hoc Resumo</button>
        <button class='tab-btn' role='tab' aria-selected='false' data-tab='tab-posthoc-detalhe'>Pós-hoc Detalhado</button>
      </div>

      <div id='tab-comparacao' class='tab-pane active' role='tabpanel'>
        <div class='table-expl'>
          Esta tabela mostra a diferença entre SEM e COM NaNO3 por grupo. Se o Efeito % for positivo, houve ganho com NaNO3; se for negativo, houve redução.
        </div>
        <div class='table-scroll'>{comp_html}</div>
        <p class='small'>Leitura: efeito positivo indica melhora com NaNO3; efeito negativo indica redução.</p>
      </div>

      <div id='tab-posthoc-resumo' class='tab-pane' role='tabpanel'>
        <div class='table-expl'>
          Esta tabela resume quantas comparações entre pares de grupos foram significativas em cada contexto após correção de Bonferroni.
        </div>
        <div class='table-scroll'>{posthoc_resumo_html}</div>
      </div>

      <div id='tab-posthoc-detalhe' class='tab-pane' role='tabpanel'>
        <div class='table-expl'>
          Esta tabela lista os pares de grupos com diferença significativa (p-ajustado &lt; 0.05) no teste U de Mann-Whitney com ajuste de Bonferroni.
        </div>
        <div class='table-scroll'>{posthoc_signif_html}</div>
        <p class='small'>Critério de significância: p-ajustado &lt; 0.05.</p>
      </div>
    </section>

    <section class='card'>
      <h2>4. Conclusões Integradas</h2>
      <ul>
        <li><span class='ok'>Biomassa:</span> resposta positiva ao NaNO3 em grupos específicos, com destaque para C1.</li>
        <li><span class='ok'>PAM:</span> forte ganho em C0, C1 e C2 no cenário COM vs SEM, principalmente no PAM final.</li>
        <li><span class='warn'>Qualidade de dado:</span> presença de outlier extremo em C3 exige cautela interpretativa e foi destacada em todos os painéis.</li>
        <li>Resultados pós-hoc confirmam diferenças pareadas em contextos-chave identificados pelos testes globais.</li>
      </ul>
    </section>

    <section class='card'>
      <h2>5. Materiais de Entrega</h2>
      <ul>
        <li><a href='Dashboard_Biomassa_COM_NaNO3.html' target='_blank'>Dashboard Biomassa COM NaNO3</a></li>
        <li><a href='Dashboard_PAM.html' target='_blank'>Dashboard PAM (Fluorescência)</a></li>
        <li><a href='Dashboard_Consolidado_Biorremediacao.html' target='_blank'>Dashboard Consolidado</a></li>
        <li><a href='Relatorio_PostHoc.html' target='_blank'>Relatório Pós-Hoc (HTML)</a></li>
        <li>Planilhas: Analise_Biomassa_COM_NaNO3.xlsx, Analise_PAM.xlsx, Comparacao_COM_vs_SEM_NaNO3.xlsx, PostHoc_Comparacoes.xlsx</li>
      </ul>
    </section>
  </div>

  <script>
    (function() {{
      const root = document.documentElement;
      const themeBtn = document.getElementById('themeToggle');
      const savedTheme = localStorage.getItem('theme-preference') || 'light';
      root.setAttribute('data-theme', savedTheme);
      themeBtn.textContent = savedTheme === 'dark' ? '☀️ Claro' : '🌙 Escuro';

      themeBtn.addEventListener('click', function() {{
        const current = root.getAttribute('data-theme') || 'light';
        const next = current === 'dark' ? 'light' : 'dark';
        root.setAttribute('data-theme', next);
        localStorage.setItem('theme-preference', next);
        themeBtn.textContent = next === 'dark' ? '☀️ Claro' : '🌙 Escuro';
      }});

      const buttons = document.querySelectorAll('.tab-btn');
      const panes = document.querySelectorAll('.tab-pane');

      buttons.forEach((btn) => {{
        btn.addEventListener('click', function() {{
          const target = this.getAttribute('data-tab');

          buttons.forEach((b) => {{
            b.classList.remove('active');
            b.setAttribute('aria-selected', 'false');
          }});

          panes.forEach((p) => p.classList.remove('active'));

          this.classList.add('active');
          this.setAttribute('aria-selected', 'true');

          const pane = document.getElementById(target);
          if (pane) pane.classList.add('active');
        }});
      }});
    }})();
  </script>
</body>
</html>
"""

with open("Relatorio_Final_Biorremediacao.html", "w", encoding="utf-8") as f:
    f.write(html)

print("OK - Relatorio_Final_Biorremediacao.html")
