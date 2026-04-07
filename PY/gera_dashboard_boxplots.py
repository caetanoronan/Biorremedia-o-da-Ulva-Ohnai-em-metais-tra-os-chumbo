import pandas as pd
import json

# Carregar dados
df_boxplot = pd.read_csv("CSV/Dados_BoxPlot.csv")
df_estatisticas = pd.read_csv("CSV/Estatisticas_Descritivas.csv")

# Preparar dados para JavaScript
dados_biomassa = df_boxplot[df_boxplot['Variavel'] == 'Biomassa Final'].to_dict('records')
dados_pam = df_boxplot[df_boxplot['Variavel'] == 'PAM Final (Fv/Fm)'].to_dict('records')

est_biomassa = df_estatisticas[df_estatisticas['Variavel'] == 'Biomassa Final'].to_dict('records')
est_pam = df_estatisticas[df_estatisticas['Variavel'] == 'PAM Final (Fv/Fm)'].to_dict('records')

# Converter para JSON
dados_biomassa_json = json.dumps(dados_biomassa)
dados_pam_json = json.dumps(dados_pam)
est_biomassa_json = json.dumps(est_biomassa)
est_pam_json = json.dumps(est_pam)

html_content = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Análise Descritiva com Box Plots - Biorremediação</title>
  <script src="libs/plotly-2.35.2.min.js"></script>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;700&family=IBM+Plex+Sans:wght@400;500;600&display=swap" rel="stylesheet" />
  <style>
    :root {{
      --bg: #f5f7fa;
      --ink: #1d2a24;
      --muted: #4b5f57;
      --card: #ffffff;
      --line: #d0d7ce;
      --accent: #17684f;
      --accent-2: #c73f13;
      --shadow: 0 4px 12px rgba(56, 44, 21, 0.08);
      --r: 8px;
    }}

    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "IBM Plex Sans", sans-serif;
      color: var(--ink);
      background: var(--bg);
      padding: 20px;
    }}

    .wrap {{
      max-width: 1400px;
      margin: 0 auto;
    }}

    h1 {{
      font-family: "Space Grotesk", sans-serif;
      font-size: 2.2rem;
      margin: 0 0 10px;
      color: #174f3f;
    }}

    .subtitle {{
      color: var(--muted);
      margin-bottom: 30px;
      font-size: 1.05rem;
    }}

    .section-title {{
      font-family: "Space Grotesk", sans-serif;
      font-size: 1.4rem;
      color: var(--accent);
      margin: 40px 0 20px;
      border-bottom: 2px solid var(--line);
      padding-bottom: 10px;
    }}

    .card {{
      background: var(--card);
      border-radius: var(--r);
      box-shadow: var(--shadow);
      padding: 20px;
      margin-bottom: 20px;
      border: 1px solid var(--line);
    }}

    .plot-container {{
      width: 100%;
      height: 500px;
      margin-bottom: 20px;
    }}

    .info-box {{
      background: #e8f4f1;
      border-left: 4px solid var(--accent);
      padding: 15px;
      border-radius: var(--r);
      margin-bottom: 20px;
      font-size: 0.95rem;
      line-height: 1.6;
    }}

    .info-box strong {{
      color: var(--accent);
    }}

    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 0.85rem;
    }}

    th, td {{
      padding: 10px;
      text-align: center;
      border-bottom: 1px solid var(--line);
    }}

    th {{
      background: #f0f5f3;
      font-weight: 600;
      color: var(--accent);
    }}

    td:first-child, th:first-child {{
      text-align: left;
    }}

    tr:hover {{
      background: #f9fdfb;
    }}

    footer {{
      margin-top: 50px;
      padding-top: 20px;
      border-top: 1px solid var(--line);
      color: var(--muted);
      font-size: 0.9rem;
      text-align: center;
    }}

    .back-link {{
      display: inline-block;
      margin-bottom: 20px;
      color: var(--accent);
      text-decoration: none;
      font-weight: 500;
    }}

    .back-link:hover {{
      text-decoration: underline;
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <a href="Dashboard_Biorremediacao.html" class="back-link">← Voltar ao Dashboard Principal</a>

    <h1>📊 Análise Descritiva com Box Plots</h1>
    <p class="subtitle">Estatísticas detalhadas de Biorremediação com NaNO3</p>

    <div class="info-box">
      <strong>ℹ️ Sobre esta análise:</strong> Os box plots apresentam a distribuição completa dos dados, 
      mostrando mediana (linha), quartis Q1 e Q3 (caixa), e mínimo/máximo (whiskers). 
      Combinado com as estatísticas descritivas abaixo, oferece visão completa da variabilidade dos dados.
    </div>

    <!-- BIOMASSA -->
    <div class="section-title">📈 Biomassa Final</div>

    <div class="card">
      <h3 style="margin: 0 0 15px; color: var(--accent);">Distribuição por Grupo e Tratamento</h3>
      <div id="boxplot-biomassa" class="plot-container"></div>
    </div>

    <div class="card">
      <h3 style="margin: 0 0 15px; color: var(--accent);">Estatísticas Descritivas - Biomassa Final</h3>
      <table id="table-biomassa"></table>
    </div>

    <!-- PAM -->
    <div class="section-title">🌿 PAM Final (Fv/Fm)</div>

    <div class="card">
      <h3 style="margin: 0 0 15px; color: var(--accent);">Distribuição por Grupo e Tratamento</h3>
      <div id="boxplot-pam" class="plot-container"></div>
    </div>

    <div class="card">
      <h3 style="margin: 0 0 15px; color: var(--accent);">Estatísticas Descritivas - PAM Final (Fv/Fm)</h3>
      <table id="table-pam"></table>
    </div>

    <footer>
      <p>Gerado automaticamente | Estatísticas calculadas com min, Q1, mediana, Q3, max</p>
      <p><a href="Relatorio_Final_Biorremediacao.html" style="color: var(--accent);">Ver Relatório Completo</a></p>
    </footer>
  </div>

  <script>
    // Dados carregados
    const dataBiomassa = {dados_biomassa_json};
    const dataPAM = {dados_pam_json};
    const estBiomassa = {est_biomassa_json};
    const estPAM = {est_pam_json};

    function gerarBoxPlot(elementId, dados, titulo) {{
      const grupos = [...new Set(dados.map(d => d.Grupo))].sort();
      const tratamentos = [...new Set(dados.map(d => d.Tratamento))];
      const cores = {{'Sem NaNO3': '#2e8b57', 'Com NaNO3': '#dc143c'}};

      const traces = [];
      for (const trat of tratamentos) {{
        for (const grupo of grupos) {{
          const valores = dados
            .filter(d => d.Grupo === grupo && d.Tratamento === trat)
            .map(d => d.Valor)
            .sort((a, b) => a - b);

          if (valores.length > 0) {{
            traces.push({{
              y: valores,
              name: `${{grupo}} - ${{trat}}`,
              type: 'box',
              boxmean: 'sd',
              marker: {{color: cores[trat]}},
              line: {{color: cores[trat]}}
            }});
          }}
        }}
      }}

      const layout = {{
        title: titulo,
        yaxis: {{ title: 'Valor' }},
        xaxis: {{ title: 'Grupo - Tratamento' }},
        showlegend: true,
        plot_bgcolor: '#f9fdfb',
        paper_bgcolor: '#ffffff',
        font: {{ family: 'IBM Plex Sans, sans-serif', color: '#1d2a24', size: 11 }},
        margin: {{ l: 60, r: 40, t: 50, b: 60 }},
        height: 500
      }};

      Plotly.newPlot(elementId, traces, layout, {{ responsive: true }});
    }}

    function gerarTabela(elementId, dados) {{
      const html = `
        <tr>
          <th>Grupo</th>
          <th>Tratamento</th>
          <th>N</th>
          <th>Mean</th>
          <th>Std</th>
          <th>Min</th>
          <th>Q1</th>
          <th>Median</th>
          <th>Q3</th>
          <th>Max</th>
        </tr>
        ${{dados.map(d => `
          <tr>
            <td>${{d.Grupo}}</td>
            <td>${{d.Tratamento}}</td>
            <td>${{d.N}}</td>
            <td>${{d.Mean.toFixed(4)}}</td>
            <td>${{d.Std.toFixed(4)}}</td>
            <td>${{d.Min.toFixed(4)}}</td>
            <td>${{d.Q1.toFixed(4)}}</td>
            <td>${{d.Median.toFixed(4)}}</td>
            <td>${{d.Q3.toFixed(4)}}</td>
            <td>${{d.Max.toFixed(4)}}</td>
          </tr>
        `).join('')}}
      `;
      document.getElementById(elementId).innerHTML = html;
    }}

    // Inicializar
    gerarBoxPlot('boxplot-biomassa', dataBiomassa, 'Box Plot - Biomassa Final');
    gerarBoxPlot('boxplot-pam', dataPAM, 'Box Plot - PAM Final (Fv/Fm)');

    gerarTabela('table-biomassa', estBiomassa);
    gerarTabela('table-pam', estPAM);
  </script>
</body>
</html>
"""

# Salvar arquivo
with open("HTML/Dashboard_Estatisticas_BoxPlot.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("✓ Dashboard_Estatisticas_BoxPlot.html gerado com sucesso!")
print(f"  - {len(dados_biomassa)} dados de Biomassa")
print(f"  - {len(dados_pam)} dados de PAM")
print(f"  - {len(est_biomassa)} linhas de estatísticas de Biomassa")
print(f"  - {len(est_pam)} linhas de estatísticas de PAM")
