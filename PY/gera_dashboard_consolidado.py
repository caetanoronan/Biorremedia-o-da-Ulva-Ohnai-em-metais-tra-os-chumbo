import json
import pandas as pd

# Carrega dados comparativos
comp = pd.read_csv("Comparacao_COM_vs_SEM_NaNO3.csv")

biomassa = comp[comp["Variavel"] == "Biomassa Final"].copy()
pam = comp[comp["Variavel"] == "PAM Final (Fv/Fm)"].copy()

# Ordenacao por grupo
ordem = ["C0", "C1", "C2", "C3"]
biomassa["Grupo"] = pd.Categorical(biomassa["Grupo"], categories=ordem, ordered=True)
pam["Grupo"] = pd.Categorical(pam["Grupo"], categories=ordem, ordered=True)
biomassa = biomassa.sort_values("Grupo")
pam = pam.sort_values("Grupo")

# Achados principais
top_biomassa = biomassa.sort_values("Efeito_Pct", ascending=False).iloc[0]
top_pam = pam.sort_values("Efeito_Pct", ascending=False).iloc[0]

html = f"""<!DOCTYPE html>
<html lang=\"pt-BR\">
<head>
  <meta charset=\"UTF-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\" />
  <title>Dashboard Consolidado - Biorremediacao</title>
  <script src=\"libs/plotly-2.35.2.min.js\"></script>
  <style>
    * {{ margin: 0; padding: 0; box-sizing: border-box; }}
    body {{
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background: linear-gradient(135deg, #0f2027 0%, #203a43 50%, #2c5364 100%);
      color: #223;
      min-height: 100vh;
      padding: 20px;
    }}
    .container {{ max-width: 1400px; margin: 0 auto; }}
    header {{
      background: white;
      border-radius: 12px;
      padding: 28px;
      margin-bottom: 24px;
      box-shadow: 0 10px 25px rgba(0, 0, 0, 0.15);
    }}
    header h1 {{ color: #0d3b66; margin-bottom: 8px; font-size: 2.1em; }}
    header p {{ color: #4f647a; font-size: 1.05em; }}

    .cards {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
      gap: 16px;
      margin-bottom: 24px;
    }}
    .card {{
      background: white;
      border-radius: 12px;
      padding: 18px;
      box-shadow: 0 8px 20px rgba(0, 0, 0, 0.12);
      border-left: 6px solid #118ab2;
    }}
    .card h3 {{ color: #0d3b66; font-size: 0.95em; margin-bottom: 8px; text-transform: uppercase; }}
    .card .value {{ color: #073b4c; font-weight: 700; font-size: 1.8em; }}
    .card .sub {{ color: #5f7286; font-size: 0.9em; margin-top: 4px; }}

    .panel {{
      background: white;
      border-radius: 12px;
      padding: 20px;
      margin-bottom: 24px;
      box-shadow: 0 8px 20px rgba(0, 0, 0, 0.12);
      overflow: hidden;
    }}
    .panel h2 {{ color: #0d3b66; margin-bottom: 12px; font-size: 1.25em; }}

    .grid-2 {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(500px, 1fr));
      gap: 16px;
    }}

    .chart {{
      width: 100%;
      height: 360px;
      min-width: 0;
      overflow: hidden;
    }}

    .chart .js-plotly-plot,
    .chart .plot-container,
    .chart .svg-container {{
      width: 100% !important;
      height: 100% !important;
    }}

    table {{ width: 100%; border-collapse: collapse; }}
    th, td {{ padding: 10px; border-bottom: 1px solid #e7edf3; text-align: left; font-size: 0.92em; }}
    th {{ background: #f4f8fb; color: #0d3b66; }}

    .links {{ display: flex; gap: 10px; flex-wrap: wrap; margin-top: 10px; }}
    .btn {{
      display: inline-block;
      text-decoration: none;
      background: #0d6efd;
      color: white;
      padding: 10px 14px;
      border-radius: 8px;
      font-weight: 600;
      font-size: 0.9em;
    }}
    .btn:hover {{ background: #0b5ed7; }}

    @media (max-width: 768px) {{
      .chart {{ height: 320px; }}
      .grid-2 {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <div class=\"container\">
    <header>
      <h1>Dashboard Consolidado - Biorremediacao</h1>
      <p>Sintese integrada de Biomassa e PAM com comparacao entre tratamentos SEM vs COM NaNO3.</p>
    </header>

    <section class=\"cards\">
      <div class=\"card\">
        <h3>Total de analises</h3>
        <div class=\"value\">4 fases</div>
        <div class=\"sub\">Biomassa, PAM, comparacao e dashboards</div>
      </div>
      <div class=\"card\">
        <h3>Maior ganho em Biomassa</h3>
        <div class=\"value\">{top_biomassa['Efeito_Pct']:.2f}%</div>
        <div class=\"sub\">Grupo {top_biomassa['Grupo']} com NaNO3</div>
      </div>
      <div class=\"card\">
        <h3>Maior ganho em PAM Final</h3>
        <div class=\"value\">{top_pam['Efeito_Pct']:.2f}%</div>
        <div class=\"sub\">Grupo {top_pam['Grupo']} com NaNO3</div>
      </div>
      <div class=\"card\">
        <h3>Amostras comparadas</h3>
        <div class=\"value\">8 linhas</div>
        <div class=\"sub\">4 grupos x 2 variaveis finais</div>
      </div>
    </section>

    <section class=\"panel\">
      <h2>Efeito do NaNO3 por grupo</h2>
      <div class=\"grid-2\">
        <div>
          <h3 style=\"color:#0d3b66; margin-bottom:6px;\">Biomassa Final (Efeito %)</h3>
          <div id=\"chartBiomassaPct\" class=\"chart\"></div>
        </div>
        <div>
          <h3 style=\"color:#0d3b66; margin-bottom:6px;\">PAM Final (Efeito %)</h3>
          <div id=\"chartPamPct\" class=\"chart\"></div>
        </div>
      </div>
    </section>

    <section class=\"panel\">
      <h2>Tabela comparativa consolidada</h2>
      <table>
        <thead>
          <tr>
            <th>Grupo</th>
            <th>Variavel</th>
            <th>SEM NaNO3</th>
            <th>COM NaNO3</th>
            <th>Efeito</th>
            <th>Efeito %</th>
          </tr>
        </thead>
        <tbody id=\"tabelaComparativa\"></tbody>
      </table>
    </section>

    <section class=\"panel\">
      <h2>Navegacao dos dashboards detalhados</h2>
      <div class=\"links\">
        <a class=\"btn\" href=\"Dashboard_Biomassa_COM_NaNO3.html\" target=\"_blank\">Abrir Biomassa</a>
        <a class=\"btn\" href=\"Dashboard_PAM.html\" target=\"_blank\">Abrir PAM</a>
        <a class=\"btn\" href=\"Dashboard_Biorremediacao.html\" target=\"_blank\">Abrir Biorremediacao</a>
      </div>
    </section>
  </div>

  <script>
    const dadosBiomassa = {json.dumps(biomassa.to_dict(orient='records'))};
    const dadosPam = {json.dumps(pam.to_dict(orient='records'))};
    const dadosTabela = {json.dumps(comp.to_dict(orient='records'))};

    function plotar() {{
      const layoutBase = {{
        autosize: true,
        margin: {{ l: 45, r: 20, t: 15, b: 45 }},
        yaxis: {{ title: 'Efeito (%)' }},
        xaxis: {{ title: 'Grupo' }}
      }};

      Plotly.newPlot('chartBiomassaPct', [{{
        x: dadosBiomassa.map(d => d.Grupo),
        y: dadosBiomassa.map(d => d.Efeito_Pct),
        type: 'bar',
        marker: {{ color: '#118ab2' }},
        text: dadosBiomassa.map(d => `${{d.Efeito_Pct.toFixed(2)}}%`),
        textposition: 'outside'
      }}], layoutBase, {{ responsive: true }});

      Plotly.newPlot('chartPamPct', [{{
        x: dadosPam.map(d => d.Grupo),
        y: dadosPam.map(d => d.Efeito_Pct),
        type: 'bar',
        marker: {{ color: '#06d6a0' }},
        text: dadosPam.map(d => `${{d.Efeito_Pct.toFixed(2)}}%`),
        textposition: 'outside'
      }}], layoutBase, {{ responsive: true }});
    }}

    function preencherTabela() {{
      const tbody = document.getElementById('tabelaComparativa');
      dadosTabela.forEach(r => {{
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td><strong>${{r.Grupo}}</strong></td>
          <td>${{r.Variavel}}</td>
          <td>${{Number(r.Sem_NaNO3).toFixed(4)}}</td>
          <td>${{Number(r.Com_NaNO3).toFixed(4)}}</td>
          <td>${{Number(r.Efeito).toFixed(4)}}</td>
          <td>${{Number(r.Efeito_Pct).toFixed(2)}}%</td>
        `;
        tbody.appendChild(tr);
      }});
    }}

    function resizeCharts() {{
      ['chartBiomassaPct', 'chartPamPct'].forEach(id => {{
        const el = document.getElementById(id);
        if (el && el.data) Plotly.Plots.resize(el);
      }});
    }}

    document.addEventListener('DOMContentLoaded', function() {{
      preencherTabela();
      plotar();
      setTimeout(resizeCharts, 80);
    }});

    window.addEventListener('resize', resizeCharts);
  </script>
</body>
</html>
"""

with open("Dashboard_Consolidado_Biorremediacao.html", "w", encoding="utf-8") as f:
    f.write(html)

print("OK - Dashboard_Consolidado_Biorremediacao.html criado")
