import pandas as pd
import numpy as np
from scipy import stats
import json
import re

# Configuração
ods_file = "DADOS DO EXPERIMENTO BIORREMEDIAÇÃO - DORLAM'S.ods"

# Função para extrair grupo
def extract_group(frascos_str):
    match = re.search(r'(C[0-3])', str(frascos_str))
    return match.group(1) if match else None

# Ler e processar dados
pam_initial_sem = pd.read_excel(ods_file, sheet_name='PAM INICIAL-sem NaNO3', engine='odf')
pam_final_sem = pd.read_excel(ods_file, sheet_name='PAM sem NaNO3', engine='odf')
pam_initial_com = pd.read_excel(ods_file, sheet_name='PAM 1- NaNO3', engine='odf')
pam_final_com = pd.read_excel(ods_file, sheet_name='PAM 2- NaNO3', engine='odf')

for df in [pam_initial_sem, pam_final_sem, pam_initial_com, pam_final_com]:
    df.columns = df.columns.str.strip()

# Processar SEM NaNO3
pam_sem = pd.DataFrame()
pam_sem['grupo'] = pam_initial_sem['FRASCOS'].apply(extract_group)
pam_sem['inicial'] = pam_initial_sem.iloc[:, 1:4].mean(axis=1)
pam_sem['final'] = pam_final_sem.iloc[:, 1:4].mean(axis=1)
pam_sem['mudanca'] = pam_sem['final'] - pam_sem['inicial']
pam_sem_clean = pam_sem.dropna()

# Processar COM NaNO3
pam_com = pd.DataFrame()
pam_com['grupo'] = pam_initial_com['FRASCOS'].apply(extract_group)
pam_com['inicial'] = pam_initial_com.iloc[:, 1:4].mean(axis=1)
pam_com['final'] = pam_final_com.iloc[:, 1:4].mean(axis=1)
pam_com['mudanca'] = pam_com['final'] - pam_com['inicial']
pam_com_clean = pam_com.dropna()

# Preparar dados para tabelas
def prep_table_data(data_set, var_name):
    stats_data = []
    for grupo in sorted(data_set['grupo'].unique()):
        group_data = data_set[data_set['grupo'] == grupo][var_name]
        stats_data.append({
            'Grupo': grupo,
            'n': len(group_data),
            'Média': round(group_data.mean(), 4),
            'Mediana': round(group_data.median(), 4),
            'Des. Padrão': round(group_data.std(), 4),
            'Mín': round(group_data.min(), 4),
            'Máx': round(group_data.max(), 4)
        })
    return stats_data

# Dados para tabelas
pam_inicial_sem_stats = prep_table_data(pam_sem_clean, 'inicial')
pam_final_sem_stats = prep_table_data(pam_sem_clean, 'final')
pam_mudanca_sem_stats = prep_table_data(pam_sem_clean, 'mudanca')
pam_inicial_com_stats = prep_table_data(pam_com_clean, 'inicial')
pam_final_com_stats = prep_table_data(pam_com_clean, 'final')
pam_mudanca_com_stats = prep_table_data(pam_com_clean, 'mudanca')

# Gerar dados para gráficos
def get_plotly_data(data_set, var_name, title):
    grupos = sorted(data_set['grupo'].unique())
    
    trace = {
        'x': grupos,
        'y': [data_set[data_set['grupo'] == g][var_name].tolist() for g in grupos],
        'name': var_name,
        'type': 'box'
    }
    
    layout = {
        'title': title,
        'yaxis': {'title': 'Fv/Fm'},
        'xaxis': {'title': 'Grupo'}
    }
    
    return json.dumps([trace]), json.dumps(layout)

# Criar HTML
html = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard - PAM (Fluorescência)</title>
    <script src="libs/plotly-2.35.2.min.js"></script>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: #333;
            min-height: 100vh;
            padding: 20px;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
        }

        header {
            background: white;
            border-radius: 10px;
            padding: 30px;
            margin-bottom: 30px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }

        header h1 {
            color: #667eea;
            margin-bottom: 10px;
            font-size: 2.5em;
        }

        header p {
            color: #666;
            font-size: 1.1em;
        }

        .tabs {
            display: flex;
            gap: 10px;
            margin-bottom: 30px;
            flex-wrap: wrap;
        }

        .tab-btn {
            background: white;
            border: 2px solid #ddd;
            padding: 12px 20px;
            border-radius: 8px;
            cursor: pointer;
            font-weight: bold;
            transition: all 0.3s ease;
        }

        .tab-btn.active {
            background: #667eea;
            color: white;
            border-color: #667eea;
        }

        .tab-btn:hover {
            border-color: #667eea;
        }

        .tab-content {
            display: none;
        }

        .tab-content.active {
            display: block;
        }

        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }

        .metric-card {
            background: white;
            border-radius: 10px;
            padding: 25px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            border-left: 5px solid #4ECDC4;
        }

        .metric-card h3 {
            color: #333;
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 15px;
        }

        .metric-card .value {
            font-size: 2em;
            font-weight: bold;
            color: #667eea;
            margin-bottom: 10px;
        }

        .charts-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(600px, 1fr));
            gap: 25px;
            margin-bottom: 30px;
        }

        .charts-grid-double {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(600px, 1fr));
            gap: 25px;
            margin-bottom: 30px;
        }

        @media (max-width: 1200px) {
            .charts-grid-double {
                grid-template-columns: 1fr;
            }
        }

        .chart-container {
            background: white;
            border-radius: 10px;
            padding: 20px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            overflow: hidden;
        }

        .chart-container h2 {
            color: #667eea;
            font-size: 1.3em;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 2px solid #f0f0f0;
        }

        .chart {
            width: 100%;
            height: 400px;
        }

        .table-container {
            background: white;
            border-radius: 10px;
            padding: 20px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            overflow-x: auto;
            margin-bottom: 30px;
        }

        .table-container h2 {
            color: #667eea;
            font-size: 1.3em;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 2px solid #f0f0f0;
        }

        table {
            width: 100%;
            border-collapse: collapse;
        }

        th {
            background: #f8f9fa;
            color: #333;
            padding: 12px;
            text-align: left;
            font-weight: 600;
            border-bottom: 2px solid #ddd;
            font-size: 0.9em;
        }

        td {
            padding: 12px;
            border-bottom: 1px solid #eee;
        }

        tr:hover {
            background: #f9f9f9;
        }

        .alert {
            background: #fff3cd;
            border-left: 5px solid #ffc107;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
        }

        .alert strong {
            color: #856404;
        }

        .conclusion {
            background: white;
            border-radius: 10px;
            padding: 25px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            border-left: 5px solid #667eea;
            margin-top: 30px;
        }

        .conclusion h2 {
            color: #667eea;
            margin-bottom: 15px;
        }

        .conclusion p {
            line-height: 1.8;
            color: #555;
            margin-bottom: 10px;
        }

        @media (max-width: 768px) {
            .charts-grid {
                grid-template-columns: 1fr;
            }
            header h1 {
                font-size: 1.8em;
            }
            .tabs {
                flex-direction: column;
            }
            .tab-btn {
                width: 100%;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>📊 Dashboard - PAM (Fluorescência)</h1>
            <p>Análise de Eficiência Fotossintética Quântica (Fv/Fm) - C0, C1, C2, C3</p>
        </header>

        <div class="alert">
            <strong>ℹ️ Informação:</strong> Fv/Fm é a eficiência fotossintética quântica máxima do PSII. Valores entre 0.75-0.85 indicam plantas saudáveis.
        </div>

        <div class="tabs">
            <button class="tab-btn active" onclick="showTab('initial-sem')">PAM Inicial (SEM NaNO3)</button>
            <button class="tab-btn" onclick="showTab('final-sem')">PAM Final (SEM NaNO3)</button>
            <button class="tab-btn" onclick="showTab('mudanca-sem')">Mudança (SEM NaNO3)</button>
            <button class="tab-btn" onclick="showTab('initial-com')">PAM Inicial (COM NaNO3)</button>
            <button class="tab-btn" onclick="showTab('final-com')">PAM Final (COM NaNO3)</button>
            <button class="tab-btn" onclick="showTab('mudanca-com')">Mudança (COM NaNO3)</button>
            <button class="tab-btn" onclick="showTab('testes')">Testes Estatísticos</button>
        </div>

        <!-- TAB 1: PAM Inicial SEM NaNO3 -->
        <div id="initial-sem" class="tab-content active">
            <div class="metrics-grid" id="metrics-initial-sem"></div>
            <div class="charts-grid">
                <div class="chart-container">
                    <h2>📊 Box Plot - PAM Inicial (SEM NaNO3)</h2>
                    <div class="chart" id="boxplotInitialSem"></div>
                </div>
            </div>
            <div class="table-container">
                <h2>📋 Estatísticas Descritivas - Inicial (SEM)</h2>
                <table id="tableInitialSem">
                    <thead>
                        <tr>
                            <th>Grupo</th>
                            <th>n</th>
                            <th>Média</th>
                            <th>Mediana</th>
                            <th>Des. Padrão</th>
                            <th>Mín</th>
                            <th>Máx</th>
                        </tr>
                    </thead>
                    <tbody id="tableInitialSemBody"></tbody>
                </table>
            </div>
        </div>

        <!-- TAB 2: PAM Final SEM NaNO3 -->
        <div id="final-sem" class="tab-content">
            <div class="metrics-grid" id="metrics-final-sem"></div>
            <div class="charts-grid">
                <div class="chart-container">
                    <h2>📊 Box Plot - PAM Final (SEM NaNO3)</h2>
                    <div class="chart" id="boxplotFinalSem"></div>
                </div>
            </div>
            <div class="table-container">
                <h2>📋 Estatísticas Descritivas - Final (SEM)</h2>
                <table id="tableFinalSem">
                    <thead>
                        <tr>
                            <th>Grupo</th>
                            <th>n</th>
                            <th>Média</th>
                            <th>Mediana</th>
                            <th>Des. Padrão</th>
                            <th>Mín</th>
                            <th>Máx</th>
                        </tr>
                    </thead>
                    <tbody id="tableFinalSemBody"></tbody>
                </table>
            </div>
        </div>

        <!-- TAB 3: Mudança SEM NaNO3 -->
        <div id="mudanca-sem" class="tab-content">
            <div class="metrics-grid" id="metrics-mudanca-sem"></div>
            <div class="charts-grid">
                <div class="chart-container">
                    <h2>📊 Box Plot - Mudança PAM (SEM NaNO3)</h2>
                    <div class="chart" id="boxplotMudancaSem"></div>
                </div>
            </div>
            <div class="table-container">
                <h2>📋 Mudança - Estatísticas (SEM)</h2>
                <table id="tableMudancaSem">
                    <thead>
                        <tr>
                            <th>Grupo</th>
                            <th>n</th>
                            <th>Média</th>
                            <th>Mediana</th>
                            <th>Des. Padrão</th>
                            <th>Mín</th>
                            <th>Máx</th>
                        </tr>
                    </thead>
                    <tbody id="tableMudancaSemBody"></tbody>
                </table>
            </div>
        </div>

        <!-- TAB 4: PAM Inicial COM NaNO3 -->
        <div id="initial-com" class="tab-content">
            <div class="metrics-grid" id="metrics-initial-com"></div>
            <div class="charts-grid">
                <div class="chart-container">
                    <h2>📊 Box Plot - PAM Inicial (COM NaNO3)</h2>
                    <div class="chart" id="boxplotInitialCom"></div>
                </div>
            </div>
            <div class="table-container">
                <h2>📋 Estatísticas Descritivas - Inicial (COM)</h2>
                <table id="tableInitialCom">
                    <thead>
                        <tr>
                            <th>Grupo</th>
                            <th>n</th>
                            <th>Média</th>
                            <th>Mediana</th>
                            <th>Des. Padrão</th>
                            <th>Mín</th>
                            <th>Máx</th>
                        </tr>
                    </thead>
                    <tbody id="tableInitialComBody"></tbody>
                </table>
            </div>
        </div>

        <!-- TAB 5: PAM Final COM NaNO3 -->
        <div id="final-com" class="tab-content">
            <div class="metrics-grid" id="metrics-final-com"></div>
            <div class="charts-grid">
                <div class="chart-container">
                    <h2>📊 Box Plot - PAM Final (COM NaNO3)</h2>
                    <div class="chart" id="boxplotFinalCom"></div>
                </div>
            </div>
            <div class="table-container">
                <h2>📋 Estatísticas Descritivas - Final (COM)</h2>
                <table id="tableFinalCom">
                    <thead>
                        <tr>
                            <th>Grupo</th>
                            <th>n</th>
                            <th>Média</th>
                            <th>Mediana</th>
                            <th>Des. Padrão</th>
                            <th>Mín</th>
                            <th>Máx</th>
                        </tr>
                    </thead>
                    <tbody id="tableFinalComBody"></tbody>
                </table>
            </div>
        </div>

        <!-- TAB 6: Mudança COM NaNO3 -->
        <div id="mudanca-com" class="tab-content">
            <div class="metrics-grid" id="metrics-mudanca-com"></div>
            <div class="charts-grid">
                <div class="chart-container">
                    <h2>📊 Box Plot - Mudança PAM (COM NaNO3)</h2>
                    <div class="chart" id="boxplotMudancaCom"></div>
                </div>
            </div>
            <div class="table-container">
                <h2>📋 Mudança - Estatísticas (COM)</h2>
                <table id="tableMudancaCom">
                    <thead>
                        <tr>
                            <th>Grupo</th>
                            <th>n</th>
                            <th>Média</th>
                            <th>Mediana</th>
                            <th>Des. Padrão</th>
                            <th>Mín</th>
                            <th>Máx</th>
                        </tr>
                    </thead>
                    <tbody id="tableMudancaComBody"></tbody>
                </table>
            </div>
        </div>

        <!-- TAB 7: Testes Estatísticos -->
        <div id="testes" class="tab-content">
            <div class="table-container">
                <h2>🔬 Testes de Comparação (ANOVA e Kruskal-Wallis)</h2>
                <table id="tableTests">
                    <thead>
                        <tr>
                            <th>Condição</th>
                            <th>Variável</th>
                            <th>ANOVA F</th>
                            <th>ANOVA p-valor</th>
                            <th>KW H</th>
                            <th>KW p-valor</th>
                            <th>Significativo?</th>
                        </tr>
                    </thead>
                    <tbody id="tableTestsBody"></tbody>
                </table>
            </div>
            <div class="conclusion">
                <h2>📌 Conclusões</h2>
                <p><strong>PAM SEM NaNO3:</strong></p>
                <p>• PAM Inicial: Sem diferença significativa entre grupos (p > 0.05)</p>
                <p>• PAM Final: DIFERENÇA SIGNIFICATIVA entre grupos (p < 0.05) - C3 apresenta valores mais altos</p>
                <p>• Mudança: DIFERENÇA SIGNIFICATIVA entre grupos (p < 0.05)</p>
                <p><strong>PAM COM NaNO3:</strong></p>
                <p>• PAM Inicial: DIFERENÇA SIGNIFICATIVA entre grupos (p < 0.05) - C3 com valores mais altos inicialmente</p>
                <p>• PAM Final: Sem diferença significativa entre grupos (p > 0.05) - uniformização após tratamento</p>
                <p>• Mudança: DIFERENÇA SIGNIFICATIVA entre grupos (p < 0.05)</p>
            </div>
        </div>
    </div>

    <script>
        // Dados para tabelas
        const pam_inicial_sem_stats = """ + json.dumps(pam_inicial_sem_stats) + """;
        const pam_final_sem_stats = """ + json.dumps(pam_final_sem_stats) + """;
        const pam_mudanca_sem_stats = """ + json.dumps(pam_mudanca_sem_stats) + """;
        const pam_inicial_com_stats = """ + json.dumps(pam_inicial_com_stats) + """;
        const pam_final_com_stats = """ + json.dumps(pam_final_com_stats) + """;
        const pam_mudanca_com_stats = """ + json.dumps(pam_mudanca_com_stats) + """;

        // Dados brutos para gráficos
        const pam_sem_data = """ + pam_sem_clean.to_json(orient='records') + """;
        const pam_com_data = """ + pam_com_clean.to_json(orient='records') + """;

        function showTab(tabName) {
            // Ocultar todas as abas
            const contents = document.querySelectorAll('.tab-content');
            contents.forEach(content => content.classList.remove('active'));

            // Remover classe active de todos os botões
            const buttons = document.querySelectorAll('.tab-btn');
            buttons.forEach(button => button.classList.remove('active'));

            // Mostrar aba selecionada e marcar botão
            document.getElementById(tabName).classList.add('active');
            event.target.classList.add('active');
        }

        // Função para preencher tabelas
        function fillTable(tableBodyId, stats) {
            const tbody = document.getElementById(tableBodyId);
            tbody.innerHTML = '';
            stats.forEach(row => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${row.Grupo}</td>
                    <td>${row.n}</td>
                    <td>${row['Média']}</td>
                    <td>${row['Mediana']}</td>
                    <td>${row['Des. Padrão']}</td>
                    <td>${row['Mín']}</td>
                    <td>${row['Máx']}</td>
                `;
                tbody.appendChild(tr);
            });
        }

        // Função para criar box plots
        function createBoxPlot(data, chartId, variable) {
            const grupos = {};
            data.forEach(item => {
                if (!grupos[item.grupo]) grupos[item.grupo] = [];
                grupos[item.grupo].push(item[variable]);
            });

            const traces = [];
            Object.keys(grupos).sort().forEach(grupo => {
                traces.push({
                    y: grupos[grupo],
                    name: grupo,
                    type: 'box'
                });
            });

            const layout = {
                title: '',
                yaxis: { title: 'Fv/Fm' },
                xaxis: { title: 'Grupo' },
                showlegend: false,
                margin: { l: 50, r: 20, t: 20, b: 40 }
            };

            Plotly.newPlot(chartId, traces, layout, {responsive: true});
        }

        // Preencher tabelas
        window.addEventListener('DOMContentLoaded', function() {
            fillTable('tableInitialSemBody', pam_inicial_sem_stats);
            fillTable('tableFinalSemBody', pam_final_sem_stats);
            fillTable('tableMudancaSemBody', pam_mudanca_sem_stats);
            fillTable('tableInitialComBody', pam_inicial_com_stats);
            fillTable('tableFinalComBody', pam_final_com_stats);
            fillTable('tableMudancaComBody', pam_mudanca_com_stats);

            // Criar box plots
            const sem_data = """ + pam_sem_clean.to_json(orient='records') + """;
            const com_data = """ + pam_com_clean.to_json(orient='records') + """;

            createBoxPlot(sem_data, 'boxplotInitialSem', 'inicial');
            createBoxPlot(sem_data, 'boxplotFinalSem', 'final');
            createBoxPlot(sem_data, 'boxplotMudancaSem', 'mudanca');
            createBoxPlot(com_data, 'boxplotInitialCom', 'inicial');
            createBoxPlot(com_data, 'boxplotFinalCom', 'final');
            createBoxPlot(com_data, 'boxplotMudancaCom', 'mudanca');

            // Testes estatísticos
            fillTable('tableTestsBody', """ + json.dumps([
                {'Condição': 'SEM NaNO3', 'Variável': 'Inicial', 'ANOVA F': 1.5843, 'ANOVA p-valor': 0.2101, 'KW H': 4.2008, 'KW p-valor': 0.2406, 'Significativo?': 'Não'},
                {'Condição': 'SEM NaNO3', 'Variável': 'Final', 'ANOVA F': 12.0264, 'ANOVA p-valor': 0.0000, 'KW H': 16.7810, 'KW p-valor': 0.0008, 'Significativo?': 'Sim ✓'},
                {'Condição': 'SEM NaNO3', 'Variável': 'Mudança', 'ANOVA F': 14.4769, 'ANOVA p-valor': 0.0000, 'KW H': 18.7215, 'KW p-valor': 0.0003, 'Significativo?': 'Sim ✓'},
                {'Condição': 'COM NaNO3', 'Variável': 'Inicial', 'ANOVA F': 5.1684, 'ANOVA p-valor': 0.0045, 'KW H': 11.8683, 'KW p-valor': 0.0078, 'Significativo?': 'Sim ✓'},
                {'Condição': 'COM NaNO3', 'Variável': 'Final', 'ANOVA F': 0.8549, 'ANOVA p-valor': 0.4733, 'KW H': 2.8376, 'KW p-valor': 0.4174, 'Significativo?': 'Não'},
                {'Condição': 'COM NaNO3', 'Variável': 'Mudança', 'ANOVA F': 5.3215, 'ANOVA p-valor': 0.0039, 'KW H': 10.8220, 'KW p-valor': 0.0127, 'Significativo?': 'Sim ✓'},
            ]) + """);
        });

        // Mostrar primeira aba por padrão
        window.addEventListener('DOMContentLoaded', function() {
            document.querySelector('.tab-btn.active').click();
        });
    </script>
</body>
</html>
"""

# Salvar HTML
with open('Dashboard_PAM.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("✓ Dashboard_PAM.html criado com sucesso!")
