import pandas as pd
import numpy as np
from scipy import stats
import re
import warnings
warnings.filterwarnings('ignore')

# Configuração
source_file = "DADOS DO EXPERIMENTO BIORREMEDIAÇÃO - DORLAM'S.xls"

def extract_group(frascos_str):
    txt = str(frascos_str).upper().strip().replace("CO", "C0")
    match = re.search(r'(C[0-3])', txt)
    return match.group(1) if match else None

def calcular_estatisticas(dados):
    """Retorna dict com todas as estatísticas descritivas"""
    if len(dados) == 0:
        return None
    return {
        'Mean': round(dados.mean(), 4),
        'Std': round(dados.std(), 4),
        'Min': round(dados.min(), 4),
        'Q1': round(np.percentile(dados, 25), 4),
        'Median': round(np.percentile(dados, 50), 4),
        'Q3': round(np.percentile(dados, 75), 4),
        'Max': round(dados.max(), 4),
        'N': len(dados)
    }

print("="*80)
print("ANÁLISE COMPARATIVA: EFEITO DO TRATAMENTO COM NaNO3")
print("="*80)

# ============================================================================
# 1. LEITURA E PROCESSAMENTO - BIOMASSA
# ============================================================================

print("\n" + "─"*80)
print("PROCESSANDO DADOS DE BIOMASSA")
print("─"*80)

biomasssa_inicial_sem = pd.read_excel(source_file, sheet_name='Biomassa inicial- sem NaNO3')
biomassa_final_sem = pd.read_excel(source_file, sheet_name='Biomassa Final sem NaNO3', header=1)
biomassa_inicial_com = pd.read_excel(source_file, sheet_name='Biomassa inicial  com NaNO3')
biomassa_final_com = pd.read_excel(source_file, sheet_name='Biomassa final com NaNO3')

for df in [biomasssa_inicial_sem, biomassa_final_sem, biomassa_inicial_com, biomassa_final_com]:
    df.columns = df.columns.str.strip()

biomassa_sem = pd.DataFrame()
biomassa_sem['grupo'] = biomasssa_inicial_sem['TRATAMENTOS'].apply(extract_group)
biomassa_sem['inicial'] = pd.to_numeric(biomasssa_inicial_sem['INICIAL'], errors='coerce')
biomassa_sem['final'] = pd.to_numeric(biomassa_final_sem['FINAL'], errors='coerce')
biomassa_sem['crescimento'] = biomassa_sem['final'] - biomassa_sem['inicial']
biomassa_sem_clean = biomassa_sem.dropna()

biomassa_com = pd.DataFrame()
biomassa_com['grupo'] = biomassa_inicial_com['TRATAMENTOS'].apply(extract_group)
biomassa_com['inicial'] = pd.to_numeric(biomassa_inicial_com['BIOMASSA'], errors='coerce')
biomassa_com['final'] = pd.to_numeric(biomassa_final_com['PESO'], errors='coerce')
biomassa_com['crescimento'] = biomassa_com['final'] - biomassa_com['inicial']
biomassa_com_clean = biomassa_com[(biomassa_com['final'] < 1000)].dropna()

print(f"Biomassa SEM NaNO3: {len(biomassa_sem_clean)} amostras")
print(f"Biomassa COM NaNO3: {len(biomassa_com_clean)} amostras (após remoção de outlier)")

# ============================================================================
# 2. LEITURA E PROCESSAMENTO - PAM
# ============================================================================

print("\n" + "─"*80)
print("PROCESSANDO DADOS DE PAM")
print("─"*80)

pam_initial_sem = pd.read_excel(source_file, sheet_name='PAM INICIAL sem NaNO3')
pam_final_sem = pd.read_excel(source_file, sheet_name='PAM FINAL sem NaNO3')
pam_initial_com = pd.read_excel(source_file, sheet_name='PAM INICIAL COM NaNO3')
pam_final_com = pd.read_excel(source_file, sheet_name='PAM FINAL COM NaNO3')

for df in [pam_initial_sem, pam_final_sem, pam_initial_com, pam_final_com]:
    df.columns = df.columns.str.strip()

pam_sem = pd.DataFrame()
pam_sem['grupo'] = pam_initial_sem['TRATAMENTOS'].apply(extract_group)
pam_sem['inicial'] = pam_initial_sem.iloc[:, -3:].apply(pd.to_numeric, errors='coerce').mean(axis=1)
pam_sem['final'] = pam_final_sem.iloc[:, -3:].apply(pd.to_numeric, errors='coerce').mean(axis=1)
pam_sem['mudanca'] = pam_sem['final'] - pam_sem['inicial']
pam_sem_clean = pam_sem.dropna()

pam_com = pd.DataFrame()
pam_com['grupo'] = pam_initial_com['TRATAMENTOS'].apply(extract_group)
pam_com['inicial'] = pam_initial_com.iloc[:, -3:].apply(pd.to_numeric, errors='coerce').mean(axis=1)
pam_com['final'] = pam_final_com.iloc[:, -3:].apply(pd.to_numeric, errors='coerce').mean(axis=1)
pam_com['mudanca'] = pam_com['final'] - pam_com['inicial']
pam_com_clean = pam_com.dropna()

print(f"PAM SEM NaNO3: {len(pam_sem_clean)} amostras")
print(f"PAM COM NaNO3: {len(pam_com_clean)} amostras")

# ============================================================================
# 3. CRIAR DATAFRAME DE COMPARAÇÃO (formato anterior)
# ============================================================================

print("\n" + "="*80)
print("SALVANDO RESULTADOS")
print("="*80)

comparacao_data = []

for grupo in sorted(biomassa_sem_clean['grupo'].unique()):
    sem_final = biomassa_sem_clean[biomassa_sem_clean['grupo'] == grupo]['final'].values
    com_final = biomassa_com_clean[biomassa_com_clean['grupo'] == grupo]['final'].values
    pam_sem_final = pam_sem_clean[pam_sem_clean['grupo'] == grupo]['final'].values
    pam_com_final = pam_com_clean[pam_com_clean['grupo'] == grupo]['final'].values
    
    if len(sem_final) > 0 and len(com_final) > 0:
        comparacao_data.append({
            'Grupo': grupo,
            'Variavel': 'Biomassa Final',
            'Sem_NaNO3': round(sem_final.mean(), 4),
            'Com_NaNO3': round(com_final.mean(), 4),
            'Efeito': round(com_final.mean() - sem_final.mean(), 4),
            'Efeito_Pct': round((com_final.mean() - sem_final.mean()) / sem_final.mean() * 100, 2)
        })
    
    if len(pam_sem_final) > 0 and len(pam_com_final) > 0:
        comparacao_data.append({
            'Grupo': grupo,
            'Variavel': 'PAM Final (Fv/Fm)',
            'Sem_NaNO3': round(pam_sem_final.mean(), 4),
            'Com_NaNO3': round(pam_com_final.mean(), 4),
            'Efeito': round(pam_com_final.mean() - pam_sem_final.mean(), 4),
            'Efeito_Pct': round((pam_com_final.mean() - pam_sem_final.mean()) / pam_sem_final.mean() * 100, 2) if pam_sem_final.mean() != 0 else 0
        })

df_comparacao = pd.DataFrame(comparacao_data)
df_comparacao.to_csv("CSV/Comparacao_COM_vs_SEM_NaNO3.csv", index=False)
print("✓ Comparacao_COM_vs_SEM_NaNO3.csv")

# ============================================================================
# 4. CRIAR DATAFRAME DE ESTATÍSTICAS DESCRITIVAS COMPLETAS
# ============================================================================

estatisticas_data = []

for grupo in sorted(biomassa_sem_clean['grupo'].unique()):
    sem_final = biomassa_sem_clean[biomassa_sem_clean['grupo'] == grupo]['final'].values
    com_final = biomassa_com_clean[biomassa_com_clean['grupo'] == grupo]['final'].values
    pam_sem_final = pam_sem_clean[pam_sem_clean['grupo'] == grupo]['final'].values
    pam_com_final = pam_com_clean[pam_com_clean['grupo'] == grupo]['final'].values
    
    if len(sem_final) > 0:
        est = calcular_estatisticas(sem_final)
        estatisticas_data.append({
            'Variavel': 'Biomassa Final',
            'Tratamento': 'Sem NaNO3',
            'Grupo': grupo,
            'N': est['N'],
            'Mean': est['Mean'],
            'Std': est['Std'],
            'Min': est['Min'],
            'Q1': est['Q1'],
            'Median': est['Median'],
            'Q3': est['Q3'],
            'Max': est['Max']
        })
    
    if len(com_final) > 0:
        est = calcular_estatisticas(com_final)
        estatisticas_data.append({
            'Variavel': 'Biomassa Final',
            'Tratamento': 'Com NaNO3',
            'Grupo': grupo,
            'N': est['N'],
            'Mean': est['Mean'],
            'Std': est['Std'],
            'Min': est['Min'],
            'Q1': est['Q1'],
            'Median': est['Median'],
            'Q3': est['Q3'],
            'Max': est['Max']
        })
    
    if len(pam_sem_final) > 0:
        est = calcular_estatisticas(pam_sem_final)
        estatisticas_data.append({
            'Variavel': 'PAM Final (Fv/Fm)',
            'Tratamento': 'Sem NaNO3',
            'Grupo': grupo,
            'N': est['N'],
            'Mean': est['Mean'],
            'Std': est['Std'],
            'Min': est['Min'],
            'Q1': est['Q1'],
            'Median': est['Median'],
            'Q3': est['Q3'],
            'Max': est['Max']
        })
    
    if len(pam_com_final) > 0:
        est = calcular_estatisticas(pam_com_final)
        estatisticas_data.append({
            'Variavel': 'PAM Final (Fv/Fm)',
            'Tratamento': 'Com NaNO3',
            'Grupo': grupo,
            'N': est['N'],
            'Mean': est['Mean'],
            'Std': est['Std'],
            'Min': est['Min'],
            'Q1': est['Q1'],
            'Median': est['Median'],
            'Q3': est['Q3'],
            'Max': est['Max']
        })

df_estatisticas = pd.DataFrame(estatisticas_data)
df_estatisticas.to_csv("CSV/Estatisticas_Descritivas.csv", index=False)
print("✓ Estatisticas_Descritivas.csv")

# ============================================================================
# 5. CRIAR DATAFRAME DE DADOS PARA BOX PLOT
# ============================================================================

boxplot_data = []

for grupo in sorted(biomassa_sem_clean['grupo'].unique()):
    sem_final = biomassa_sem_clean[biomassa_sem_clean['grupo'] == grupo]['final'].values
    com_final = biomassa_com_clean[biomassa_com_clean['grupo'] == grupo]['final'].values
    pam_sem_final = pam_sem_clean[pam_sem_clean['grupo'] == grupo]['final'].values
    pam_com_final = pam_com_clean[pam_com_clean['grupo'] == grupo]['final'].values
    
    for val in sem_final:
        boxplot_data.append({'Variavel': 'Biomassa Final', 'Tratamento': 'Sem NaNO3', 'Grupo': grupo, 'Valor': round(val, 4)})
    for val in com_final:
        boxplot_data.append({'Variavel': 'Biomassa Final', 'Tratamento': 'Com NaNO3', 'Grupo': grupo, 'Valor': round(val, 4)})
    for val in pam_sem_final:
        boxplot_data.append({'Variavel': 'PAM Final (Fv/Fm)', 'Tratamento': 'Sem NaNO3', 'Grupo': grupo, 'Valor': round(val, 4)})
    for val in pam_com_final:
        boxplot_data.append({'Variavel': 'PAM Final (Fv/Fm)', 'Tratamento': 'Com NaNO3', 'Grupo': grupo, 'Valor': round(val, 4)})

df_boxplot = pd.DataFrame(boxplot_data)
df_boxplot.to_csv("CSV/Dados_BoxPlot.csv", index=False)
print("✓ Dados_BoxPlot.csv")

print("\n✓ Análise comparativa concluída com sucesso!")
print(f"\nArquivos gerados:")
print(f"  - Estatisticas_Descritivas.csv (com min, Q1, mediana, Q3, max)")
print(f"  - Dados_BoxPlot.csv (para visualizações)")
print(f"  - Comparacao_COM_vs_SEM_NaNO3.csv (comparação simples)")
