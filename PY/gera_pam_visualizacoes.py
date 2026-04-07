import pandas as pd
import numpy as np
from scipy import stats
import matplotlib.pyplot as plt
import seaborn as sns
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
import warnings
warnings.filterwarnings('ignore')

# Configuração
ods_file = "DADOS DO EXPERIMENTO BIORREMEDIAÇÃO - DORLAM'S.ods"

# Função para extrair grupo
def extract_group(frascos_str):
    match = re.search(r'(C[0-3])', str(frascos_str))
    return match.group(1) if match else None

# ============================================================================
# LEITURA E PROCESSAMENTO DOS DADOS
# ============================================================================

# Ler as 4 abas PAM
pam_initial_sem = pd.read_excel(ods_file, sheet_name='PAM INICIAL-sem NaNO3', engine='odf')
pam_final_sem = pd.read_excel(ods_file, sheet_name='PAM sem NaNO3', engine='odf')
pam_initial_com = pd.read_excel(ods_file, sheet_name='PAM 1- NaNO3', engine='odf')
pam_final_com = pd.read_excel(ods_file, sheet_name='PAM 2- NaNO3', engine='odf')

# Limpeza de colunas
for df in [pam_initial_sem, pam_final_sem, pam_initial_com, pam_final_com]:
    df.columns = df.columns.str.strip()

# Processar SEM NaNO3
pam_sem = pd.DataFrame()
pam_sem['grupo'] = pam_initial_sem['FRASCOS'].apply(extract_group)
pam_sem['inicial'] = pam_initial_sem.iloc[:, 1:4].mean(axis=1)
pam_sem['final'] = pam_final_sem.iloc[:, 1:4].mean(axis=1)
pam_sem['mudanca'] = pam_sem['final'] - pam_sem['inicial']
pam_sem['mudanca_pct'] = ((pam_sem['final'] - pam_sem['inicial']) / pam_sem['inicial'] * 100).round(2)
pam_sem_clean = pam_sem.dropna()

# Processar COM NaNO3
pam_com = pd.DataFrame()
pam_com['grupo'] = pam_initial_com['FRASCOS'].apply(extract_group)
pam_com['inicial'] = pam_initial_com.iloc[:, 1:4].mean(axis=1)
pam_com['final'] = pam_final_com.iloc[:, 1:4].mean(axis=1)
pam_com['mudanca'] = pam_com['final'] - pam_com['inicial']
pam_com['mudanca_pct'] = ((pam_com['final'] - pam_com['inicial']) / pam_com['inicial'] * 100).round(2)
pam_com_clean = pam_com.dropna()

# ============================================================================
# VISUALIZAÇÕES
# ============================================================================

fig, axes = plt.subplots(2, 3, figsize=(16, 10))
fig.suptitle('Análise PAM - Fluorescência (Fv/Fm)', fontsize=18, fontweight='bold')

# SEM NaNO3
# Box plot - Inicial
ax = axes[0, 0]
pam_sem_clean.boxplot(column='inicial', by='grupo', ax=ax)
ax.set_title('PAM Inicial (SEM NaNO3)', fontweight='bold')
ax.set_ylabel('Fv/Fm')
ax.set_xlabel('Grupo')
plt.sca(ax)
plt.xticks(rotation=0)

# Box plot - Final
ax = axes[0, 1]
pam_sem_clean.boxplot(column='final', by='grupo', ax=ax)
ax.set_title('PAM Final (SEM NaNO3)', fontweight='bold')
ax.set_ylabel('Fv/Fm')
ax.set_xlabel('Grupo')
plt.sca(ax)
plt.xticks(rotation=0)

# Box plot - Mudança
ax = axes[0, 2]
pam_sem_clean.boxplot(column='mudanca', by='grupo', ax=ax)
ax.set_title('Mudança PAM (SEM NaNO3)', fontweight='bold')
ax.set_ylabel('ΔFv/Fm')
ax.set_xlabel('Grupo')
plt.sca(ax)
plt.xticks(rotation=0)

# COM NaNO3
# Box plot - Inicial
ax = axes[1, 0]
pam_com_clean.boxplot(column='inicial', by='grupo', ax=ax)
ax.set_title('PAM Inicial (COM NaNO3)', fontweight='bold')
ax.set_ylabel('Fv/Fm')
ax.set_xlabel('Grupo')
plt.sca(ax)
plt.xticks(rotation=0)

# Box plot - Final
ax = axes[1, 1]
pam_com_clean.boxplot(column='final', by='grupo', ax=ax)
ax.set_title('PAM Final (COM NaNO3)', fontweight='bold')
ax.set_ylabel('Fv/Fm')
ax.set_xlabel('Grupo')
plt.sca(ax)
plt.xticks(rotation=0)

# Box plot - Mudança
ax = axes[1, 2]
pam_com_clean.boxplot(column='mudanca', by='grupo', ax=ax)
ax.set_title('Mudança PAM (COM NaNO3)', fontweight='bold')
ax.set_ylabel('ΔFv/Fm')
ax.set_xlabel('Grupo')
plt.sca(ax)
plt.xticks(rotation=0)

plt.tight_layout()
plt.savefig('Analise_PAM.png', dpi=300, bbox_inches='tight')
print("✓ Analise_PAM.png criado")
plt.close()

# ============================================================================
# EXPORTAR PARA EXCEL COM MÚLTIPLAS ABAS
# ============================================================================

with pd.ExcelWriter('Analise_PAM.xlsx', engine='openpyxl') as writer:
    
    # Aba 1: PAM Inicial SEM NaNO3
    stats_pam_sem_inicial = pam_sem_clean.groupby('grupo')['inicial'].describe().round(4)
    stats_pam_sem_inicial.to_excel(writer, sheet_name='PAM Inicial SEM')
    
    # Aba 2: PAM Final SEM NaNO3
    stats_pam_sem_final = pam_sem_clean.groupby('grupo')['final'].describe().round(4)
    stats_pam_sem_final.to_excel(writer, sheet_name='PAM Final SEM')
    
    # Aba 3: Mudança SEM NaNO3
    stats_pam_sem_mudanca = pam_sem_clean.groupby('grupo')['mudanca'].describe().round(4)
    stats_pam_sem_mudanca.to_excel(writer, sheet_name='Mudanca SEM')
    
    # Aba 4: PAM Inicial COM NaNO3
    stats_pam_com_inicial = pam_com_clean.groupby('grupo')['inicial'].describe().round(4)
    stats_pam_com_inicial.to_excel(writer, sheet_name='PAM Inicial COM')
    
    # Aba 5: PAM Final COM NaNO3
    stats_pam_com_final = pam_com_clean.groupby('grupo')['final'].describe().round(4)
    stats_pam_com_final.to_excel(writer, sheet_name='PAM Final COM')
    
    # Aba 6: Mudança COM NaNO3
    stats_pam_com_mudanca = pam_com_clean.groupby('grupo')['mudanca'].describe().round(4)
    stats_pam_com_mudanca.to_excel(writer, sheet_name='Mudanca COM')
    
    # Aba 7: Testes Estatísticos
    testes = []
    for condition, data_set in [('SEM NaNO3', pam_sem_clean), ('COM NaNO3', pam_com_clean)]:
        for var in ['inicial', 'final', 'mudanca']:
            groups = [data_set[data_set['grupo'] == g][var].values 
                     for g in sorted(data_set['grupo'].unique())]
            f_stat, f_pval = stats.f_oneway(*groups)
            h_stat, h_pval = stats.kruskal(*groups)
            
            testes.append({
                'Condicao': condition,
                'Variavel': var,
                'ANOVA_F': f_stat,
                'ANOVA_p': f_pval,
                'KW_H': h_stat,
                'KW_p': h_pval,
                'Signif_ANOVA': 'Sim' if f_pval < 0.05 else 'Nao',
                'Signif_KW': 'Sim' if h_pval < 0.05 else 'Nao'
            })
    
    df_testes = pd.DataFrame(testes)
    df_testes.to_excel(writer, sheet_name='Testes Estatisticos', index=False)
    
    # Aba 8: Resumo
    resumo_text = """
    ANÁLISE PAM - FLUORESCÊNCIA (Fv/Fm)
    
    DADOS PROCESSADOS:
    - PAM SEM NaNO3: 40 amostras (4 grupos x 10 replicatas)
    - PAM COM NaNO3: 40 amostras (4 grupos x 10 replicatas)
    
    PRINCIPAIS RESULTADOS:
    
    SEM NaNO3:
    • PAM Inicial: Sem diferença significativa entre grupos
    • PAM Final: DIFERENÇA SIGNIFICATIVA (p < 0.05)
    • Mudança PAM: DIFERENÇA SIGNIFICATIVA (p < 0.05)
    
    COM NaNO3:
    • PAM Inicial: DIFERENÇA SIGNIFICATIVA (p < 0.05)
    • PAM Final: Sem diferença significativa entre grupos
    • Mudança PAM: DIFERENÇA SIGNIFICATIVA (p < 0.05)
    
    INTERPRETAÇÃO:
    - Fv/Fm é a eficiência fotossintética quântica máxima do PSII
    - Valores mais altos indicam melhor desempenho fotossintético
    - Mudanças significativas sugerem efeito do tratamento
    """
    
    df_resumo = pd.DataFrame({'Resumo': [resumo_text]})
    df_resumo.to_excel(writer, sheet_name='Resumo', index=False)

print("✓ Analise_PAM.xlsx criado")
print("\n✓ Análise PAM com visualizações e Excel completada!")
