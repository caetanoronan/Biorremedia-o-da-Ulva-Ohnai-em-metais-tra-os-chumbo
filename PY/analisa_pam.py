import pandas as pd
import numpy as np
from scipy import stats
import matplotlib.pyplot as plt
import seaborn as sns
import re
import warnings
warnings.filterwarnings('ignore')

# Configuração
ods_file = "DADOS DO EXPERIMENTO BIORREMEDIAÇÃO - DORLAM'S.ods"

# Função para extrair grupo (C0, C1, C2, C3) da coluna FRASCOS
def extract_group(frascos_str):
    match = re.search(r'(C[0-3])', str(frascos_str))
    return match.group(1) if match else None

# ============================================================================
# 1. LEITURA DOS DADOS PAM
# ============================================================================

print("="*80)
print("ANÁLISE DE DADOS PAM (Fluorescência)")
print("="*80)

# Ler as 4 abas PAM
pam_initial_sem = pd.read_excel(ods_file, sheet_name='PAM INICIAL-sem NaNO3', engine='odf')
pam_final_sem = pd.read_excel(ods_file, sheet_name='PAM sem NaNO3', engine='odf')
pam_initial_com = pd.read_excel(ods_file, sheet_name='PAM 1- NaNO3', engine='odf')
pam_final_com = pd.read_excel(ods_file, sheet_name='PAM 2- NaNO3', engine='odf')

# Limpeza de colunas (remover espaço em branco)
for df in [pam_initial_sem, pam_final_sem, pam_initial_com, pam_final_com]:
    df.columns = df.columns.str.strip()

# ============================================================================
# 2. PROCESSAMENTO SEM NaNO3
# ============================================================================

print("\n" + "─"*80)
print("PAM SEM NaNO3")
print("─"*80)

# Extrair grupo e calcular média das replicatas
pam_sem = pd.DataFrame()
pam_sem['grupo'] = pam_initial_sem['FRASCOS'].apply(extract_group)
pam_sem['inicial'] = pam_initial_sem.iloc[:, 1:4].mean(axis=1)
pam_sem['final'] = pam_final_sem.iloc[:, 1:4].mean(axis=1)
pam_sem['mudanca'] = pam_sem['final'] - pam_sem['inicial']
pam_sem['mudanca_pct'] = ((pam_sem['final'] - pam_sem['inicial']) / pam_sem['inicial'] * 100).round(2)

# Remover linhas com NaN
pam_sem_clean = pam_sem.dropna()

print(f"\nDados processados: {len(pam_sem_clean)} amostras")
print("\nPrimeiras 10 linhas:")
print(pam_sem_clean.head(10))

# ============================================================================
# 3. PROCESSAMENTO COM NaNo3
# ============================================================================

print("\n" + "─"*80)
print("PAM COM NaNO3")
print("─"*80)

# Extrair grupo e calcular média das replicatas
pam_com = pd.DataFrame()
pam_com['grupo'] = pam_initial_com['FRASCOS'].apply(extract_group)
pam_com['inicial'] = pam_initial_com.iloc[:, 1:4].mean(axis=1)
pam_com['final'] = pam_final_com.iloc[:, 1:4].mean(axis=1)
pam_com['mudanca'] = pam_com['final'] - pam_com['inicial']
pam_com['mudanca_pct'] = ((pam_com['final'] - pam_com['inicial']) / pam_com['inicial'] * 100).round(2)

# Remover linhas com NaN
pam_com_clean = pam_com.dropna()

print(f"\nDados processados: {len(pam_com_clean)} amostras")
print("\nPrimeiras 10 linhas:")
print(pam_com_clean.head(10))

# ============================================================================
# 4. ESTATÍSTICAS DESCRITIVAS - SEM NaNO3
# ============================================================================

print("\n" + "="*80)
print("ESTATÍSTICAS DESCRITIVAS - PAM SEM NaNO3")
print("="*80)

for var in ['inicial', 'final', 'mudanca']:
    print(f"\n{var.upper()}:")
    stats_sem = pam_sem_clean.groupby('grupo')[var].describe()
    print(stats_sem)
    
    # Teste de normalidade (Shapiro-Wilk)
    print(f"\nTeste de Normalidade (Shapiro-Wilk) - {var}:")
    for grupo in sorted(pam_sem_clean['grupo'].unique()):
        data = pam_sem_clean[pam_sem_clean['grupo'] == grupo][var]
        stat, pval = stats.shapiro(data)
        normal = "✓ Normal" if pval > 0.05 else "✗ Não-normal"
        print(f"  {grupo}: p={pval:.4f} {normal}")

# ============================================================================
# 5. ESTATÍSTICAS DESCRITIVAS - COM NaNO3
# ============================================================================

print("\n" + "="*80)
print("ESTATÍSTICAS DESCRITIVAS - PAM COM NaNO3")
print("="*80)

for var in ['inicial', 'final', 'mudanca']:
    print(f"\n{var.upper()}:")
    stats_com = pam_com_clean.groupby('grupo')[var].describe()
    print(stats_com)
    
    # Teste de normalidade (Shapiro-Wilk)
    print(f"\nTeste de Normalidade (Shapiro-Wilk) - {var}:")
    for grupo in sorted(pam_com_clean['grupo'].unique()):
        data = pam_com_clean[pam_com_clean['grupo'] == grupo][var]
        stat, pval = stats.shapiro(data)
        normal = "✓ Normal" if pval > 0.05 else "✗ Não-normal"
        print(f"  {grupo}: p={pval:.4f} {normal}")

# ============================================================================
# 6. TESTES ESTATÍSTICOS - SEM NaNO3
# ============================================================================

print("\n" + "="*80)
print("TESTES ESTATÍSTICOS - PAM SEM NaNO3")
print("="*80)

for var in ['inicial', 'final', 'mudanca']:
    print(f"\n{var.upper()}:")
    
    # ANOVA
    groups = [pam_sem_clean[pam_sem_clean['grupo'] == g][var].values 
              for g in sorted(pam_sem_clean['grupo'].unique())]
    f_stat, f_pval = stats.f_oneway(*groups)
    print(f"  ANOVA: F={f_stat:.4f}, p={f_pval:.6f}", end="")
    print(" → SIGNIFICATIVO ✓" if f_pval < 0.05 else " → Não significativo")
    
    # Kruskal-Wallis
    h_stat, h_pval = stats.kruskal(*groups)
    print(f"  Kruskal-Wallis: H={h_stat:.4f}, p={h_pval:.6f}", end="")
    print(" → SIGNIFICATIVO ✓" if h_pval < 0.05 else " → Não significativo")

# ============================================================================
# 7. TESTES ESTATÍSTICOS - COM NaNO3
# ============================================================================

print("\n" + "="*80)
print("TESTES ESTATÍSTICOS - PAM COM NaNO3")
print("="*80)

for var in ['inicial', 'final', 'mudanca']:
    print(f"\n{var.upper()}:")
    
    # ANOVA
    groups = [pam_com_clean[pam_com_clean['grupo'] == g][var].values 
              for g in sorted(pam_com_clean['grupo'].unique())]
    f_stat, f_pval = stats.f_oneway(*groups)
    print(f"  ANOVA: F={f_stat:.4f}, p={f_pval:.6f}", end="")
    print(" → SIGNIFICATIVO ✓" if f_pval < 0.05 else " → Não significativo")
    
    # Kruskal-Wallis
    h_stat, h_pval = stats.kruskal(*groups)
    print(f"  Kruskal-Wallis: H={h_stat:.4f}, p={h_pval:.6f}", end="")
    print(" → SIGNIFICATIVO ✓" if h_pval < 0.05 else " → Não significativo")

# ============================================================================
# 8. SALVAR DADOS PROCESSADOS
# ============================================================================

print("\n" + "="*80)
print("SALVANDO ARQUIVOS")
print("="*80)

# CSV sem NaNO3
pam_sem_clean.to_csv('PAM_SEM_NaNO3.csv', index=False)
print("✓ PAM_SEM_NaNO3.csv")

# CSV com NaNO3
pam_com_clean.to_csv('PAM_COM_NaNO3.csv', index=False)
print("✓ PAM_COM_NaNO3.csv")

print("\n✓ Análise PAM concluída!")
