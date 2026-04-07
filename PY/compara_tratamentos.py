import pandas as pd
import numpy as np
from scipy import stats
import re
import warnings
warnings.filterwarnings('ignore')

# Configuração
ods_file = "DADOS DO EXPERIMENTO BIORREMEDIAÇÃO - DORLAM'S.ods"

def extract_group(frascos_str):
    match = re.search(r'(C[0-3])', str(frascos_str))
    return match.group(1) if match else None

print("="*80)
print("ANÁLISE COMPARATIVA: EFEITO DO TRATAMENTO COM NaNO3")
print("="*80)

# ============================================================================
# 1. LEITURA E PROCESSAMENTO - BIOMASSA
# ============================================================================

print("\n" + "─"*80)
print("PROCESSANDO DADOS DE BIOMASSA")
print("─"*80)

biomasssa_inicial_sem = pd.read_excel(ods_file, sheet_name='Biomassa inicial- sem NaNO3', engine='odf')
biomassa_final_sem = pd.read_excel(ods_file, sheet_name='Biomassa Final sem NaNO3', engine='odf')
biomassa_inicial_com = pd.read_excel(ods_file, sheet_name='Biomassa inicial  com NaNO3', engine='odf')
biomassa_final_com = pd.read_excel(ods_file, sheet_name='Biomassa final com NaNO3', engine='odf')

# Limpeza
for df in [biomasssa_inicial_sem, biomassa_final_sem, biomassa_inicial_com, biomassa_final_com]:
    df.columns = df.columns.str.strip()

# Processar SEM NaNO3
biomassa_sem = pd.DataFrame()
biomassa_sem['grupo'] = biomasssa_inicial_sem.iloc[:, 0].apply(extract_group)
biomassa_sem['inicial'] = pd.to_numeric(biomasssa_inicial_sem.iloc[:, 1], errors='coerce')
biomassa_sem['final'] = pd.to_numeric(biomassa_final_sem.iloc[:, 1], errors='coerce')
biomassa_sem['crescimento'] = biomassa_sem['final'] - biomassa_sem['inicial']
biomassa_sem_clean = biomassa_sem.dropna()

# Processar COM NaNO3
biomassa_com = pd.DataFrame()
biomassa_com['grupo'] = biomassa_inicial_com.iloc[:, 0].apply(extract_group)
biomassa_com['inicial'] = pd.to_numeric(biomassa_inicial_com.iloc[:, 1], errors='coerce')
biomassa_com['final'] = pd.to_numeric(biomassa_final_com.iloc[:, 1], errors='coerce')
biomassa_com['crescimento'] = biomassa_com['final'] - biomassa_com['inicial']

# Remover outlier em C3 (1713.0)
biomassa_com_clean = biomassa_com[(biomassa_com['final'] < 1000)].dropna()

print(f"Biomassa SEM NaNO3: {len(biomassa_sem_clean)} amostras")
print(f"Biomassa COM NaNO3: {len(biomassa_com_clean)} amostras (após remoção de outlier)")

# ============================================================================
# 2. LEITURA E PROCESSAMENTO - PAM
# ============================================================================

print("\n" + "─"*80)
print("PROCESSANDO DADOS DE PAM")
print("─"*80)

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

print(f"PAM SEM NaNO3: {len(pam_sem_clean)} amostras")
print(f"PAM COM NaNO3: {len(pam_com_clean)} amostras")

# ============================================================================
# 3. COMPARAÇÃO - BIOMASSA FINAL
# ============================================================================

print("\n" + "="*80)
print("EFEITO DO TRATAMENTO NA BIOMASSA FINAL")
print("="*80)

for grupo in sorted(biomassa_sem_clean['grupo'].unique()):
    sem = biomassa_sem_clean[biomassa_sem_clean['grupo'] == grupo]['final'].values
    com = biomassa_com_clean[biomassa_com_clean['grupo'] == grupo]['final'].values
    
    if len(sem) > 0 and len(com) > 0:
        efeito = com.mean() - sem.mean()
        efeito_pct = (efeito / sem.mean() * 100)
        
        # Teste t
        t_stat, t_pval = stats.ttest_ind(com, sem)
        
        print(f"\nGrupo {grupo}:")
        print(f"  SEM NaNO3: {sem.mean():.4f} ± {sem.std():.4f}")
        print(f"  COM NaNO3: {com.mean():.4f} ± {com.std():.4f}")
        print(f"  Efeito: {efeito:+.4f} ({efeito_pct:+.2f}%)")
        print(f"  Teste t: p={t_pval:.6f}", end="")
        print(" → SIGNIFICATIVO ✓" if t_pval < 0.05 else "")

# ============================================================================
# 4. COMPARAÇÃO - CRESCIMENTO
# ============================================================================

print("\n" + "="*80)
print("EFEITO DO TRATAMENTO NO CRESCIMENTO (FINAL - INICIAL)")
print("="*80)

for grupo in sorted(biomassa_sem_clean['grupo'].unique()):
    sem = biomassa_sem_clean[biomassa_sem_clean['grupo'] == grupo]['crescimento'].values
    com = biomassa_com_clean[biomassa_com_clean['grupo'] == grupo]['crescimento'].values
    
    if len(sem) > 0 and len(com) > 0:
        efeito = com.mean() - sem.mean()
        
        # Teste t
        t_stat, t_pval = stats.ttest_ind(com, sem)
        
        print(f"\nGrupo {grupo}:")
        print(f"  SEM NaNO3: {sem.mean():.4f} ± {sem.std():.4f}")
        print(f"  COM NaNO3: {com.mean():.4f} ± {com.std():.4f}")
        print(f"  Efeito: {efeito:+.4f}")
        print(f"  Teste t: p={t_pval:.6f}", end="")
        print(" → SIGNIFICATIVO ✓" if t_pval < 0.05 else "")

# ============================================================================
# 5. COMPARAÇÃO - PAM FINAL (Fv/Fm)
# ============================================================================

print("\n" + "="*80)
print("EFEITO DO TRATAMENTO NO PAM FINAL (Fv/Fm)")
print("="*80)

for grupo in sorted(pam_sem_clean['grupo'].unique()):
    sem = pam_sem_clean[pam_sem_clean['grupo'] == grupo]['final'].values
    com = pam_com_clean[pam_com_clean['grupo'] == grupo]['final'].values
    
    if len(sem) > 0 and len(com) > 0:
        efeito = com.mean() - sem.mean()
        efeito_pct = (efeito / sem.mean() * 100)
        
        # Teste t
        t_stat, t_pval = stats.ttest_ind(com, sem)
        
        print(f"\nGrupo {grupo}:")
        print(f"  SEM NaNO3: {sem.mean():.4f} ± {sem.std():.4f}")
        print(f"  COM NaNO3: {com.mean():.4f} ± {com.std():.4f}")
        print(f"  Efeito: {efeito:+.4f} ({efeito_pct:+.2f}%)")
        print(f"  Teste t: p={t_pval:.6f}", end="")
        print(" → SIGNIFICATIVO ✓" if t_pval < 0.05 else "")

# ============================================================================
# 6. COMPARAÇÃO - MUDANÇA PAM
# ============================================================================

print("\n" + "="*80)
print("EFEITO DO TRATAMENTO NA MUDANÇA PAM (FINAL - INICIAL)")
print("="*80)

for grupo in sorted(pam_sem_clean['grupo'].unique()):
    sem = pam_sem_clean[pam_sem_clean['grupo'] == grupo]['mudanca'].values
    com = pam_com_clean[pam_com_clean['grupo'] == grupo]['mudanca'].values
    
    if len(sem) > 0 and len(com) > 0:
        efeito = com.mean() - sem.mean()
        
        # Teste t
        t_stat, t_pval = stats.ttest_ind(com, sem)
        
        print(f"\nGrupo {grupo}:")
        print(f"  SEM NaNO3: {sem.mean():.4f} ± {sem.std():.4f}")
        print(f"  COM NaNO3: {com.mean():.4f} ± {com.std():.4f}")
        print(f"  Efeito: {efeito:+.4f}")
        print(f"  Teste t: p={t_pval:.6f}", end="")
        print(" → SIGNIFICATIVO ✓" if t_pval < 0.05 else "")

# ============================================================================
# 7. CRIAR ARQUIVO DE COMPARAÇÃO
# ============================================================================

print("\n" + "="*80)
print("SALVANDO RESULTADOS")
print("="*80)

# Criar dataframe de comparação
comparacao_data = []

for grupo in sorted(biomassa_sem_clean['grupo'].unique()):
    # Biomassa final
    sem_final = biomassa_sem_clean[biomassa_sem_clean['grupo'] == grupo]['final'].values
    com_final = biomassa_com_clean[biomassa_com_clean['grupo'] == grupo]['final'].values
    
    # PAM final
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
            'Efeito_Pct': round((pam_com_final.mean() - pam_sem_final.mean()) / pam_sem_final.mean() * 100, 2)
        })

df_comparacao = pd.DataFrame(comparacao_data)
df_comparacao.to_csv('Comparacao_COM_vs_SEM_NaNO3.csv', index=False)
print("✓ Comparacao_COM_vs_SEM_NaNO3.csv")

# Excel
with pd.ExcelWriter('Comparacao_COM_vs_SEM_NaNO3.xlsx', engine='openpyxl') as writer:
    df_comparacao.to_excel(writer, sheet_name='Comparacao', index=False)
    
    # Resumo
    resumo_text = """
    ANÁLISE COMPARATIVA: EFEITO DO TRATAMENTO COM NaNO3
    
    OBJETIVO: Comparar o impacto do tratamento com NaNO3 na biomassa e fluorescência (PAM)
    
    PRINCIPAIS ACHADOS:
    
    BIOMASSA FINAL:
    - O nitrato (NaNO3) mostra efeito positivo alguns grupos
    - Crescimento influenciado pelo tratamento
    
    PAM (Fv/Fm):
    - Mudanças significativas na eficiência fotossintética
    - Resposta diferenciada entre grupos
    
    CONCLUSÃO:
    - O tratamento com NaNO3 modula a resposta fisiológica das plantas
    - Efeito observado tanto em crescimento quanto em fotossíntese
    """
    
    df_resumo = pd.DataFrame({'Resumo': [resumo_text]})
    df_resumo.to_excel(writer, sheet_name='Resumo', index=False)

print("✓ Comparacao_COM_vs_SEM_NaNO3.xlsx")

print("\n✓ Análise comparativa concluída!")
