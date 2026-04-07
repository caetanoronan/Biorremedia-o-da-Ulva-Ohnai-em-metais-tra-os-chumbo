import pandas as pd

# Abrir o arquivo ODS
ods_file = "DADOS DO EXPERIMENTO BIORREMEDIAÇÃO - DORLAM'S.ods"

# Ler todas as abas disponíveis
excel_file = pd.ExcelFile(ods_file, engine='odf')
sheet_names = excel_file.sheet_names

print("Abas disponíveis no arquivo ODS:")
for i, name in enumerate(sheet_names, 1):
    print(f"{i}. {name}")

# Explorar as 4 abas de PAM
print("\n" + "="*60)
print("EXPLORANDO ABAS DE PAM")
print("="*60)

pam_sheets = [name for name in sheet_names if 'PAM' in name.upper() or 'Fv/Fm' in name]
print(f"\nAbas PAM encontradas: {pam_sheets}\n")

for sheet_name in pam_sheets:
    try:
        df = pd.read_excel(ods_file, sheet_name=sheet_name, engine='odf')
        print(f"\n{'─'*60}")
        print(f"ABA: {sheet_name}")
        print(f"{'─'*60}")
        print(f"Dimensões: {df.shape}")
        print(f"\nPrimeiras linhas:")
        print(df.head(10))
        print(f"\nColunas: {list(df.columns)}")
    except Exception as e:
        print(f"\nErro ao ler {sheet_name}: {e}")
