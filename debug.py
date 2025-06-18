import xlwings as xw
import os
import shutil
from datetime import datetime

# === CONFIGURAÇÕES ===
CAMINHO_ORIGINAL = r"H:\\COLECÇÕES TORRE\\PV 2026\\TORRE UOMO\\QCol. Base_TUPV26_DF_V72.xlsm"
DIR_TEMP = r"H:\\Informatica\\LEONARDOCRUZ\\Suporte\\QColBaseChecker\\temp_data"
os.makedirs(DIR_TEMP, exist_ok=True)

NOME_SHEET = "QColeção"
CELULA_DEBUG = "CK10"

def timestamp():
    return datetime.now().strftime("%Y-%m-%d-%H_%M_%S")


def gerar_nome_temp(path):
    nome = os.path.basename(path)
    raiz, ext = os.path.splitext(nome)
    return os.path.join(DIR_TEMP, f"{raiz}_temp_{timestamp()}{ext}")


# === INÍCIO ===
print("\n=== DEBUG: VERIFICAÇÃO DA CÉLULA CK10 COM FÓRMULA ===")

# Copiar ficheiro original para temporário
print("A copiar ficheiro original para temporário...")
caminho_temp = gerar_nome_temp(CAMINHO_ORIGINAL)
shutil.copy2(CAMINHO_ORIGINAL, caminho_temp)
print("Cópia criada em:", caminho_temp)

# Abrir o ficheiro temporário
print("A abrir o ficheiro temporário...")
app = xw.App(visible=False)
wb = app.books.open(caminho_temp, update_links=False, read_only=True)
ws = wb.sheets[NOME_SHEET]
celula = ws.range(CELULA_DEBUG)

# === TENTAR LER FORMULAS DE CK10 ===
print(f"\n🔎 Aceder à célula {CELULA_DEBUG} na folha '{NOME_SHEET}'...")

try:
    formula_api = celula.api.Formula
    print(f"[API] .Formula       → {formula_api}")
except Exception as e:
    print(f"[API] .Formula       → ERRO: {e}")

try:
    formula_local = celula.api.FormulaLocal
    print(f"[API] .FormulaLocal  → {formula_local}")
except Exception as e:
    print(f"[API] .FormulaLocal  → ERRO: {e}")

try:
    formula_xlwings = celula.formula
    print(f"[xlwings] .formula   → {formula_xlwings}")
except Exception as e:
    print(f"[xlwings] .formula   → ERRO: {e}")

# Fechar tudo
wb.close()
app.quit()

print("\n✅ Debug concluído.")
