import xlwings as xw
import os
import shutil
from datetime import datetime

# === CONFIGURAÇÕES ===
CAMINHO_ORIGINAL = r"H:\\COLECÇÕES TORRE\\PV 2026\\TORRE UOMO\\QCol. Base_TUPV26_V73_JF.xlsm"
DIR_TEMP = r"H:\\Informatica\\LEONARDOCRUZ\\Suporte\\QColBaseChecker\\temp_data"
os.makedirs(DIR_TEMP, exist_ok=True)

NOME_SHEET = "QColeção"
COL_INICIO = 1   # Coluna A
COL_FIM = 219    # Coluna HJ
LINHA_INICIO = 10

# === HELPERS ===


def timestamp():
    return datetime.now().strftime("%Y-%m-%d-%H_%M_%S")


def gerar_nome_temp(path):
    nome = os.path.basename(path)
    raiz, ext = os.path.splitext(nome)
    return os.path.join(DIR_TEMP, f"{raiz}_temp_{timestamp()}{ext}")


# === INÍCIO ===
print("\n=== INÍCIO DO PROCESSO DE VALIDAÇÃO DE FÓRMULAS EXTERNAS ===")

# Copiar ficheiro original para temporário
print("A copiar ficheiro atual...")
caminho_temp = gerar_nome_temp(CAMINHO_ORIGINAL)
shutil.copy2(CAMINHO_ORIGINAL, caminho_temp)
print("Cópia criada em:", caminho_temp)

# Abrir o ficheiro no Excel
print("A abrir o ficheiro...")
app = xw.App(visible=False)
wb = app.books.open(caminho_temp, update_links=False, read_only=False)
ws = wb.sheets[NOME_SHEET]

# === LIMPEZA DA FOLHA ===
print("A remover filtros e a mostrar todas as linhas e colunas...")

try:
    if ws.api.AutoFilterMode:
        ws.api.AutoFilterMode = False
except:
    pass

for row in ws.range("A1").expand("down").rows:
    row.api.EntireRow.Hidden = False

for col in ws.range("A1").expand("right").columns:
    col.api.EntireColumn.Hidden = False

# === VALIDAÇÃO ===
ultima_linha = ws.range("A" + str(ws.cells.last_cell.row)).end("up").row
LINHA_FIM = max(ultima_linha, LINHA_INICIO + 1)
diferencas = []

print(f"\nA validar da linha {LINHA_INICIO} até {LINHA_FIM}...\n")

for linha in range(LINHA_INICIO, LINHA_FIM + 1):
    print(f"  ➤ A validar linha {linha} de {LINHA_FIM}...")
    if linha == LINHA_FIM:
        print(f"\n  ➤ A construir relatório... aguarde...\n")
    for col in range(COL_INICIO, COL_FIM + 1):
        celula = ws.cells(linha, col)
        celula_nome = f"{xw.utils.col_name(col)}{linha}"

        try:
            f_atual_raw = celula.api.Formula
            f_atual = f_atual_raw.lstrip("=") if f_atual_raw else None
        except Exception as e:
            print(f"[Erro em {celula_nome}]: {e}")
            continue

        if not f_atual:
            continue

        f_atual_lower = f_atual.lower()

        if any(ext in f_atual_lower for ext in [".xls", ".xlsx", ".xlsm"]):
            if "#ref!" in f_atual_lower:
                tipo = "Erro externo - #REF!"
            elif "#n/a" in f_atual_lower:
                tipo = "Erro externo - #N/A"
            else:
                tipo = "Fórmula com referência externa"

            diferencas.append(
                (celula_nome, linha, xw.utils.col_name(
                    col), tipo, "", f_atual.strip())
            )

# Fechar ficheiro original
wb.close()
app.quit()

# === RELATÓRIO ===
log_excel = os.path.join(DIR_TEMP, f"QColBase_Erros_{timestamp()}.xlsx")
app_out = xw.App(visible=False)
wb_out = app_out.books.add()
ws_out = wb_out.sheets[0]

# Cabeçalho
ws_out.range("A1").value = ["Célula", "Linha", "Coluna",
                            "Classificação", "Fórmula Esperada", "Atual"]

# Inserir linhas no Excel
for i, linha in enumerate(diferencas, start=2):
    ws_out.range(f"A{i}").value = linha

# Guardar relatório
wb_out.save(log_excel)
wb_out.close()
app_out.quit()

print("\n✅ Validação concluída com sucesso.")
print("Relatório salvo em:", log_excel)
