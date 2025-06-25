import os
import shutil
import xlwings as xw
from typing import Optional, Tuple
from utils import gerar_temp, _timestamp, DIR_TEMP

LINHA_HEADER = 9
SHEET_NAME = "QColeção"

def comparar_colunas_e_gerar_temporarios(
    caminho_1_orig: str,
    caminho_2_orig: str
) -> Optional[Tuple[str, str]]:
    nome1 = os.path.basename(caminho_1_orig)
    nome2 = os.path.basename(caminho_2_orig)

    # Copiar para temp
    temp1 = gerar_temp(caminho_1_orig)
    temp2 = gerar_temp(caminho_2_orig)
    shutil.copy2(caminho_1_orig, temp1)
    shutil.copy2(caminho_2_orig, temp2)
    print("Cópias temporárias criadas:")
    print(" ", temp1)
    print(" ", temp2)

    # abrir Excel em background
    app = xw.App(visible=False)
    app.display_alerts = False
    app.api.AskToUpdateLinks = False
    app.api.EnableEvents = False
    app.api.Calculation = -4135  # manual

    wb1 = app.books.open(temp1, update_links=False, read_only=True)
    wb2 = app.books.open(temp2, update_links=False, read_only=True)
    ws1 = wb1.sheets[SHEET_NAME]
    ws2 = wb2.sheets[SHEET_NAME]

    # desbloquear filtros
    for ws in (ws1, ws2):
        try:
            if ws.api.AutoFilterMode:
                ws.api.AutoFilterMode = False
        except:
            pass
        for r in ws.range("A1").expand("down").rows:
            r.api.EntireRow.Hidden = False
        for c in ws.range("A1").expand("right").columns:
            c.api.EntireColumn.Hidden = False

    # ler cabeçalhos
    c1 = [(c.value or "").strip() for c in ws1.range(f"A{LINHA_HEADER}").expand("right")]
    c2 = [(c.value or "").strip() for c in ws2.range(f"A{LINHA_HEADER}").expand("right")]

    wb1.close()
    wb2.close()
    app.quit()

    # detectar diferenças
    max_cols = max(len(c1), len(c2))
    diffs = []
    excl1 = []
    excl2 = []

    for i in range(max_cols):
        a = c1[i] if i < len(c1) else ""
        b = c2[i] if i < len(c2) else ""
        if a != b:
            col = xw.utils.col_name(i+1)
            diffs.append((col, a, b))

    for i, nm in enumerate(c1):
        if nm and nm not in c2:
            excl1.append((xw.utils.col_name(i+1), nm))
    for j, nm in enumerate(c2):
        if nm and nm not in c1:
            excl2.append((xw.utils.col_name(j+1), nm))

    # report
    if diffs or excl1 or excl2:
        print("\nForam detectadas diferenças nos cabeçalhos:")
        if diffs:
            print(f" - Diferenças de posição/nome ({len(diffs)}):")
            for col, a, b in diffs:
                print(f"   {col}{LINHA_HEADER} | {nome1}:'{a}' ≠ {nome2}:'{b}'")
        if excl1:
            print(f"\n - Exclusivas em {nome1} ({len(excl1)}):")
            for col, nm in excl1:
                print(f"   {col}{LINHA_HEADER} = '{nm}'")
        if excl2:
            print(f"\n - Exclusivas em {nome2} ({len(excl2)}):")
            for col, nm in excl2:
                print(f"   {col}{LINHA_HEADER} = '{nm}'")
        print("\n⛔  Corrija os cabeçalhos antes de prosseguir.")
        return None

    # sem diferencas
    return temp1, temp2
