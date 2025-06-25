import os
import shutil
import xlwings as xw
from typing import Optional, Tuple
from utils import gerar_temp_path, DIR_TEMP

LINHA_HEADER = 9
SHEET_NAME = "QColeção"


def comparar_colunas_e_gerar_temporarios(
    caminho_original: str,
    caminho_referencia: str
) -> Optional[Tuple[str, str]]:
    nome1 = os.path.basename(caminho_original)
    nome2 = os.path.basename(caminho_referencia)

    # 1) criar cópias temporárias
    temp1 = gerar_temp_path(caminho_original)
    temp2 = gerar_temp_path(caminho_referencia)
    shutil.copy2(caminho_original, temp1)
    shutil.copy2(caminho_referencia, temp2)
    print(f"[2/4] Cópias temporárias criadas:\n   • {temp1}\n   • {temp2}")

    # 2) abrir no Excel em background
    app = xw.App(visible=False)
    app.display_alerts = False
    app.api.EnableEvents = False
    app.api.AskToUpdateLinks = False
    app.api.Calculation = -4135

    wb1 = app.books.open(temp1, update_links=False, read_only=True)
    wb2 = app.books.open(temp2, update_links=False, read_only=True)
    ws1 = wb1.sheets[SHEET_NAME]
    ws2 = wb2.sheets[SHEET_NAME]

    # desbloquear filtros e mostrar tudo
    for ws in (ws1, ws2):
        try:
            if ws.api.AutoFilterMode:
                ws.api.AutoFilterMode = False
        except:
            pass
        for row in ws.range("A1").expand("down").rows:
            row.api.EntireRow.Hidden = False
        for col in ws.range("A1").expand("right").columns:
            col.api.EntireColumn.Hidden = False

    # ler e strip
    c1 = [(c.value or "").strip()
          for c in ws1.range(f"A{LINHA_HEADER}").expand("right")]
    c2 = [(c.value or "").strip()
          for c in ws2.range(f"A{LINHA_HEADER}").expand("right")]

    wb1.close()
    wb2.close()
    app.quit()

    # detectar diferenças
    max_len = max(len(c1), len(c2))
    diffs, excl1, excl2 = [], [], []

    for i in range(max_len):
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
        print(f"[3/5] Diferenças encontradas nos cabeçalhos:")
        if diffs:
            print(f"   • Mudança de valor/posição: {len(diffs)} itens")
            for col, a, b in diffs:
                print(f"     - {col}{LINHA_HEADER}: '{a}' ≠ '{b}'")
        if excl1:
            print(f"   • Exclusivas em {nome1}: {len(excl1)}")
            for col, nm in excl1:
                print(f"     - {col}{LINHA_HEADER} = '{nm}'")
        if excl2:
            print(f"   • Exclusivas em {nome2}: {len(excl2)}")
            for col, nm in excl2:
                print(f"     - {col}{LINHA_HEADER} = '{nm}'")
        print("\nCorrija manualmente os cabeçalhos antes de continuar! ⚠️\n")
        return None

    print("\n[3/5] Cabeçalhos OK. Nenhuma diferença encontrada.✅")
    return temp1, temp2
