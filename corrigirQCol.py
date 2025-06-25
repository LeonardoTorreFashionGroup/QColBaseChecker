import os
import shutil
import xlwings as xw
from typing import Optional
from utils import gerar_corrigido_path, DIR_TEMP

SHEET_NAME = "QColeção"
LOG_START_ROW = 2  # começa após cabeçalhos na linha 1

def corrigir_qcol(
    caminho_original: str,
    caminho_log: str
) -> Optional[str]:
    if not os.path.isfile(caminho_original):
        print("Original não encontrado:", caminho_original); return None
    if not os.path.isfile(caminho_log):
        print("Log não encontrado:", caminho_log);       return None

    corrigido = gerar_corrigido_path(caminho_original)
    print("\n[5/5] A aplicar correções:")
    shutil.copy2(caminho_original, corrigido)

    app = xw.App(visible=False)
    app.display_alerts = False
    app.api.EnableEvents = False
    app.api.AskToUpdateLinks = False
    app.api.Calculation = -4135

    wb_c = app.books.open(corrigido)
    ws_c = wb_c.sheets[SHEET_NAME]
    wb_l = app.books.open(caminho_log, read_only=True)
    ws_l = wb_l.sheets[0]

    # ler registros
    last = ws_l.range(f"A{LOG_START_ROW}").end("down").row
    data = ws_l.range(f"A{LOG_START_ROW}:E{last}").value or []

    for idx, row in enumerate(data, start=1):
        cell_ref = row[0]    # coluna A
        formula    = row[4]  # coluna E
        print(f"     ⌛ A corrigir célula {idx} de {len(data)}...", end="\r")
        ws_c.range(cell_ref).api.FormulaLocal = formula
    print()

    wb_c.save(); wb_c.close()
    wb_l.close(); app.quit()

    print("Correções aplicadas. Ficheiro em: ", corrigido)
    return corrigido
