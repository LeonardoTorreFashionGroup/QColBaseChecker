import os
import shutil
import xlwings as xw
from typing import Optional
from utils import gerar_temp, _timestamp, DIR_TEMP

SHEET_NAME = "QColeção"
COL_INICIO = 1
COL_FIM = 219
LINHA_INICIO = 10


def verificar_versus_referencia(
    caminho_original: str,
    caminho_referencia: str
) -> Optional[str]:
    print("\n=== INÍCIO DA VERIFICAÇÃO COMPLETA ===")

    # copiar para temp
    print("A copiar ficheiro original...")
    tmp1 = gerar_temp(caminho_original)
    shutil.copy2(caminho_original, tmp1)
    print("  Cópia criada:", tmp1)

    print("A copiar ficheiro de referência...")
    tmp2 = gerar_temp(caminho_referencia)
    shutil.copy2(caminho_referencia, tmp2)
    print("  Cópia criada:", tmp2)

    # abrir em silêncio
    print("A abrir Excel em modo silencioso...")
    app = xw.App(visible=False)
    app.display_alerts = False
    app.api.AskToUpdateLinks = False
    app.api.EnableEvents = False
    app.api.Calculation = -4135

    wb1 = app.books.open(tmp1, update_links=False, read_only=False)
    wb2 = app.books.open(tmp2, update_links=False, read_only=True)
    ws1 = wb1.sheets[SHEET_NAME]
    ws2 = wb2.sheets[SHEET_NAME]

    # remover filtros/ocultações
    try:
        if ws1.api.AutoFilterMode:
            ws1.api.AutoFilterMode = False
    except:
        pass
    for r in ws1.range("A1").expand("down").rows:
        r.api.EntireRow.Hidden = False
    for c in ws1.range("A1").expand("right").columns:
        c.api.EntireColumn.Hidden = False

    # determinar linhas
    ultima = ws1.range(f"A{ws1.cells.last_cell.row}").end("up").row
    fim = max(ultima, LINHA_INICIO + 1)
    print(f"\nA validar da linha {LINHA_INICIO} até {fim}...")

    erros = []
    for linha in range(LINHA_INICIO, fim + 1):
        print(f"  Validando linha {linha}/{fim}", end="\r")
        rng1 = ws1.range((linha, COL_INICIO), (linha, COL_FIM)).api
        rng2 = ws2.range((linha, COL_INICIO), (linha, COL_FIM)).api
        for off in range(1, COL_FIM - COL_INICIO + 2):
            c1 = rng1.Cells(1, off)
            c2 = rng2.Cells(1, off)
            addr = f"{xw.utils.col_name(COL_INICIO + off - 1)}{linha}"
            txt = c1.Text or ""
            f1 = c1.FormulaLocal or ""
            f2 = c2.FormulaLocal or ""
            if (
                "#REF!" in txt
                or "#N/A" in txt
                or any(ext in f1.lower() for ext in (".xls", ".xlsx", ".xlsm"))
            ):
                tipo = (
                    "Erro #REF!" if "#REF!" in txt
                    else "Erro #N/A" if "#N/A" in txt
                    else "Erro dependência externa"
                )
                atual = (f1 or txt).strip()
                esperado = f2.strip()
                erros.append(
                    (addr, linha, addr[:-len(str(linha))], tipo, esperado, atual))
    print()  # quebra progresso

    wb1.close()
    wb2.close()
    app.quit()

    if not erros:
        print("Nenhuma diferença de fórmula encontrada.\n")
        return None

    # construir relatório
    print("\nA construir relatório...")
    log = os.path.join(DIR_TEMP, f"QColBase_Erros_{_timestamp()}.xlsx")
    app_out = xw.App(visible=False)
    wb_out = app_out.books.add()
    ws_out = wb_out.sheets[0]
    ws_out.range("A1").value = [
        "Célula", "Linha", "Coluna", "Classificação", "Fórmula Esperada", "Fórmula Atual"
    ]
    for idx, row in enumerate(erros, start=2):
        r = list(row)
        if r[5].startswith("="):
            r[5] = "'" + r[5]
        if r[4].startswith("="):
            r[4] = "'" + r[4]
        ws_out.range(f"A{idx}").value = r

    wb_out.save(log)
    wb_out.close()
    app_out.quit()

    print("Relatório guardado em:", log, "\n")
    return log
