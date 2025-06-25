import os
import shutil
import xlwings as xw
from typing import Optional
from utils import gerar_temp_path, DIR_TEMP, timestamp

SHEET_NAME = "QColeção"
COL_INICIO = 1
COL_FIM = 219
LINHA_INICIO = 10


def verificar_versus_referencia(
    caminho_original: str,
    caminho_referencia: str
) -> Optional[str]:
    print("\n[3/4] Iniciando verificação de fórmulas…")

    # cópias temporárias
    tmp1 = gerar_temp_path(caminho_original)
    tmp2 = gerar_temp_path(caminho_referencia)
    shutil.copy2(caminho_original, tmp1)
    shutil.copy2(caminho_referencia, tmp2)
    print(f"   • {os.path.basename(tmp1)}\n   • {os.path.basename(tmp2)}")

    # abrir
    app = xw.App(visible=False)
    app.display_alerts = False
    app.api.EnableEvents = False
    app.api.AskToUpdateLinks = False
    app.api.Calculation = -4135

    wb1 = app.books.open(tmp1, update_links=False, read_only=False)
    wb2 = app.books.open(tmp2, update_links=False, read_only=True)
    ws1 = wb1.sheets[SHEET_NAME]
    ws2 = wb2.sheets[SHEET_NAME]

    # remove filtros
    try:
        if ws1.api.AutoFilterMode:
            ws1.api.AutoFilterMode = False
    except:
        pass
    for r in ws1.range("A1").expand("down").rows:
        r.api.EntireRow.Hidden = False
    for c in ws1.range("A1").expand("right").columns:
        c.api.EntireColumn.Hidden = False

    # determinar última linha
    ultima = ws1.range(f"A{ws1.cells.last_cell.row}").end("up").row
    fim = max(ultima, LINHA_INICIO + 1)
    print(f"   Validando linhas {LINHA_INICIO} → {fim}")

    erros = []
    for linha in range(LINHA_INICIO, fim + 1):
        for offset in range(1, COL_FIM - COL_INICIO + 2):
            c1 = ws1.api.Cells(linha, COL_INICIO + offset - 1)
            c2 = ws2.api.Cells(linha, COL_INICIO + offset - 1)
            addr = f"{xw.utils.col_name(COL_INICIO+offset-1)}{linha}"
            txt = c1.Text or ""
            f1 = c1.FormulaLocal or ""
            f2 = c2.FormulaLocal or ""
            cond_ref = "#REF!" in txt
            cond_na = "#N/A" in txt
            cond_ext = any(ext in f1.lower()
                           for ext in (".xls", ".xlsx", ".xlsm"))
            if cond_ref or cond_na or cond_ext:
                tipo = "Erro #REF!" if cond_ref else "Erro #N/A" if cond_na else "Erro externa"
                atual = (f1 or txt).strip()
                esperado = f2.strip()
                erros.append(
                    (addr, linha, addr[:-len(str(linha))], tipo, esperado, atual))
        print(f"   Linha {linha}/{fim} processada", end="\r")
    print()

    wb1.close()
    wb2.close()
    app.quit()

    if not erros:
        print("Sem diferenças de fórmula encontradas.")
        return None

    # gerar relatório
    log = os.path.join(DIR_TEMP, f"QColBase_Erros_{timestamp()}.xlsx")
    out = xw.App(visible=False)
    wb_out = out.books.add()
    ws_out = wb_out.sheets[0]
    ws_out.range("A1").value = ["Celula", "Linha", "Coluna",
                                "Tipo", "Formula_Esperado", "Formula_Atual"]
    for i, row in enumerate(erros, start=2):
        r = list(row)
        if r[5].startswith("="):
            r[5] = "'" + r[5]
        if r[4].startswith("="):
            r[4] = "'" + r[4]
        ws_out.range(f"A{i}").value = r
    wb_out.save(log)
    wb_out.close()
    out.quit()
    print(f"⚠️ Erros detectados. Relatório: {log}")
    return log