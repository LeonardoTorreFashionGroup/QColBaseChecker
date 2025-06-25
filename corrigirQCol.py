import os
import shutil
import xlwings as xw
from datetime import datetime
from typing import Optional
from utils import DIR_TEMP, _timestamp

SHEET_NAME = "QColeção"
LOG_HEADERS_ROW = 1  # cabeçalhos na linha 1 do log
LOG_DATA_START = 2   # dados começam na linha 2

def _gerar_nome_corrigido(orig_path: str) -> str:
    base, ext = os.path.splitext(os.path.basename(orig_path))
    ts = _timestamp()
    return os.path.join(DIR_TEMP, f"{base}_corrigido_{ts}{ext}")

def corrigir_qcol(
    caminho_original: str,
    caminho_log: str
) -> Optional[str]:
    """
    A partir do ficheiro original e do relatório de erros (Excel),
    cria uma cópia em DIR_TEMP com sufixo '_corrigido_<timestamp>',
    lê cada linha do log, extrai a 'Célula' e a 'Fórmula Esperada'
    e aplica essa fórmula no ficheiro copiado. Regressa o caminho
    do ficheiro corrigido ou None em caso de erro.
    """
    # 1) validar existência
    if not os.path.isfile(caminho_original):
        print("❌ Ficheiro original não encontrado:", caminho_original)
        return None
    if not os.path.isfile(caminho_log):
        print("❌ Relatório de erros não encontrado:", caminho_log)
        return None

    os.makedirs(DIR_TEMP, exist_ok=True)

    # 2) criar cópia para correção
    corrigido = _gerar_nome_corrigido(caminho_original)
    print("\n=== 3. APLICAR CORREÇÕES ===")
    print("1) Criando cópia para correção…")
    shutil.copy2(caminho_original, corrigido)
    print("   Cópia criada em:", corrigido)

    # 3) abrir ficheiro corrigido e log
    print("2) Abrindo Excel em modo silencioso…")
    app = xw.App(visible=False)
    app.display_alerts = False
    app.api.AskToUpdateLinks = False
    app.api.EnableEvents = False
    app.api.Calculation = -4135  # manual

    wb_corr = app.books.open(corrigido)
    ws_corr = wb_corr.sheets[SHEET_NAME]

    wb_log = app.books.open(caminho_log, update_links=False, read_only=True)
    ws_log = wb_log.sheets[0]

    # 4) ler todas as linhas do log
    print("3) Lendo log de erros…")
    # encontra últimas linhas preenchidas na coluna A
    last = ws_log.range(f"A{LOG_DATA_START}").end("down").row
    registros = ws_log.range(f"A{LOG_DATA_START}:F{last}").value or []

    # 5) aplicar correções
    print(f"4) Aplicando {len(registros)} correção(ões)…")
    for idx, row in enumerate(registros, start=1):
        celula = row[0]             # coluna A: Célula (ex. 'W10')
        formula_esperada = row[4]   # coluna E: Fórmula Esperada
        print(f"   [{idx}/{len(registros)}] {celula} ← {formula_esperada}", end="\r")
        # usar FormulaLocal para manter idioma
        ws_corr.range(celula).api.FormulaLocal = formula_esperada
    print()  # nova linha após progresso

    # 6) guardar e fechar
    print("5) Gravando e fechando ficheiro corrigido…")
    wb_corr.save()
    wb_corr.close()
    wb_log.close()
    app.quit()

    print("✅ Correção concluída. Ficheiro gerado em:", corrigido, "\n")
    return corrigido
