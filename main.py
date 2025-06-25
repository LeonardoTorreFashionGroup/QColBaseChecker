import sys
from pathlib import Path

from headers import comparar_colunas_e_gerar_temporarios
from validarQCol import verificar_versus_referencia
from corrigirQCol import corrigir_qcol

USO_ORIG = r"H:\COLECÇÕES TORRE\PV 2026\TORRE UOMO\QCol. Base_TUPV26_V73_JF.xlsm"
USO_REF = r"H:\Processos Gerais\PV20- Gestão do produto\Quadros da coleção base\QCol. Base_V39 LFC.xlsm"

USO_ORIG = r"H:\COLECÇÕES TORRE\PV 2026\TORRE UOMO\QCol. Base_TUPV26_TESTE_LFC.xlsm"  # TEST
USO_REF = r"H:\Processos Gerais\PV20- Gestão do produto\Quadros da coleção base\QCol. Base_V39 LFC.xlsm"  # TEST


USO_ORIG = r"H:\Informatica\LEONARDOCRUZ\Projetos\QColChecker\temp_data\QCol. Base_TUPV26_TESTE_LFC_corrigido_2025-06-25-12_14_37.xlsm"


def main():
    print("=== INÍCIO ===")
    print(" ORIGINAL   :", USO_ORIG)
    print(" REFERÊNCIA :", USO_REF, "\n")
    orig, ref = USO_ORIG, USO_REF

    # validar headers
    print("=== 1. VALIDAR HEADERS ===")
    res = comparar_colunas_e_gerar_temporarios(orig, ref)
    if not res:
        sys.exit(1)
    temp_orig, temp_ref = res

    # perguntar antes de prosseguir
    resp = input(
        "\nOs cabeçalhos conferem. Deseja avançar para as validações de fórmula? (S/N): ").strip().lower()
    if resp != 's':
        print("\nOperação cancelada pelo utilizador.")
        sys.exit(0)

    # validar fórmulas
    print("\n=== 2. VALIDAÇÕES DO QUADRO DE COLEÇÃO ===")
    rel = verificar_versus_referencia(orig, ref)
    if rel:
        print("\nERROS ENCONTRADOS! VEJA O RELATÓRIO EM:\n   ", rel)
        # sys.exit(1)

    resp = input(
        "Deseja aplicar correções automáticas ao ficheiro? (S/N): ").strip().lower()
    if resp == 's':
        path_corr = corrigir_qcol(USO_ORIG, rel)
        if not path_corr:
            sys.exit(1)
        print("Ficheiro corrigido disponível em:", path_corr)
    else:
        print("Correções automáticas canceladas pelo utilizador.")


if __name__ == "__main__":
    main()
