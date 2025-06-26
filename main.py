import sys
from pathlib import Path
from headers    import comparar_colunas_e_gerar_temporarios
from validarQCol import verificar_versus_referencia
from corrigirQCol import corrigir_qcol

USO_ORIG = r"H:\COLECÇÕES TORRE\PV 2026\TORRE UOMO\QCol. Base_TUPV26_V74_LFC.xlsm"
USO_REF  = r"H:\Processos Gerais\PV20- Gestão do produto\Quadros da coleção base\QCol. Base_V39 LFC.xlsm"

#USO_ORIG = r"H:\COLECÇÕES TORRE\PV 2026\TORRE UOMO\QCol. Base_TUPV26_TESTE_LFC.xlsm" #TEST
#USO_REF  = r"H:\Processos Gerais\PV20- Gestão do produto\Quadros da coleção base\QCol. Base_V39 LFC.xlsm" #TEST


def main():
    # argumentos ou default
    if len(sys.argv) == 3:
        orig, ref = sys.argv[1], sys.argv[2]
    else:
        print("[1/5] Ficheiros originais em uso:")
        print("   • Quadro Coleção :", USO_ORIG)
        print("   • Quadro Base    :", USO_REF, "\n")
        orig, ref = USO_ORIG, USO_REF

    # comparar cabeçalhos
    temps = comparar_colunas_e_gerar_temporarios(orig, ref)
    if not temps:
        sys.exit(1)
    temp_orig, temp_ref = temps


    resp = input("\nCabeçalhos OK. Avançar p/ validação de fórmulas? (S/N): ").strip().lower()
    if resp != 's':
        print("Operação cancelada pelo utilizador."); sys.exit(0)

    # validar fórmulas
    log_path = verificar_versus_referencia(orig, ref)
    if not log_path:
        print("\nTudo OK, sem erros de fórmula.")
        return

    resp2 = input("\nDeseja aplicar correções automáticas? (S/N): ").strip().lower()
    if resp2 == 's':
        result = corrigir_qcol(orig, log_path)
        if not result:
            sys.exit(1)
    else:
        print("Correções automáticas não aplicadas.")

if __name__ == "__main__":
    main()
