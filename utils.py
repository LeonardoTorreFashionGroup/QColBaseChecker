import os
import shutil
from datetime import datetime

# Pasta para todos os temp
DIR_TEMP = r"H:\Informatica\LEONARDOCRUZ\Projetos\QColChecker\temp_data"
os.makedirs(DIR_TEMP, exist_ok=True)


def _timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d-%H_%M_%S")


def gerar_temp(orig_path: str) -> str:
    """
    Gera caminho em DIR_TEMP com o nome original + '_temp_' + timestamp.
    Exemplo: 'QCol. Base.xlsm' - '.../QCol. Base_temp_2025-06-25-11_00_00.xlsm'
    """
    base, ext = os.path.splitext(os.path.basename(orig_path))
    return os.path.join(DIR_TEMP, f"{base}_temp_{_timestamp()}{ext}")
