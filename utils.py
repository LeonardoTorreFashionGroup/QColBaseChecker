import os
from datetime import datetime

DIR_TEMP = r"H:\Informatica\LEONARDOCRUZ\Projetos\QColChecker\temp_data"
os.makedirs(DIR_TEMP, exist_ok=True)

def timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d_%H%M%S")

def gerar_temp_path(orig_path: str) -> str:
    """
    Gera um nome em DIR_TEMP com sufixo '_temp_<timestamp>'.
    """
    base, ext = os.path.splitext(os.path.basename(orig_path))
    return os.path.join(DIR_TEMP, f"{base}_temp_{timestamp()}{ext}")

def gerar_corrigido_path(orig_path: str) -> str:
    """
    Gera um nome em DIR_TEMP com sufixo '_corrigido_<timestamp>'.
    """
    base, ext = os.path.splitext(os.path.basename(orig_path))
    return os.path.join(DIR_TEMP, f"{base}_corrigido_{timestamp()}{ext}")
