from dataclasses import dataclass
from pathlib import Path
import io
import pandas as pd


@dataclass
class DadosEntrada:
    arquivo: Path
    bytes_planilha: bytes
    colaboradores: pd.DataFrame
    projetos: pd.DataFrame
    operacoes: list[str]


def carregar_dados(caminho: str | Path) -> DadosEntrada:
    caminho = Path(caminho)
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
    conteudo = caminho.read_bytes()
    xls = pd.ExcelFile(io.BytesIO(conteudo))
    obrigatorias = {'Colaboradores', 'Projetos', 'Cronograma'}
    faltantes = obrigatorias - set(xls.sheet_names)
    if faltantes:
        raise ValueError(f"Abas obrigatórias ausentes: {', '.join(sorted(faltantes))}")
    colaboradores = pd.read_excel(xls, sheet_name='Colaboradores', header=None)
    projetos = pd.read_excel(xls, sheet_name='Projetos', skiprows=1)
    if len(projetos.columns) < 5:
        raise ValueError("A aba 'Projetos' não contém as colunas esperadas.")
    colunas_data = {str(c).strip().upper() for c in projetos.columns[-2:]}
    if not {'INÍCIO', 'PRAZO'}.issubset(colunas_data):
        raise ValueError("As duas últimas colunas da aba 'Projetos' devem ser INÍCIO e PRAZO.")
    operacoes = projetos.columns[2:-2].tolist()
    return DadosEntrada(caminho, conteudo, colaboradores, projetos, operacoes)
