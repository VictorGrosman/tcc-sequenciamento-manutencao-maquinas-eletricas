"""Heurísticas SJF tradicional e SJF com prioridade dinâmica por folga."""
from pathlib import Path
from dados import DadosEntrada
from sequenciamento import ETAPAS, FLUXO, IGNORAR, gerar_horario_turno, executar_sequenciamento


def executar_sjf(dados: DadosEntrada):
    """SJF tradicional: prioriza a menor carga horária total do projeto."""
    wb, ws, projetos, ops, makespan, _ = executar_sequenciamento(dados, 'SJF', False)
    return wb, ws, projetos, ops, makespan


def executar_sjf_folga(dados: DadosEntrada):
    """SJF modificado: menor folga do projeto; desempate pela carga total."""
    wb, ws, projetos, ops, makespan, _ = executar_sequenciamento(dados, 'SJF', True)
    return wb, ws, projetos, ops, makespan


def salvar_resultado(workbook, caminho_saida):
    caminho_saida = Path(caminho_saida)
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(caminho_saida)
    return caminho_saida
