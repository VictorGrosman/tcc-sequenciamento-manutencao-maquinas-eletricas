"""Heurísticas SPT tradicional e SPT com prioridade dinâmica por folga."""
from dados import DadosEntrada
from sequenciamento import ETAPAS, FLUXO, IGNORAR, gerar_horario_turno, executar_sequenciamento

CRITERIO = 'hh_total'


def executar_spt(dados: DadosEntrada):
    """SPT tradicional: menor tempo de processamento da operação elegível."""
    return executar_sequenciamento(dados, 'SPT', False)


def executar_spt_folga(dados: DadosEntrada):
    """SPT modificado: menor folga do projeto; desempate pelo tempo da operação."""
    return executar_sequenciamento(dados, 'SPT', True)
