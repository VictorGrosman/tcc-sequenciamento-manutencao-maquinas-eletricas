"""Busca local por inserção de operações completas após SJF/SPT tradicional.

Uma inserção altera a ordem de prioridade (não desloca células diretamente).
Cada candidato é reprogramado integralmente pelo núcleo de sequenciamento,
que continua responsável por precedências, competências e capacidade.
"""
from __future__ import annotations

from dataclasses import dataclass
from collections import defaultdict

from gantt import _turno
from sequenciamento import (JORNADA_MAXIMA, _mapear_operacoes,
                            _nome_atividade, dias_atraso_uteis,
                            executar_sequenciamento, norm, ler_alocacoes_celula, ler_bloqueios_resultado)


@dataclass
class ResultadoBusca:
    resultado: tuple
    atraso_inicial: int
    atraso_final: int
    movimentos_aceitos: int
    candidatos_avaliados: int


def _objetivo(projetos):
    if any(p['data_conclusao'] is None for p in projetos):
        return None
    return sum(dias_atraso_uteis(p['prazo'], p['data_conclusao']) for p in projetos)


def validar_cronograma(resultado, dados, tolerancia=1e-6):
    """Auditoria independente do cronograma materializado, inclusive precedências.

    Não valida feriados: o algoritmo de origem trabalha com segunda a sexta.
    """
    wb, ws, projetos, ops, _, _ = resultado
    preds, ignoradas = _mapear_operacoes(ops)
    raw = dados.colaboradores
    limites = {}
    for i in range(len(ops)):
        try:
            valor = raw.iloc[1, i + 2]
            import pandas as pd
            limites[i] = max(1, int(valor)) if not pd.isna(valor) else 1
        except (IndexError, TypeError, ValueError):
            limites[i] = 1
    habilidades = {}
    for r in range(2, len(raw)):
        nome = str(raw.iloc[r, 1]).strip()
        if nome and nome.lower() != 'nan':
            habilidades[norm(nome)] = {i for i in range(len(ops))
                                       if i + 2 < raw.shape[1] and raw.iloc[r, i + 2] == 1}
    atividades = {}
    for p in projetos:
        for tid in range(len(ops)):
            chave = _nome_atividade(p['nome'], ops[tid])
            if chave in atividades:
                raise ValueError(f'Identificador de atividade duplicado: {chave}')
            atividades[chave] = (p, tid)
    horas = defaultdict(float)
    inicio = {}
    fim = {}
    alocados = defaultdict(set)
    bloqueios = ler_bloqueios_resultado(wb)
    for row in range(2, ws.max_row + 1, 2):
        nome = ws.cell(row, 1).value
        if not nome:
            continue
        nome_norm = norm(str(nome).strip())
        skill = next((n for n in habilidades if nome_norm == n or nome_norm in n or n in nome_norm), None)
        for col in range(2, ws.max_column + 1):
            if (row, col) in bloqueios:
                continue  # registro original preservado, não é alocação deste cenário
            atividade_celula = ws.cell(row + 1, col).value
            if not atividade_celula:
                continue
            registros = ler_alocacoes_celula(ws.cell(row, col).value, atividade_celula)
            soma_dia = 0.0
            intervalos = []
            for turno, atividade in registros:
                if atividade not in atividades:
                    raise ValueError(f'Atividade desconhecida: {atividade}')
                p, tid = atividades[atividade]
                data = ws.cell(1, col).value
                if hasattr(data, 'date'):
                    data = data.date()
                if data is None or data.weekday() >= 5 or data < p['inicio']:
                    raise ValueError(f'Dia inválido para {atividade}: {data}')
                if skill is None or tid not in habilidades[skill]:
                    raise ValueError(f'Colaborador sem competência: {nome}, {atividade}')
                h, _ = _turno(turno)
                if h <= tolerancia:
                    raise ValueError(f'Turno inválido: {nome}, {data}, {turno}')
                soma_dia += h
                import re
                for hi, mi, hf, mf in re.findall(r'(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})', turno):
                    ini, fim_turno = int(hi) * 60 + int(mi), int(hf) * 60 + int(mf)
                    if fim_turno <= ini or not (450 <= ini < 720 and fim_turno <= 720 or
                                                 780 <= ini < 1050 and fim_turno <= 1050):
                        raise ValueError(f'Horário inválido: {nome}, {data}, {turno}')
                    intervalos.append((ini, fim_turno))
                chave = (p['id'], tid)
                horas[chave] += h
                inicio[chave] = min(inicio.get(chave, col), col)
                fim[chave] = max(fim.get(chave, col), col)
                alocados[(p['id'], tid, col)].add(row)
                if len(alocados[(p['id'], tid, col)]) > limites[tid]:
                    raise ValueError(f'Limite de colaboradores excedido: {atividade}, {data}')
            if soma_dia > JORNADA_MAXIMA + tolerancia:
                raise ValueError(f'Jornada excedida: {nome}, {data}, {soma_dia}')
            intervalos.sort()
            if any(intervalos[i][1] > intervalos[i + 1][0] for i in range(len(intervalos) - 1)):
                raise ValueError(f'Turnos sobrepostos: {nome}, {data}')
    for p in projetos:
        for tid, tarefa in p['tarefas'].items():
            chave = (p['id'], tid)
            esperado = tarefa['hh_total']
            if abs(horas[chave] - esperado) > tolerancia:
                raise ValueError(f'Carga divergente: {p["nome"]}, {ops[tid]}: {horas[chave]} != {esperado}')
            if tid in ignoradas:
                continue
            if esperado > tolerancia:
                for pred in preds[tid]:
                    if p['tarefas'][pred]['hh_total'] > tolerancia and not fim.get((p['id'], pred), -1) < inicio[chave]:
                        raise ValueError(f'Precedência violada: {p["nome"]}: {ops[pred]} -> {ops[tid]}')
    return True


def executar_busca_local(dados, metodo='SJF', max_iteracoes=5, max_candidatos=150, resultado_inicial=None, incremental=False):
    """Primeira melhora: tenta antecipar uma operação de projeto atrasado.

    Unidade de movimento: operação completa, representada por (projeto, operação).
    A execução pode ocupar vários dias e colaboradores, como na heurística original.
    """
    metodo = metodo.upper()
    if metodo not in {'SJF', 'SPT'}:
        raise ValueError('Método deve ser SJF ou SPT.')
    if max_iteracoes < 0 or max_candidatos < 0:
        raise ValueError('Os limites da busca devem ser não negativos.')
    atual = resultado_inicial if resultado_inicial is not None else executar_sequenciamento(dados, metodo=metodo, incremental=incremental)
    validar_cronograma(atual, dados)
    z = _objetivo(atual[2])
    if z is None:
        raise ValueError('O cronograma inicial não concluiu todos os projetos no horizonte.')
    inicial = z
    # A ordem é a das primeiras alocações observadas no cronograma original.
    ordem = [(p['id'], tid) for p, tid in atual[5]]
    aceitos = avaliados = 0
    for _ in range(max_iteracoes):
        atrasados = {p['id'] for p in atual[2]
                     if p['data_conclusao'] is not None and
                     dias_atraso_uteis(p['prazo'], p['data_conclusao']) > 0}
        if not atrasados or not z or avaliados >= max_candidatos:
            break
        melhorou = False
        # Operações de projetos atrasados, da última para a primeira, para permitir
        # antecipação sem mudar as precedências impostas pelo escalonador.
        for i in range(len(ordem) - 1, 0, -1):
            if ordem[i][0] not in atrasados:
                continue
            for destino in range(i - 1, -1, -1):
                if avaliados >= max_candidatos:
                    break
                proposta = ordem.copy()
                operacao = proposta.pop(i)
                proposta.insert(destino, operacao)
                candidato = executar_sequenciamento(dados, metodo=metodo, ordem_prioridade=proposta, incremental=incremental)
                avaliados += 1
                if _objetivo(candidato[2]) is None:
                    continue
                validar_cronograma(candidato, dados)
                z_novo = _objetivo(candidato[2])
                if z_novo < z:
                    atual, z, ordem = candidato, z_novo, proposta
                    aceitos += 1
                    melhorou = True
                    break
            if melhorou or avaliados >= max_candidatos:
                break
        if not melhorou:
            break
    return ResultadoBusca(atual, inicial, z, aceitos, avaliados)
