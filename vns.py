"""VNS (Variable Neighborhood Search) para sequenciamento de projetos.

Parte do MESMO cronograma tradicional usado pela busca local por inserção.
Explora inserção, troca e inversão de blocos da lista de prioridades.
Cada candidato é reconstruído pelo sequenciador original, que impõe
precedências, competências, número máximo de colaboradores e jornada.

Objetivo: soma ponderada de atraso total e makespan, normalizados em relação
à solução tradicional para que os pesos tenham interpretação comparável.
"""
from __future__ import annotations

from dataclasses import dataclass
import random

import pandas as pd

from busca_local import validar_cronograma
from sequenciamento import dias_atraso_uteis, executar_sequenciamento


@dataclass
class ResultadoVNS:
    resultado: tuple
    atraso_inicial: int
    atraso_final: int
    makespan_inicial: int
    makespan_final: int
    movimentos_aceitos: int
    candidatos_avaliados: int
    peso_atraso: float
    peso_makespan: float


def indicadores(projetos):
    """(atraso agregado, makespan em dias úteis), ou None se incompleto."""
    if not projetos or any(p['data_conclusao'] is None for p in projetos):
        return None
    atraso = sum(dias_atraso_uteis(p['prazo'], p['data_conclusao']) for p in projetos)
    inicio = min(p['inicio'] for p in projetos)
    fim = max(p['data_conclusao'] for p in projetos)
    makespan = len(pd.bdate_range(inicio, fim)) if fim >= inicio else 0
    return atraso, makespan


def _mover(ordem, tipo, rng):
    """Modifica prioridades, nunca diretamente as células alocadas."""
    n = len(ordem)
    if n < 2:
        return ordem.copy()
    i, j = sorted(rng.sample(range(n), 2))
    nova = ordem.copy()
    if tipo == 0:  # inserção
        item = nova.pop(j)
        nova.insert(i, item)
    elif tipo == 1:  # troca
        nova[i], nova[j] = nova[j], nova[i]
    else:  # inversão de um trecho
        nova[i:j + 1] = reversed(nova[i:j + 1])
    return nova


def executar_vns(dados, metodo='SJF', resultado_inicial=None,
                 max_iteracoes=5, max_candidatos=150,
                 peso_atraso=0.5, peso_makespan=0.5, semente=42, incremental=False):
    """VNS com perturbação e descida local, com orçamento global de avaliações.

    Iterações: ciclos de vizinhanças; candidatos: recálculos totais permitidos.
    Só guarda soluções que MELHOREM a função objetivo ponderada; a perturbação
    permite explorar temporariamente outras sequências dentro de cada ciclo.
    """
    metodo = metodo.upper()
    if metodo not in {'SJF', 'SPT'}:
        raise ValueError('Método deve ser SJF ou SPT.')
    if max_iteracoes < 0 or max_candidatos < 0:
        raise ValueError('Limites de busca devem ser não negativos.')
    if peso_atraso < 0 or peso_makespan < 0 or peso_atraso + peso_makespan <= 0:
        raise ValueError('Defina pesos não negativos, com pelo menos um positivo.')

    rng = random.Random(semente)
    inicial = resultado_inicial if resultado_inicial is not None else executar_sequenciamento(dados, metodo=metodo, incremental=incremental)
    validar_cronograma(inicial, dados)
    referencia = indicadores(inicial[2])
    if referencia is None:
        raise ValueError('O cronograma inicial não concluiu todos os projetos.')

    def pontuar(medidas):
        if medidas is None:
            return float('inf')
        # Normalização fixa pela solução tradicional, sem alterar os denominadores.
        return (peso_atraso * medidas[0] / max(1, referencia[0])
                + peso_makespan * medidas[1] / max(1, referencia[1]))

    melhor = inicial
    melhor_ordem = [(p['id'], tid) for p, tid in inicial[5]]
    melhor_medidas = referencia
    melhor_score = pontuar(referencia)
    avaliados = 0
    aceitos = 0
    cache = {tuple(melhor_ordem): (inicial, referencia)}

    def avaliar(ordem):
        nonlocal avaliados
        chave = tuple(ordem)
        if chave in cache:
            return cache[chave]
        if avaliados >= max_candidatos:
            return None, None
        candidato = executar_sequenciamento(dados, metodo=metodo, ordem_prioridade=ordem, incremental=incremental)
        avaliados += 1
        medidas = indicadores(candidato[2])
        if medidas is not None:
            validar_cronograma(candidato, dados)
        cache[chave] = (candidato, medidas)
        return candidato, medidas

    if len(melhor_ordem) < 2:
        return ResultadoVNS(melhor, *[referencia[0], referencia[0], referencia[1], referencia[1]],
                            0, 0, peso_atraso, peso_makespan)

    for _ in range(max_iteracoes):
        k = 0
        while k < 3 and avaliados < max_candidatos:
            # Perturbação: um ou mais movimentos na vizinhança k.
            ordem_atual = melhor_ordem.copy()
            for _ in range(k + 1):
                ordem_atual = _mover(ordem_atual, k, rng)
            candidato, medidas = avaliar(ordem_atual)
            if medidas is None:
                k += 1
                continue
            score = pontuar(medidas)

            # Busca local: amostra vizinhos das três estruturas, aceitando
            # apenas melhorias locais; o ponto perturbado pode ser pior.
            tentativas = 0
            while tentativas < 12 and avaliados < max_candidatos:
                tipo = rng.randrange(3)
                nova_ordem = _mover(ordem_atual, tipo, rng)
                novo_candidato, novas_medidas = avaliar(nova_ordem)
                tentativas += 1
                if novas_medidas is not None and pontuar(novas_medidas) < score - 1e-12:
                    ordem_atual, candidato, medidas = nova_ordem, novo_candidato, novas_medidas
                    score = pontuar(medidas)
                    tentativas = 0

            if score < melhor_score - 1e-12:
                melhor = candidato
                melhor_ordem = ordem_atual
                melhor_medidas = medidas
                melhor_score = score
                aceitos += 1
                k = 0
            else:
                k += 1

    return ResultadoVNS(melhor, referencia[0], melhor_medidas[0],
                        referencia[1], melhor_medidas[1], aceitos, avaliados,
                        peso_atraso, peso_makespan)
