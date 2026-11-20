"""Núcleo comum das heurísticas SJF e SPT para o TCC.

A versão tradicional preserva o critério original. A versão com folga dinâmica
coloca a menor folga estimada do projeto como primeiro critério de despacho.
"""
from __future__ import annotations

import io
import re
import math
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
import pandas as pd

from dados import DadosEntrada

ETAPAS = {
    'DESMONTAGEM':['desmontagem'], 'PERITAGEM':['peritagem'],
    'LIMPEZA_ROTOR':['limpeza','rotor'], 'BALANCEAR_ROTOR':['balance','rotor'],
    'PINTURA_ROTOR':['pintura','rotor'], 'ENSAIOS_ROTOR':['ensaio','rotor'],
    'LIMPEZA_ESTATOR':['limpeza','estator'], 'PINTURA_ESTATOR':['pintura','estator'],
    'RETIRADA_BOBINAS':['retirada','bobina'], 'FABRICACAO_BOBINAS':['fabricacao','bobina'],
    'MONTAGEM_BOBINAS':['montagem','bobina'], 'ENSAIOS_ESTATOR':['ensaio','estator'],
    'LIMPEZA_EXCITATRIZ':['limpeza','excitatriz'], 'BALANC_EXCITATRIZ':['balance','excitatriz'],
    'PINTURA_EXCITATRIZ':['pintura','excitatriz'], 'ENSAIOS_EXCITATRIZ':['ensaio','excitatriz'],
    'TROCADOR_CALOR':['trocador'], 'BASES_DISPOSITIVOS':['bases'],
    'MONTAGEM_GERAL':['montagem','geral'], 'ENSAIOS_FINAIS':['ensaio','finai']}

FLUXO = {
    'DESMONTAGEM':[], 'PERITAGEM':['DESMONTAGEM'],
    'LIMPEZA_ROTOR':['PERITAGEM'], 'BALANCEAR_ROTOR':['LIMPEZA_ROTOR'],
    'PINTURA_ROTOR':['BALANCEAR_ROTOR'], 'ENSAIOS_ROTOR':['PINTURA_ROTOR'],
    'LIMPEZA_ESTATOR':['PERITAGEM'], 'PINTURA_ESTATOR':['LIMPEZA_ESTATOR'],
    'RETIRADA_BOBINAS':['LIMPEZA_ESTATOR'], 'FABRICACAO_BOBINAS':['PERITAGEM'],
    'MONTAGEM_BOBINAS':['RETIRADA_BOBINAS','FABRICACAO_BOBINAS'],
    'ENSAIOS_ESTATOR':['PINTURA_ESTATOR','MONTAGEM_BOBINAS'],
    'LIMPEZA_EXCITATRIZ':['PERITAGEM'], 'BALANC_EXCITATRIZ':['LIMPEZA_EXCITATRIZ'],
    'PINTURA_EXCITATRIZ':['BALANC_EXCITATRIZ'], 'ENSAIOS_EXCITATRIZ':['PINTURA_EXCITATRIZ'],
    'TROCADOR_CALOR':['PERITAGEM'], 'BASES_DISPOSITIVOS':['PERITAGEM'],
    'MONTAGEM_GERAL':['ENSAIOS_ROTOR','ENSAIOS_ESTATOR','ENSAIOS_EXCITATRIZ'],
    'ENSAIOS_FINAIS':['MONTAGEM_GERAL']}
IGNORAR = ['bandagem']
JORNADA_MAXIMA = 9.0


def norm(valor):
    return unicodedata.normalize('NFKD', str(valor)).encode('ascii', 'ignore').decode().lower().strip()


def como_data(valor, padrao=None):
    if valor is None or pd.isna(valor):
        return padrao
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    convertido = pd.to_datetime(valor, errors='coerce', dayfirst=True)
    return padrao if pd.isna(convertido) else convertido.date()


def dias_uteis_assinados(inicio: date, fim: date) -> int:
    """Dias úteis disponíveis (positivos) ou vencidos (negativos).

    No sentido positivo, conta início e fim. No sentido negativo, conta os
    dias úteis posteriores ao prazo até a data atual, sem contar o próprio prazo.
    """
    if fim >= inicio:
        return int(len(pd.bdate_range(inicio, fim)))
    return -int(len(pd.bdate_range(fim + timedelta(days=1), inicio)))


def dias_atraso_uteis(prazo: date, conclusao: date) -> int:
    """Conta somente os dias úteis posteriores ao prazo até a conclusão."""
    if conclusao <= prazo:
        return 0
    return int(len(pd.bdate_range(prazo + timedelta(days=1), conclusao)))


def _relogio(minuto):
    return f"{minuto // 60:02d}:{minuto % 60:02d}"


def gerar_horario_intervalo(inicio_hh: float, horas: float) -> str:
    """Converte horas líquidas dentro de 07:30-12:00 e 13:00-17:30."""
    inicio = round(inicio_hh * 60)
    fim = round((inicio_hh + horas) * 60)
    if inicio < 0 or fim > 540 or fim <= inicio:
        raise ValueError(f'Intervalo de jornada inválido: {inicio_hh}, {horas}')
    partes = []
    if inicio < 270:
        a, b = 450 + inicio, 450 + min(fim, 270)
        if b > a:
            partes.append(f'{_relogio(a)}-{_relogio(b)}')
    if fim > 270:
        a, b = 780 + max(0, inicio - 270), 780 + fim - 270
        if b > a:
            partes.append(f'{_relogio(a)}-{_relogio(b)}')
    return ' | '.join(partes)


def gerar_horario_turno(horas: float) -> str:
    """Compatibilidade: jornada iniciada às 07:30."""
    return gerar_horario_intervalo(0, horas)


def ler_alocacoes_celula(turnos, atividades):
    """Cada linha de turno corresponde à mesma linha de atividade.

    Mantém compatibilidade com a planilha antiga (uma única atividade).
    """
    if not atividades:
        return []
    linhas_atividade = str(atividades).split('\n')
    linhas_turno = str(turnos or '').split('\n')
    if len(linhas_atividade) != len(linhas_turno):
        raise ValueError('Quantidade de turnos diferente da quantidade de atividades na célula.')
    return list(zip(linhas_turno, linhas_atividade))


def bloqueios_calendario(ws, datas_col):
    """Extrai indisponibilidades ANTES de limpar o calendário.

    Qualquer registro já preenchido (horário ou atividade) é compromisso fixo
    e bloqueia o dia inteiro. Mesclagens horizontais também são respeitadas.
    Retorna {(linha_colaborador, coluna_data): motivo}.
    """
    ocupados = {}
    for row in range(2, ws.max_row + 1, 2):
        if not ws.cell(row, 1).value:
            continue
        for col in datas_col:
            textos = [str(ws.cell(r, col).value).strip()
                      for r in (row, row + 1)
                      if ws.cell(r, col).value is not None
                      and str(ws.cell(r, col).value).strip()]
            if textos:
                ocupados[(row, col)] = ' | '.join(textos)
    for intervalo in ws.merged_cells.ranges:
        if intervalo.max_col <= 1 or intervalo.min_row < 2:
            continue
        valor = ws.cell(intervalo.min_row, intervalo.min_col).value
        if valor is None or not str(valor).strip():
            continue
        for row in range(2, ws.max_row + 1, 2):
            if not ws.cell(row, 1).value or not (
                    intervalo.min_row <= row + 1 and intervalo.max_row >= row):
                continue
            for col in range(max(2, intervalo.min_col), intervalo.max_col + 1):
                if col in datas_col:
                    ocupados.setdefault((row, col), str(valor).strip())
    return ocupados


def ler_bloqueios_resultado(wb):
    """Bloqueios explícitos gravados pelo sequenciador em cada solução."""
    if 'Indisponibilidades' not in wb.sheetnames:
        return set()
    ws = wb['Indisponibilidades']
    return {(int(l), int(c)) for l, c, *_ in ws.iter_rows(
        min_row=2, values_only=True) if l is not None and c is not None}


def _mapear_operacoes(ops):
    ops_norm = [norm(o) for o in ops]
    idx = {}
    for etapa, palavras in ETAPAS.items():
        achados = [i for i, nome in enumerate(ops_norm) if all(p in nome for p in palavras)]
        if len(achados) != 1:
            raise ValueError(f"Etapa {etapa}: esperava uma coluna, encontrei {len(achados)}: {[ops[i] for i in achados]}")
        idx[etapa] = achados[0]
    ignoradas = {i for i, nome in enumerate(ops_norm) if any(k in nome for k in IGNORAR)}
    preds = {i: [] for i in range(len(ops))}
    for etapa, anteriores in FLUXO.items():
        preds[idx[etapa]] = [idx[p] for p in anteriores]
    return preds, ignoradas


def _data_projeto(proj_df, linha, nomes_colunas, nome_coluna, padrao):
    coluna = next((c for c in nomes_colunas if norm(c) == norm(nome_coluna)), None)
    if coluna is None:
        return padrao
    return como_data(proj_df.iloc[linha][coluna], padrao)


def _profundidades(preds):
    memo = {}
    def prof(t):
        if t not in memo:
            memo[t] = 0 if not preds[t] else 1 + max(prof(p) for p in preds[t])
        return memo[t]
    for t in preds:
        prof(t)
    return memo


def _resolver_operacoes_vazias(proj, preds, coluna_atual):
    mudou = True
    while mudou:
        mudou = False
        for tid, tarefa in proj['tarefas'].items():
            if tarefa['concluida'] or tarefa['hh_total'] > 0:
                continue
            anteriores = [proj['tarefas'][p] for p in preds[tid]]
            if all(p['concluida'] for p in anteriores):
                tarefa['concluida'] = True
                tarefa['dia_conclusao'] = max([p['dia_conclusao'] for p in anteriores] + [coluna_atual - 1])
                mudou = True


def _dias_criticos_restantes(proj, preds):
    """Estimativa de caminho crítico em dias úteis, usando até 9 h/operação/dia."""
    memo = {}
    def caminho(tid):
        if tid not in memo:
            tarefa = proj['tarefas'][tid]
            restante = max(0.0, tarefa['hh_restante'])
            duracao = int(math.ceil(restante / JORNADA_MAXIMA))
            memo[tid] = duracao + max((caminho(p) for p in preds[tid]), default=0)
        return memo[tid]
    return max((caminho(t) for t in proj['tarefas']), default=0)


def _folga(proj, data_atual, preds):
    data_base = max(data_atual, proj['inicio'])
    disponiveis = dias_uteis_assinados(data_base, proj['prazo'])
    restante = _dias_criticos_restantes(proj, preds)
    return disponiveis - restante


def _nome_atividade(nome_projeto, nome_operacao):
    if '-' in nome_projeto:
        partes = nome_projeto.split('-')
        empresa = '-'.join(partes[:-1]).strip()
        numero = partes[-1].strip()
        return f"RS {numero} - {nome_operacao} - {empresa}"
    return f"{nome_projeto} - {nome_operacao}"


def executar_sequenciamento(dados: DadosEntrada, metodo='SJF', prioridade_folga=False, ordem_prioridade=None, incremental=False):
    metodo = metodo.upper()
    if metodo not in {'SJF', 'SPT'}:
        raise ValueError("Método deve ser 'SJF' ou 'SPT'.")

    ops, raw, proj_df = dados.operacoes, dados.colaboradores, dados.projetos
    preds, ignoradas = _mapear_operacoes(ops)
    profundidade = _profundidades(preds)
    posicao_insercao = ({chave: i for i, chave in enumerate(ordem_prioridade)}
                        if ordem_prioridade is not None else None)

    max_workers = {}
    for i in range(len(ops)):
        try:
            valor = raw.iloc[1, i + 2]
            max_workers[i] = max(1, int(valor)) if not pd.isna(valor) else 1
        except (ValueError, TypeError, IndexError):
            max_workers[i] = 1

    skills = {}
    for r in range(2, len(raw)):
        nome = str(raw.iloc[r, 1]).strip()
        if nome.lower() == 'nan' or not nome:
            continue
        skills[norm(nome)] = {i for i in range(len(ops)) if i + 2 < raw.shape[1] and raw.iloc[r, i + 2] == 1}

    # Identifica as colunas de data na aba Projetos sem depender da posição fixa.
    colunas = list(proj_df.columns)
    col_inicio = next((c for c in colunas if norm(c) == 'inicio'), None)
    col_prazo = next((c for c in colunas if norm(c) == 'prazo'), None)
    if col_inicio is None or col_prazo is None:
        raise ValueError("A aba 'Projetos' precisa conter as colunas 'INÍCIO' e 'PRAZO'.")

    wb = openpyxl.load_workbook(io.BytesIO(dados.bytes_planilha))
    if 'Cronograma' not in wb.sheetnames:
        raise ValueError("A planilha precisa conter a aba 'Cronograma'.")
    ws = wb['Cronograma']

    datas_col = {}
    for col in range(2, ws.max_column + 1):
        data = como_data(ws.cell(1, col).value)
        if data is not None:
            datas_col[col] = data
    if not datas_col:
        raise ValueError("Não foram encontradas datas na primeira linha da aba 'Cronograma'.")

    data_inicio_cal = min(datas_col.values())
    # Estende o horizonte até pelo menos 90 dias úteis após o maior prazo/data existente.
    ultima_data = max(max(datas_col.values()), max(como_data(proj_df.iloc[i][col_prazo], data_inicio_cal) for i in range(len(proj_df))))
    data_limite = ultima_data
    adicionados = 0
    while adicionados < 90:
        data_limite += timedelta(days=1)
        if data_limite.weekday() < 5:
            nova_col = ws.max_column + 1
            ws.cell(1, nova_col).value = datetime.combine(data_limite, datetime.min.time())
            datas_col[nova_col] = data_limite
            adicionados += 1

    # Planejamento inicial: comportamento antigo (texto preenchido bloqueia o dia).
    # Incremental: a agenda é imutável; registros de trabalho são reconhecidos
    # depois de criar os projetos e abatidos das cargas restantes.
    bloqueios = {} if incremental else bloqueios_calendario(ws, datas_col)
    if not incremental:
        for row in range(2, ws.max_row + 1, 2):
            for col in range(2, ws.max_column + 1):
                if (row, col) in bloqueios:
                    continue
                ws.cell(row, col).value = None
                ws.cell(row + 1, col).value = None

    projetos = []
    for r in range(len(proj_df)):
        nome = str(proj_df.iloc[r, 1]).strip()
        if nome.lower() == 'nan' or not nome:
            continue
        inicio = como_data(proj_df.iloc[r][col_inicio], data_inicio_cal)
        prazo = como_data(proj_df.iloc[r][col_prazo], inicio)
        tarefas = {}
        total = 0.0
        for i in range(len(ops)):
            valor = proj_df.iloc[r, i + 2]
            hh = 0.0 if pd.isna(valor) or i in ignoradas else max(0.0, float(valor))
            total += hh
            tarefas[i] = {'hh_restante': hh, 'hh_total': hh, 'concluida': False,
                          'dia_conclusao': -1, 'iniciada': False, 'dia_inicio': None}
        projetos.append({'id': r, 'nome': nome, 'inicio': inicio, 'prazo': prazo,
                         'total_hh': total, 'tarefas': tarefas, 'data_conclusao': None})

    if not projetos:
        raise ValueError("Nenhum projeto válido foi encontrado na aba 'Projetos'.")

    # Reconhece e preserva o trabalho previamente marcado, inclusive turnos
    # parciais, e distingue esses registros de férias/afastamentos.
    ocupacao_fixa = {}    # (linha, coluna) -> lista de intervalos reais (minutos)
    fixos_por_dia = {}   # (projeto_id, operacao_id, coluna) -> conjunto de linhas
    if incremental:
        from collections import defaultdict
        from gantt import _turno
        catalogo = {}
        for proj in projetos:
            for tid in proj['tarefas']:
                nome = _nome_atividade(proj['nome'], ops[tid])
                if nome in catalogo:
                    raise ValueError(f'Identificador de atividade repetido: {nome}')
                catalogo[nome] = (proj, tid)
        fixos_por_dia = defaultdict(set)
        for row in range(2, ws.max_row + 1, 2):
            if not ws.cell(row, 1).value:
                continue
            for col, data in datas_col.items():
                atv = ws.cell(row + 1, col).value
                turno = ws.cell(row, col).value
                if atv is None and turno is None:
                    continue
                if not atv or any(a.strip() not in catalogo for a in str(atv).split('\n')):
                    bloqueios[(row, col)] = str(atv or turno)  # ausência: dia inteiro
                    continue
                registros = ler_alocacoes_celula(turno, atv)
                intervalos = []
                for horario, nome_atv in registros:
                    proj, tid = catalogo[nome_atv.strip()]
                    if data.weekday() >= 5 or data < proj['inicio']:
                        raise ValueError(f'Atividade fixa em data inválida: {nome_atv}, {data}')
                    horas, _ = _turno(horario)
                    if horas <= 0:
                        raise ValueError(f'Horário fixo inválido: {nome_atv}, {horario}')
                    tarefa = proj['tarefas'][tid]
                    tarefa['hh_restante'] -= horas
                    if tarefa['hh_restante'] < -1e-6:
                        raise ValueError(f'HH fixas excedem HH previstas: {nome_atv}')
                    tarefa['iniciada'] = True
                    tarefa['dia_inicio'] = min(tarefa['dia_inicio'] or col, col)
                    tarefa['dia_conclusao'] = max(tarefa['dia_conclusao'], col)
                    fixos_por_dia[(proj['id'], tid, col)].add(row)
                    pares = re.findall(r'(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})', horario)
                    if not pares:
                        raise ValueError(f'Horário não reconhecido: {horario}')
                    for h1, m1, h2, m2 in pares:
                        a, b = int(h1)*60+int(m1), int(h2)*60+int(m2)
                        if not ((450 <= a < b <= 720) or (780 <= a < b <= 1050)):
                            raise ValueError(f'Turno fora da jornada: {nome_atv}, {horario}')
                        intervalos.append((a, b))
                intervalos.sort()
                if any(intervalos[i][1] > intervalos[i+1][0] for i in range(len(intervalos)-1)):
                    raise ValueError(f'Sobreposição de horário existente: linha {row}, data {data}')
                if sum((b-a) for a,b in intervalos) > 540 + 1e-6:
                    raise ValueError(f'Jornada existente excede 9h: linha {row}, data {data}')
                ocupacao_fixa[(row, col)] = intervalos
        for proj in projetos:
            for tarefa in proj['tarefas'].values():
                if tarefa['hh_restante'] <= 1e-6:
                    tarefa['hh_restante'] = 0.0
                    if tarefa['iniciada']:
                        tarefa['concluida'] = True
        # Se uma operação fixa tem sucessora já alocada, a auditoria final
        # verificará a precedência. Nunca deslocamos as células originais.

    if 'Indisponibilidades' in wb.sheetnames:
        del wb['Indisponibilidades']
    ws_bloc = wb.create_sheet('Indisponibilidades')
    ws_bloc.append(['Linha colaborador', 'Coluna calendário', 'Colaborador', 'Data', 'Registro original'])
    for (row, col), texto in sorted(bloqueios.items()):
        ws_bloc.append([row, col, ws.cell(row, 1).value, datas_col[col], texto])
    ws_bloc.freeze_panes = 'A2'

    def horario_livre(intervalos_ocupados, horas_maximas):
        """Retorna turnos não sobrepostos (pode atravessar o intervalo de almoço)."""
        vagas = []
        for inicio, fim in ((450, 720), (780, 1050)):
            cursor = inicio
            for a, b in sorted(intervalos_ocupados):
                if b <= cursor or a >= fim:
                    continue
                if a > cursor:
                    vagas.append((cursor, min(a, fim)))
                cursor = max(cursor, b)
                if cursor >= fim:
                    break
            if cursor < fim:
                vagas.append((cursor, fim))
        minutos = max(0, round(horas_maximas * 60))
        usados = []
        for a, b in vagas:
            if minutos <= 0:
                break
            duracao = min(b-a, minutos)
            if duracao:
                usados.append((a, a+duracao))
                minutos -= duracao
        return usados

    def formatar_intervalos(intervalos):
        return ' | '.join(f'{_relogio(a)}-{_relogio(b)}' for a,b in intervalos)

    # Mantém a ordem original para desempates; SJF tradicional ordena por carga total.
    ordem_original = {p['id']: i for i, p in enumerate(projetos)}
    makespan_col = 0
    sequencia = []
    colaboradores_agendados = []
    for row in range(2, ws.max_row + 1, 2):
        nome = ws.cell(row, 1).value
        if nome:
            colaboradores_agendados.append((row, str(nome).strip()))

    if incremental:
        for proj in projetos:
            if all(t['concluida'] or t['hh_total'] <= 1e-9 for t in proj['tarefas'].values()):
                finais = [t['dia_conclusao'] for t in proj['tarefas'].values() if t['dia_conclusao'] >= 0]
                if finais:
                    proj['data_conclusao'] = datas_col[max(finais)]

    # Somente dias úteis recebem alocações; colunas de sábado/domingo ficam vazias.
    for col in sorted(datas_col):
        data_atual = datas_col[col]
        if data_atual.weekday() >= 5:
            continue
        for proj in projetos:
            _resolver_operacoes_vazias(proj, preds, col)

        alocacoes_dia = {p['id']: {i: set(fixos_por_dia.get((p['id'], i, col), set())) for i in range(len(ops))} for p in projetos}
        horas_colaborador_dia = {}
        for row, nome_colab in colaboradores_agendados:
            if (row, col) in bloqueios:
                continue  # férias, licença, atestado ou outro compromisso prévio
            cel_turno = ws.cell(row, col)
            cel_atividade = ws.cell(row + 1, col)
            # O mesmo colaborador pode executar várias operações independentes
            # no dia; a soma das durações permanece limitada a 9h líquidas.
            chave_nome = norm(nome_colab)
            nome_skill = next((n for n in skills if chave_nome == n or chave_nome in n or n in chave_nome), None)
            if nome_skill is None:
                continue
            while True:
                horas_ja_alocadas = horas_colaborador_dia.get(chave_nome, 0.0)
                fixos = ocupacao_fixa.get((row, col), [])
                horas_ocupadas = sum((b-a)/60 for a,b in fixos)
                capacidade_restante = max(0.0, JORNADA_MAXIMA - horas_ja_alocadas - horas_ocupadas)
                if capacidade_restante <= 1e-9:
                    break
                candidatas = []
                for proj in projetos:
                    if data_atual < proj['inicio']:
                        continue
                    folga = _folga(proj, data_atual, preds) if prioridade_folga else 0
                    for tid, tarefa in proj['tarefas'].items():
                        if tarefa['concluida'] or tarefa['hh_restante'] <= 1e-9:
                            continue
                        if incremental and tarefa['iniciada'] and tarefa['dia_conclusao'] >= col:
                            continue  # trabalho fixo futuro: complementar somente depois dele
                        # Predecessoras devem acabar em dia ANTERIOR.
                        # Não liberamos sucessoras no mesmo dia, mesmo com HH sobrando.
                        predecessoras_ok = all(
                            proj['tarefas'][pid]['concluida'] and proj['tarefas'][pid]['dia_conclusao'] < col
                            for pid in preds[tid]
                        )
                        if not predecessoras_ok or tid not in skills[nome_skill]:
                            continue
                        if len(alocacoes_dia[proj['id']][tid]) >= max_workers[tid] and row not in alocacoes_dia[proj['id']][tid]:
                            continue
                        if metodo == 'SJF':
                            chave = ((folga, proj['total_hh'], ordem_original[proj['id']], tid)
                                     if prioridade_folga else
                                     (proj['total_hh'], ordem_original[proj['id']], tid))
                        else:
                            chave = ((folga, tarefa['hh_total'], profundidade[tid], ordem_original[proj['id']], tid)
                                     if prioridade_folga else
                                     (tarefa['hh_total'], profundidade[tid], ordem_original[proj['id']], tid))
                        if posicao_insercao is not None:
                            chave = (posicao_insercao.get((proj['id'], tid), len(posicao_insercao)),) + chave
                        candidatas.append((chave, proj, tid))
                if not candidatas:
                    break
                candidatas.sort(key=lambda item: item[0])
                _, proj, tid = candidatas[0]
                tarefa = proj['tarefas'][tid]
                horas = min(capacidade_restante, tarefa['hh_restante'])
                if horas <= 1e-9:
                    break
                atividade = _nome_atividade(proj['nome'], ops[tid])
                if incremental:
                    ocupados = fixos + horas_colaborador_dia.get((row, col, 'novos'), [])
                    trechos = horario_livre(ocupados, horas)
                    horas = sum((b-a)/60 for a,b in trechos)
                    if horas <= 1e-9:
                        break
                    turno = formatar_intervalos(trechos)
                    horas_colaborador_dia.setdefault((row, col, 'novos'), []).extend(trechos)
                else:
                    turno = gerar_horario_intervalo(horas_ja_alocadas, horas)
                cel_atividade.value = (str(cel_atividade.value) + '\n' if cel_atividade.value else '') + atividade
                cel_turno.value = (str(cel_turno.value) + '\n' if cel_turno.value else '') + turno
                cel_turno.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical='center')
                cel_atividade.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical='center')
                quantidade = len(str(cel_atividade.value).split('\n'))
                ws.row_dimensions[row].height = max(ws.row_dimensions[row].height or 15, 15 * quantidade)
                ws.row_dimensions[row + 1].height = max(ws.row_dimensions[row + 1].height or 15, 30 * quantidade)
                if not tarefa['iniciada']:
                    tarefa['iniciada'] = True
                    tarefa['dia_inicio'] = col
                    sequencia.append((proj, tid))
                tarefa['hh_restante'] = max(0.0, tarefa['hh_restante'] - horas)
                if tarefa['hh_restante'] <= 1e-9:
                    tarefa['concluida'] = True
                    tarefa['dia_conclusao'] = col
                # Conta colaboradores distintos, não o número de linhas no mesmo dia.
                alocacoes_dia[proj['id']][tid].add(row)
                horas_colaborador_dia[chave_nome] = horas_ja_alocadas + horas
                makespan_col = max(makespan_col, col)

        for proj in projetos:
            _resolver_operacoes_vazias(proj, preds, col + 1)
            if all(t['concluida'] for t in proj['tarefas'].values()) and proj['data_conclusao'] is None:
                conclusoes = [t['dia_conclusao'] for t in proj['tarefas'].values() if t['dia_conclusao'] >= 0]
                col_conclusao = max(conclusoes, default=col)
                proj['data_conclusao'] = datas_col.get(col_conclusao, data_atual)

        if all(p['data_conclusao'] is not None for p in projetos):
            break

    # Resolve novamente projetos compostos apenas por operações sem carga.
    for proj in projetos:
        if proj['data_conclusao'] is None and all(t['concluida'] for t in proj['tarefas'].values()):
            proj['data_conclusao'] = proj['inicio']

    # Registra indicadores para auditoria e comparação no TCC.
    if 'Indicadores' in wb.sheetnames:
        del wb['Indicadores']
    ws_ind = wb.create_sheet('Indicadores')
    ws_ind.append(['Projeto', 'Data de início', 'Prazo', 'Conclusão', 'Carga total (H/H)',
                   'Atraso (dias úteis)', 'Entregue no prazo', 'Folga estimada no início (dias úteis)', 'Status'])
    for proj in projetos:
        conclusao = proj['data_conclusao']
        atraso = None if conclusao is None else dias_atraso_uteis(proj['prazo'], conclusao)
        no_prazo = None if conclusao is None else conclusao <= proj['prazo']
        folga_inicial = dias_uteis_assinados(max(data_inicio_cal, proj['inicio']), proj['prazo']) - _dias_criticos_restantes(
            {'tarefas': {i: {'hh_restante': t['hh_total']} for i, t in proj['tarefas'].items()}}, preds)
        status = 'Concluído' if conclusao else 'Não concluído no horizonte'
        ws_ind.append([proj['nome'], proj['inicio'], proj['prazo'], conclusao, proj['total_hh'], atraso,
                       no_prazo, folga_inicial, status])
    for col, largura in enumerate([42, 15, 15, 15, 18, 20, 18, 34, 28], start=1):
        ws_ind.column_dimensions[openpyxl.utils.get_column_letter(col)].width = largura
    ws_ind.freeze_panes = 'A2'
    ws_ind.auto_filter.ref = ws_ind.dimensions

    if 'Sequência' in wb.sheetnames:
        del wb['Sequência']
    ws_seq = wb.create_sheet('Sequência')
    ws_seq.append(['Ordem', 'Projeto', 'Operação', 'H/H', 'Início', 'Fim'])
    def data_col(col):
        if not col:
            return 'não iniciada'
        valor = ws.cell(1, col).value
        return valor.strftime('%d/%m/%Y') if hasattr(valor, 'strftime') else valor
    sequencia_relatorio = (sorted(((p, tid) for p in projetos for tid, t in p['tarefas'].items()
                                  if t['iniciada']),
                                 key=lambda item: (item[0]['tarefas'][item[1]]['dia_inicio'], item[0]['id'], item[1]))
                          if incremental else sequencia)
    for ordem, (proj, tid) in enumerate(sequencia_relatorio, start=1):
        tarefa = proj['tarefas'][tid]
        fim = data_col(tarefa['dia_conclusao']) if tarefa['concluida'] else 'não concluída'
        ws_seq.append([ordem, proj['nome'], ops[tid], tarefa['hh_total'], data_col(tarefa['dia_inicio']), fim])
    for letra, largura in zip('ABCDEF', [8, 42, 34, 10, 14, 14]):
        ws_seq.column_dimensions[letra].width = largura
    ws_seq.freeze_panes = 'A2'

    return wb, ws, projetos, ops, makespan_col, sequencia
