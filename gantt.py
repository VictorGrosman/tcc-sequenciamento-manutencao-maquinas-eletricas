import re
import datetime as dt
from pathlib import Path
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.patches import Patch
from sequenciamento import ler_alocacoes_celula, ler_bloqueios_resultado

REGEX_TURNO = re.compile(r"(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})")


def _turno(texto):
    total = 0.0
    fim = None
    for h1, m1, h2, m2 in REGEX_TURNO.findall(str(texto or '')):
        ini = int(h1) + int(m1) / 60
        fim = int(h2) + int(m2) / 60
        total += fim - ini
    return total, fim


def _atividade(texto):
    partes = [p.strip() for p in str(texto).split(' - ')]
    if len(partes) >= 3 and partes[0].upper().startswith('RS'):
        return f'{partes[0]} - {partes[-1]}', ' - '.join(partes[1:-1])
    if len(partes) == 2:
        return partes[0], partes[1]
    return str(texto).strip(), 'Outros'


def _nome_metodo(metodo):
    nomes = {
        'SJF': 'SJF (Shortest Job First)',
        'SPT': 'SPT (Shortest Processing Time)',
        'SJF_FOLGA': 'SJF com folga dinâmica',
        'SPT_FOLGA': 'SPT com folga dinâmica',
        'SJF_INSERCAO': 'SJF + busca local (inserção)',
        'SPT_INSERCAO': 'SPT + busca local (inserção)',
        'SJF_VNS': 'SJF + VNS',
        'SPT_VNS': 'SPT + VNS',
    }
    chave = str(metodo or '').upper()
    return nomes.get(chave, str(metodo or 'Método não informado'))


def gerar_gantt(arquivo_excel, pasta_saida='.', metodo=''):
    arquivo_excel = Path(arquivo_excel)
    pasta = Path(pasta_saida)
    pasta.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(arquivo_excel, data_only=True)
    ws = wb['Cronograma']
    bloqueios = ler_bloqueios_resultado(wb)
    nome_metodo = _nome_metodo(metodo)
    sufixo = str(metodo or 'Metodo').replace(' ', '_')

    datas = {}
    for col in range(2, ws.max_column + 1):
        v = ws.cell(1, col).value
        datas[col] = v.date() if isinstance(v, dt.datetime) else (v if isinstance(v, dt.date) else None)

    registros = []
    for row in range(2, ws.max_row + 1, 2):
        nome = ws.cell(row, 1).value
        if not nome:
            continue
        for col in range(2, ws.max_column + 1):
            if (row, col) in bloqueios:
                continue  # ausência/compromisso original não é trabalho programado
            atividade_celula = ws.cell(row + 1, col).value
            if not atividade_celula:
                continue
            for turno, atividade in ler_alocacoes_celula(ws.cell(row, col).value, atividade_celula):
                horas, fim = _turno(turno)
                projeto, op = _atividade(atividade)
                registros.append({'Colaborador': str(nome).strip(), 'Coluna': col, 'Data': datas[col],
                                  'Turno': turno, 'Horas': horas, 'Hora_Fim': fim,
                                  'Projeto': projeto, 'Operação': op, 'Atividade': atividade})
    if not registros:
        raise ValueError("Nenhuma atividade encontrada na aba 'Cronograma'.")

    df = pd.DataFrame(registros)
    col_ini, col_fim = df.Coluna.min(), df.Coluna.max()
    datas_validas = all(datas.get(c) is not None for c in range(col_ini, col_fim + 1))
    n_dias_uteis = df['Data'].nunique()
    n_dias_horizonte = len(pd.bdate_range(df['Data'].min(), df['Data'].max())) if datas_validas else n_dias_uteis
    fim_vals = df.loc[df.Coluna == col_fim, 'Hora_Fim'].dropna()
    hora_fim = fim_vals.max() if not fim_vals.empty else None

    resumo_proj = (df.groupby('Projeto').agg(Inicio=('Coluna', 'min'), Fim=('Coluna', 'max'),
                   Horas=('Horas', 'sum'), Alocacoes=('Atividade', 'count'),
                   Dias_uteis_com_atividade=('Data', 'nunique')).reset_index())
    resumo_proj = resumo_proj.rename(columns={'Dias_uteis_com_atividade': 'Duração (dias úteis com atividade)'})
    if datas_validas:
        resumo_proj['Data início'] = resumo_proj.Inicio.map(datas)
        resumo_proj['Data fim'] = resumo_proj.Fim.map(datas)
    resumo_proj = resumo_proj.sort_values('Fim').drop(columns=['Inicio', 'Fim'])

    resumo_colab = df.groupby('Colaborador').agg(Horas=('Horas', 'sum'),
                        Dias_trabalhados=('Data', 'nunique')).reset_index()
    # A capacidade disponível exclui os dias bloqueados no intervalo efetivo.
    bloqueios_por_nome = {}
    for row in range(2, ws.max_row + 1, 2):
        nome = ws.cell(row, 1).value
        if nome:
            bloqueios_por_nome.setdefault(str(nome).strip(), set()).update(
                datas.get(col) for linha, col in bloqueios if linha == row
                and datas.get(col) is not None and df.Data.min() <= datas[col] <= df.Data.max()
                and datas[col].weekday() < 5)
    resumo_colab['Dias indisponíveis'] = resumo_colab.Colaborador.map(
        lambda nome: len(bloqueios_por_nome.get(nome, set())))
    resumo_colab['Dias disponíveis'] = (n_dias_horizonte - resumo_colab['Dias indisponíveis']).clip(lower=0)
    resumo_colab['Utilização (%)'] = resumo_colab.apply(
        lambda r: round(r.Horas / (9 * r['Dias disponíveis']) * 100, 1)
        if r['Dias disponíveis'] > 0 else 0.0, axis=1)

    def x(col):
        return mdates.date2num(datas[col]) if datas_validas else col

    def cores(rotulos):
        paleta = list(plt.get_cmap('tab20').colors) + list(plt.get_cmap('tab20b').colors) + list(plt.get_cmap('tab20c').colors)
        return {r: paleta[i % len(paleta)] for i, r in enumerate(sorted(rotulos))}

    def config(ax):
        if datas_validas:
            ax.xaxis_date()
            ax.xaxis.set_major_locator(mdates.AutoDateLocator())
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m/%Y'))
            plt.setp(ax.get_xticklabels(), rotation=45, ha='right')
        ax.grid(axis='x', linestyle=':', alpha=.5)
        ax.set_axisbelow(True)

    def marcar_fim(ax):
        """Marca o limite do último dia trabalhado e mostra a data correspondente."""
        limite = x(col_fim) + 1
        rotulo = (f"Fim: {datas[col_fim]:%d/%m/%Y}" if datas_validas
                  else f"Fim: coluna {col_fim}")
        ax.axvline(limite, color='red', linestyle='--', linewidth=1.6)
        ax.annotate(rotulo, xy=(limite, 1), xycoords=('data', 'axes fraction'),
                    xytext=(-4, 5), textcoords='offset points', color='red',
                    fontsize=9, ha='right', va='bottom', annotation_clip=False)

    fig, ax = plt.subplots(figsize=(16, max(5, .55 * df.Projeto.nunique() + 2)))
    projs = resumo_proj.Projeto.tolist()
    ops = sorted(df['Operação'].unique())
    cor_op = cores(ops)
    for i, proj in enumerate(projs):
        sub = df[df.Projeto == proj]
        for col, grupo in sub.groupby('Coluna'):
            ops_dia = sorted(grupo['Operação'].unique())
            altura = .8 / len(ops_dia)
            for k, op in enumerate(ops_dia):
                ax.barh(i - .4 + k * altura, 1, left=x(col), height=altura, align='edge',
                        color=cor_op[op], edgecolor='white', linewidth=.4)
    ax.set_yticks(range(len(projs)))
    ax.set_yticklabels(projs, fontsize=8)
    ax.invert_yaxis()
    config(ax)
    marcar_fim(ax)
    ax.set_title(f'Gantt por projeto – {nome_metodo}')
    ax.legend(handles=[Patch(color=cor_op[o], label=o) for o in ops], title='Operação',
              loc='upper left', bbox_to_anchor=(1.01, 1), fontsize=7)
    fig.tight_layout()
    caminho_proj = pasta / f'Gantt_Projetos_{sufixo}.png'
    fig.savefig(caminho_proj, dpi=200, bbox_inches='tight')
    plt.close(fig)

    colabs = sorted(df.Colaborador.unique())
    cor_proj = cores(projs)
    fig, ax = plt.subplots(figsize=(16, max(5, .5 * len(colabs) + 2)))
    for i, colab in enumerate(colabs):
        for _, r in df[df.Colaborador == colab].iterrows():
            ax.barh(i, 1, left=x(r.Coluna), height=.7, color=cor_proj[r.Projeto],
                    edgecolor='white', linewidth=.4)
    ax.set_yticks(range(len(colabs)))
    ax.set_yticklabels(colabs, fontsize=9)
    ax.invert_yaxis()
    config(ax)
    marcar_fim(ax)
    ax.set_title(f'Gantt por colaborador – {nome_metodo}')
    if len(projs) <= 25:
        ax.legend(handles=[Patch(color=cor_proj[p], label=p) for p in projs], title='Projeto',
                  loc='upper left', bbox_to_anchor=(1.01, 1), fontsize=7)
    fig.tight_layout()
    caminho_colab = pasta / f'Gantt_Colaboradores_{sufixo}.png'
    fig.savefig(caminho_colab, dpi=200, bbox_inches='tight')
    plt.close(fig)

    indicadores = pd.read_excel(arquivo_excel, sheet_name='Indicadores') if 'Indicadores' in wb.sheetnames else pd.DataFrame()
    caminho_resumo = pasta / f'Resumo_Makespan_Gantt_{sufixo}.xlsx'
    resumo_geral = pd.DataFrame({'Indicador': ['Método', 'Início do cronograma', 'Fim (makespan)',
        'Horário de término', 'Duração (dias corridos)', 'Dias úteis com alocação',
        'Projetos', 'Colaboradores', 'Total de H/H alocado'],
        'Valor': [nome_metodo, str(df.Data.min()) if datas_validas else f'coluna {col_ini}',
        str(df.Data.max()) if datas_validas else f'coluna {col_fim}',
        '—' if hora_fim is None else f'{int(hora_fim):02d}:{int(round((hora_fim % 1) * 60)):02d}',
        (df.Data.max() - df.Data.min()).days + 1 if datas_validas else '—', n_dias_uteis,
        df.Projeto.nunique(), df.Colaborador.nunique(), round(df.Horas.sum(), 1)]})
    with pd.ExcelWriter(caminho_resumo, engine='openpyxl') as writer:
        resumo_geral.to_excel(writer, sheet_name='Makespan', index=False)
        resumo_proj.to_excel(writer, sheet_name='Projetos', index=False)
        resumo_colab.to_excel(writer, sheet_name='Colaboradores', index=False)
        df.drop(columns=['Coluna', 'Hora_Fim', 'Data']).to_excel(writer, sheet_name='Alocações', index=False)
        if not indicadores.empty:
            indicadores.to_excel(writer, sheet_name='Indicadores de prazo', index=False)

    return {'gantt_projetos': caminho_proj, 'gantt_colaboradores': caminho_colab, 'resumo': caminho_resumo}
