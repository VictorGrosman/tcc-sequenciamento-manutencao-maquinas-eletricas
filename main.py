"""Compara, a partir da mesma solução inicial, SJF/SPT, busca local e VNS."""
from pathlib import Path
from time import perf_counter
from threading import Event, Thread
import csv
import argparse

from dados import carregar_dados
from Shortest_Job_First import salvar_resultado
from gantt import gerar_gantt
from busca_local import executar_busca_local, validar_cronograma
from sequenciamento import executar_sequenciamento
from vns import executar_vns, indicadores

PASTA_BASE = Path(__file__).parent
ARQUIVO_DADOS = PASTA_BASE / 'rev.09 - Cronograma Nishi Python Atualizado.xlsx'
MAX_ITERACOES = 5
MAX_CANDIDATOS = 150
PESO_ATRASO = 0.5
PESO_MAKESPAN = 0.5
SEMENTE_VNS = 42


def medir_execucao(funcao):
    """Retorna resultado, tempo, pico de RSS e RSS inicial, em MiB."""
    try:
        import psutil
        processo = psutil.Process()
    except ImportError:
        processo = None

    parada = Event()
    inicio_ram = processo.memory_info().rss if processo else None
    pico_ram = [inicio_ram or 0]
    if processo:
        def monitorar():
            while not parada.wait(0.05):
                try:
                    pico_ram[0] = max(pico_ram[0], processo.memory_info().rss)
                except psutil.Error:
                    pass
        monitor = Thread(target=monitorar, daemon=True)
        monitor.start()

    inicio = perf_counter()
    try:
        resultado = funcao()
    finally:
        tempo = perf_counter() - inicio
        if processo:
            parada.set()
            monitor.join()
            pico_ram[0] = max(pico_ram[0], processo.memory_info().rss)
    mib = 1024 ** 2
    return resultado, tempo, pico_ram[0] / mib if processo else None, inicio_ram / mib if processo else None


def data_final(resultado):
    conclusoes = [p['data_conclusao'] for p in resultado[2]]
    return max(conclusoes) if conclusoes and all(conclusoes) else None


def executar_cenario(nome, metodo, tipo, dados, tradicional=None, incremental=False):
    """Cenários partem do MESMO objeto tradicional, sem encadear otimizações."""
    if tipo == 'tradicional':
        pasta = PASTA_BASE / f'solucao_{metodo}_tradicional' 
        nome_arquivo = f'Cronograma_{metodo}_Tradicional.xlsx'
    elif tipo == 'insercao':
        pasta = PASTA_BASE / f'solucao_{metodo}_insercao'
        nome_arquivo = f'Cronograma_{metodo}_Insercao.xlsx'
    else:
        pasta = PASTA_BASE / f'solucao_{metodo}_VNS'
        nome_arquivo = f'Cronograma_{metodo}_VNS.xlsx'
    if incremental:
        pasta = pasta.with_name(pasta.name + '_incremental')
    pasta.mkdir(parents=True, exist_ok=True)

    def trabalho():
        if tipo == 'tradicional':
            resultado = executar_sequenciamento(dados, metodo=metodo, incremental=incremental)
            validar_cronograma(resultado, dados)
            busca = None
        elif tipo == 'insercao':
            busca = executar_busca_local(dados, metodo=metodo,
                                        max_iteracoes=MAX_ITERACOES,
                                        max_candidatos=MAX_CANDIDATOS,
                                        resultado_inicial=tradicional, incremental=incremental)
            resultado = busca.resultado
        else:
            busca = executar_vns(dados, metodo=metodo, resultado_inicial=tradicional,
                                max_iteracoes=MAX_ITERACOES,
                                max_candidatos=MAX_CANDIDATOS,
                                peso_atraso=PESO_ATRASO, peso_makespan=PESO_MAKESPAN,
                                semente=SEMENTE_VNS, incremental=incremental)
            resultado = busca.resultado
        caminho = pasta / nome_arquivo
        salvar_resultado(resultado[0], caminho)
        gerar_gantt(caminho, pasta, metodo=metodo if tipo == 'tradicional' else f'{metodo}_{"INSERCAO" if tipo == "insercao" else "VNS"}')
        return resultado, busca

    (resultado, busca), tempo, pico_ram, inicio_ram = medir_execucao(trabalho)
    medidas = indicadores(resultado[2])
    conclusao = data_final(resultado)
    final_fmt = conclusao.strftime('%d/%m/%Y') if conclusao else 'não concluído'
    print(f'\n{nome}: conclusão global em {final_fmt}; '
          f'makespan {medidas[1] if medidas else "N/D"} dias úteis; '
          f'atraso total {medidas[0] if medidas else "N/D"} dias úteis.')
    if busca is not None:
        print(f'  Candidatos avaliados: {busca.candidatos_avaliados}; melhorias aceitas: {busca.movimentos_aceitos}.')
    print(f'  Tempo: {tempo:.2f}s; RAM inicial: {inicio_ram:.1f} MiB; pico: {pico_ram:.1f} MiB'
          if pico_ram is not None else f'  Tempo: {tempo:.2f}s; RAM indisponível.')
    print(f'  Arquivo: {pasta / nome_arquivo}')

    linha = {'Cenário': nome, 'Método base': metodo,
             'Data final global': final_fmt,
             'Makespan (dias úteis)': medidas[1] if medidas else '',
             'Atraso total (dias úteis)': medidas[0] if medidas else '',
             'Tempo (s)': round(tempo, 3),
             'RAM inicial (MiB)': inicio_ram, 'Pico RAM (MiB)': pico_ram,
             'Candidatos testados': busca.candidatos_avaliados if busca else '',
             'Movimentos aceitos': busca.movimentos_aceitos if busca else '',
             'Atraso inicial (dias úteis)': busca.atraso_inicial if busca else '',
             'Atraso final (dias úteis)': busca.atraso_final if busca else ''}
    return resultado, linha


def main():
    parser = argparse.ArgumentParser(description='Comparação de cronogramas SJF, SPT, busca local e VNS')
    parser.add_argument('--arquivo', type=Path, default=ARQUIVO_DADOS,
                        help='Excel contendo Colaboradores, Projetos e Cronograma')
    parser.add_argument('--incremental', action='store_true',
                        help='Preserva todas as alocações e compromissos existentes; preenche apenas horários livres')
    args = parser.parse_args()
    dados = carregar_dados(args.arquivo)
    modo = 'INCREMENTAL' if args.incremental else 'INICIAL'
    print(f'Modo {modo}; planilha: {args.arquivo}')
    resumo = []
    for metodo in ('SJF', 'SPT'):
        # Tradicional fica na memória e serve como ponto de partida isolado
        # para busca local e VNS. Não usa a solução da busca local no VNS.
        tradicional, linha = executar_cenario(f'{metodo} tradicional', metodo, 'tradicional', dados, incremental=args.incremental)
        resumo.append(linha)
        _, linha = executar_cenario(f'{metodo} + busca local', metodo, 'insercao', dados, tradicional, incremental=args.incremental)
        resumo.append(linha)
        _, linha = executar_cenario(f'{metodo} + VNS', metodo, 'vns', dados, tradicional, incremental=args.incremental)
        resumo.append(linha)
    destino = PASTA_BASE / ('Resumo_Desempenho_Incremental.csv' if args.incremental else 'Resumo_Desempenho_Cenarios.csv')
    with destino.open('w', newline='', encoding='utf-8-sig') as f:
        escritor = csv.DictWriter(f, fieldnames=list(resumo[0]))
        escritor.writeheader()
        escritor.writerows(resumo)
    print(f'\nComparação dos 6 cenários: {destino}')


if __name__ == '__main__':
    main()
