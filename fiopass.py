#!/usr/bin/env python3
"""
fiopass.py - Gera formulários FIOTEC (Passagens e Diárias) a partir de arquivo de entrada.

Uso: python fiopass.py <arquivo_entrada.xls>
"""

import sys
import os
from html.parser import HTMLParser
from datetime import datetime
import openpyxl

TEMPLATE_FILENAME = 'Anexo I_PLANILHA PASSAGENS E DIÁRIAS 2026.xlsx'


def get_resource_path(filename):
    """Resolve o caminho do arquivo tanto em execução normal quanto em bundle PyInstaller."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)


class TableParser(HTMLParser):
    """Lê arquivo XLS no formato HTML e extrai as linhas da tabela."""

    def __init__(self):
        super().__init__()
        self.rows = []
        self._current_row = []
        self._current_cell = ''
        self._in_cell = False

    def handle_starttag(self, tag, attrs):
        if tag == 'tr':
            self._current_row = []
        elif tag in ('td', 'th'):
            self._in_cell = True
            self._current_cell = ''

    def handle_endtag(self, tag):
        if tag in ('td', 'th'):
            self._current_row.append(self._current_cell.strip())
            self._in_cell = False
        elif tag == 'tr' and self._current_row:
            self.rows.append(self._current_row)

    def handle_data(self, data):
        if self._in_cell:
            self._current_cell += data


def parse_input(filepath):
    with open(filepath, encoding='utf-8') as f:
        content = f.read()
    parser = TableParser()
    parser.feed(content)
    return parser.rows


def get_col(row, index):
    return row[index].strip() if index < len(row) else ''


def format_date(value):
    try:
        return datetime.strptime(value, '%Y-%m-%d').strftime('%d/%m/%Y')
    except ValueError:
        return value


def generate_form(row, output_dir):
    cpf = get_col(row, 0).strip('"').strip()

    wb = openpyxl.load_workbook(get_resource_path(TEMPLATE_FILENAME))
    ws = wb.active

    ws['C4'] = get_col(row, 2)            # Identificação do evento
    ws['C5'] = get_col(row, 3)            # Data
    ws['C6'] = get_col(row, 4)            # Local

    servicos = [s.strip() for s in get_col(row, 5).split(',')]
    ws['C7'] = 'Passagens :  (X)' if 'Passagens Aéreas'    in servicos else 'Passagens :  ( )'
    ws['D7'] = 'Diárias: (X)'     if 'Diárias'             in servicos else 'Diárias: ( )'
    ws['E7'] = 'Terrestre: (X)'   if 'Passagens Terrestres' in servicos else 'Terrestre: ( )'
    ws['F7'] = 'Aluguel de carro: (X)' if 'Aluguel de carro' in servicos else 'Aluguel de carro: ( )'

    ws['B15'] = get_col(row, 6)             # Nome completo
    ws['C15'] = format_date(get_col(row, 7)) # Data de nascimento
    ws['D15'] = get_col(row, 9)             # Cargo/Função
    ws['E15'] = get_col(row, 8)             # CPF
    ws['F15'] = get_col(row, 11)            # Nome banco
    ws['G15'] = get_col(row, 12)            # Agência
    ws['H15'] = get_col(row, 13)            # DV agência
    ws['I15'] = get_col(row, 14)            # Conta corrente
    ws['J15'] = get_col(row, 15)            # DV conta
    ws['K15'] = get_col(row, 16)            # Poupança

    # Trechos: cada slot ocupa 8 colunas a partir de 19, na ordem: Tipo de Trecho
    # (Ida/Volta), Origem, Destino, Tipo Localidade de Destino, Data, Período,
    # Tipo Deslocamento, Observações sobre o deslocamento. Até 10 slots
    # reservados; para de ler no primeiro slot sem origem preenchida.
    TRECHO_START = 19
    TRECHO_WIDTH = 8
    MAX_TRECHOS  = 10
    OFFSET_ORIGEM, OFFSET_DESTINO, OFFSET_DATA, OFFSET_TURNO, OFFSET_TIPO_DESLOCAMENTO = 1, 2, 4, 5, 6

    trechos = []
    for i in range(MAX_TRECHOS):
        base = TRECHO_START + i * TRECHO_WIDTH
        origem = get_col(row, base + OFFSET_ORIGEM)
        if not origem:
            break
        trechos.append({
            'origem':           origem,
            'destino':          get_col(row, base + OFFSET_DESTINO),
            'data':             get_col(row, base + OFFSET_DATA),
            'turno':            get_col(row, base + OFFSET_TURNO),
            'tipo_deslocamento': get_col(row, base + OFFSET_TIPO_DESLOCAMENTO),
        })

    def origem_com_tipo(t):
        return f"{t['origem']}\n{t['tipo_deslocamento']}" if t['tipo_deslocamento'] else t['origem']

    # Pareia trechos consecutivos em viagens de ida-e-volta, na ordem em que
    # aparecem (independente da direção declarada): o trecho i fornece
    # origem/destino/data de ida da linha, e o trecho i+1 fornece a data de
    # volta. Um único trecho gera uma linha só de ida.
    output_rows = []
    if len(trechos) == 1:
        t = trechos[0]
        output_rows.append({
            'origem': origem_com_tipo(t), 'destino': t['destino'],
            'data_ida': t['data'], 'hora_ida': t['turno'],
            'data_volta': '', 'hora_volta': '',
        })
    else:
        for i in range(len(trechos) - 1):
            t, t_next = trechos[i], trechos[i + 1]
            output_rows.append({
                'origem': origem_com_tipo(t), 'destino': t['destino'],
                'data_ida': t['data'], 'hora_ida': t['turno'],
                'data_volta': t_next['data'], 'hora_volta': t_next['turno'],
            })

    SEGMENT_ROWS = [15, 16, 17, 18]
    for out_row, seg in zip(SEGMENT_ROWS, output_rows):
        ws[f'L{out_row}'] = seg['origem']
        ws[f'M{out_row}'] = seg['destino']
        ws[f'N{out_row}'] = format_date(seg['data_ida'])
        ws[f'O{out_row}'] = seg['hora_ida']
        ws[f'P{out_row}'] = format_date(seg['data_volta'])
        ws[f'Q{out_row}'] = seg['hora_volta']

    ws['C19'] = get_col(row, 99)    # Justificativa fora do prazo
    ws['C20'] = get_col(row, 100)   # Observações

    output_path = os.path.join(output_dir, f'{cpf}.xlsx')
    wb.save(output_path)
    return output_path


def run_generation(input_file, base_output_dir, progress_callback=None, selected_rows=None):
    """
    Lê input_file e gera um xlsx por registro em base_output_dir/YYYYMMDD_HHMM/.
    progress_callback(msg) é chamado a cada evento de progresso.
    selected_rows, se fornecido, é usado em vez de parse_input(input_file).
    Retorna o caminho da pasta de saída criada.
    """
    run_stamp = datetime.now().strftime('%Y%m%d_%H%M')
    output_dir = os.path.join(base_output_dir, run_stamp)
    os.makedirs(output_dir, exist_ok=True)

    rows = selected_rows if selected_rows is not None else parse_input(input_file)

    if progress_callback:
        progress_callback(f'{len(rows)} registro(s) encontrado(s) em {os.path.basename(input_file)}')

    for i, row in enumerate(rows, 1):
        output_path = generate_form(row, output_dir)
        cpf = get_col(row, 0).strip('"')
        if progress_callback:
            progress_callback(f'  [{i}] CPF {cpf} → {os.path.basename(output_path)}')

    return output_dir


def main():
    if len(sys.argv) < 2:
        print(f'Uso: python {os.path.basename(sys.argv[0])} <arquivo_entrada.xls>')
        sys.exit(1)

    input_file = sys.argv[1]

    if not os.path.isfile(input_file):
        print(f'Erro: arquivo não encontrado: {input_file}')
        sys.exit(1)

    template_path = get_resource_path(TEMPLATE_FILENAME)
    if not os.path.isfile(template_path):
        print(f'Erro: template não encontrado: {template_path}')
        sys.exit(1)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    run_generation(input_file, script_dir, progress_callback=print)
    print('Concluído.')


if __name__ == '__main__':
    main()
