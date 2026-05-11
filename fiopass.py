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
    ws['C7'] = 'Passagens :  (X)' if 'Passagens'      in servicos else 'Passagens :  ( )'
    ws['D7'] = 'Diárias: (X)'     if 'Diárias'        in servicos else 'Diárias: ( )'
    ws['E7'] = 'Terrestre: (X)'   if 'Terrestre'       in servicos else 'Terrestre: ( )'
    ws['F7'] = 'Aluguel de carro: (X)' if 'Aluguel de carro' in servicos else 'Aluguel de carro: ( )'

    ws['B15'] = get_col(row, 6)             # Nome completo
    ws['C15'] = format_date(get_col(row, 7)) # Data de nascimento
    ws['D15'] = get_col(row, 9)             # Cargo/Função
    ws['E15'] = get_col(row, 8)             # CPF
    ws['F15'] = get_col(row, 10)            # Nome banco
    ws['G15'] = get_col(row, 11)            # Agência
    ws['H15'] = get_col(row, 12)            # DV agência
    ws['I15'] = get_col(row, 13)            # Conta corrente
    ws['J15'] = get_col(row, 14)            # DV conta
    ws['K15'] = get_col(row, 15)            # Poupança

    # Trechos: cada segmento ocupa 9 colunas a partir de S(18), AB(27), AK(36), AT(45)
    # Offsets dentro do segmento: origem=0, destino=1, data_ida=3, hora_ida=4, data_volta=6, hora_volta=7
    SEGMENT_STARTS = [18, 27, 36, 45]
    SEGMENT_ROWS   = [15, 16, 17, 18]
    SEGMENT_FIELDS = [
        ('L', 0, False),
        ('M', 1, False),
        ('N', 3, True),
        ('O', 4, False),
        ('P', 6, True),
        ('Q', 7, False),
    ]
    for start_col, out_row in zip(SEGMENT_STARTS, SEGMENT_ROWS):
        for out_col, offset, is_date in SEGMENT_FIELDS:
            value = get_col(row, start_col + offset)
            ws[f'{out_col}{out_row}'] = format_date(value) if is_date else value

    ws['C19'] = get_col(row, 108)   # Justificativa fora do prazo
    ws['C20'] = get_col(row, 109)   # Observações

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
