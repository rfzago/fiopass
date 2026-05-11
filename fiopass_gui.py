#!/usr/bin/env python3
"""
fiopass_gui.py - Interface gráfica do FioPass.
"""

import os
import sys
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from fiopass import run_generation, parse_input, get_col


class FioPassApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title('FioPass - Gerador de Formulários FIOTEC')
        self.resizable(False, False)
        self._output_dir = None
        self._queue = queue.Queue()
        self._rows = []
        self._row_vars = []
        self._build_ui()
        self._center()
        self._poll_queue()

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f'+{x}+{y}')

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.configure(padx=16, pady=12)

        # Arquivo de entrada
        frm_in = ttk.LabelFrame(self, text='Arquivo de entrada', padding=8)
        frm_in.pack(fill='x', pady=(0, 8))
        self._input_var = tk.StringVar()
        ttk.Entry(frm_in, textvariable=self._input_var, width=56).pack(side='left', expand=True, fill='x')
        ttk.Button(frm_in, text='Selecionar…', command=self._browse_input).pack(side='left', padx=(6, 0))

        # Seleção de registros (oculto até carregar arquivo)
        self._frm_sel = ttk.LabelFrame(self, text='Registros encontrados', padding=8)
        hdr = ttk.Frame(self._frm_sel)
        hdr.pack(fill='x', pady=(0, 4))
        ttk.Button(hdr, text='Selecionar todos', command=self._select_all).pack(side='left')
        ttk.Button(hdr, text='Desmarcar todos', command=self._deselect_all).pack(side='left', padx=(6, 0))
        self._lbl_count = ttk.Label(hdr, text='')
        self._lbl_count.pack(side='right')

        self._chk_canvas = tk.Canvas(self._frm_sel, height=140, highlightthickness=0)
        sb_chk = ttk.Scrollbar(self._frm_sel, orient='vertical', command=self._chk_canvas.yview)
        self._chk_canvas.configure(yscrollcommand=sb_chk.set)
        sb_chk.pack(side='right', fill='y')
        self._chk_canvas.pack(fill='x', expand=True)

        self._chk_inner = ttk.Frame(self._chk_canvas)
        self._chk_win = self._chk_canvas.create_window((0, 0), window=self._chk_inner, anchor='nw')
        self._chk_inner.bind(
            '<Configure>',
            lambda e: self._chk_canvas.configure(scrollregion=self._chk_canvas.bbox('all')),
        )
        self._chk_canvas.bind(
            '<Configure>',
            lambda e: self._chk_canvas.itemconfig(self._chk_win, width=e.width),
        )

        # Pasta de saída
        self._frm_out = ttk.LabelFrame(self, text='Pasta de saída', padding=8)
        self._frm_out.pack(fill='x', pady=(0, 12))
        self._output_var = tk.StringVar()
        ttk.Entry(self._frm_out, textvariable=self._output_var, width=56).pack(side='left', expand=True, fill='x')
        ttk.Button(self._frm_out, text='Selecionar…', command=self._browse_output).pack(side='left', padx=(6, 0))

        # Botão gerar
        self._btn_generate = ttk.Button(self, text='Gerar Formulários', command=self._start)
        self._btn_generate.pack(pady=(0, 10))

        # Log
        frm_log = ttk.LabelFrame(self, text='Log', padding=8)
        frm_log.pack(fill='both', expand=True, pady=(0, 8))
        self._log = tk.Text(frm_log, height=12, state='disabled',
                            wrap='word', font=('Courier', 10))
        sb = ttk.Scrollbar(frm_log, command=self._log.yview)
        self._log.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self._log.pack(fill='both', expand=True)

        # Botão abrir pasta (oculto até concluir)
        self._btn_open = ttk.Button(self, text='Abrir pasta de saída', command=self._open_output)

    # ── Ações ─────────────────────────────────────────────────────────────────

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title='Selecionar arquivo de entrada',
            filetypes=[('Arquivos Excel', '*.xls *.xlsx'), ('Todos os arquivos', '*.*')],
        )
        if path:
            self._input_var.set(path)
            self._load_rows(path)

    def _load_rows(self, path):
        try:
            rows = parse_input(path)
        except Exception as exc:
            messagebox.showerror('Erro', f'Não foi possível ler o arquivo:\n{exc}')
            return

        self._rows = rows
        self._row_vars = []

        for w in self._chk_inner.winfo_children():
            w.destroy()

        for row in rows:
            cpf = get_col(row, 8).strip('"').strip()
            periodo = get_col(row, 3)
            var = tk.BooleanVar(value=True)
            self._row_vars.append(var)
            ttk.Checkbutton(
                self._chk_inner,
                text=f'{cpf}   {periodo}',
                variable=var,
                command=self._update_count,
            ).pack(anchor='w', padx=4, pady=1)

        if rows:
            self._frm_sel.pack(fill='x', pady=(0, 8), before=self._frm_out)
        else:
            self._frm_sel.pack_forget()

        self._update_count()

    def _browse_output(self):
        path = filedialog.askdirectory(title='Selecionar pasta de saída')
        if path:
            self._output_var.set(path)

    def _select_all(self):
        for v in self._row_vars:
            v.set(True)
        self._update_count()

    def _deselect_all(self):
        for v in self._row_vars:
            v.set(False)
        self._update_count()

    def _update_count(self):
        selected = sum(v.get() for v in self._row_vars)
        total = len(self._row_vars)
        self._lbl_count.config(text=f'{selected}/{total} selecionados')

    def _start(self):
        input_file  = self._input_var.get().strip()
        output_base = self._output_var.get().strip()

        if not input_file:
            messagebox.showwarning('Atenção', 'Selecione o arquivo de entrada.')
            return
        if not output_base:
            messagebox.showwarning('Atenção', 'Selecione a pasta de saída.')
            return

        selected_rows = [row for row, var in zip(self._rows, self._row_vars) if var.get()]
        if self._row_vars and not selected_rows:
            messagebox.showwarning('Atenção', 'Selecione ao menos um registro.')
            return

        self._clear_log()
        self._btn_open.pack_forget()
        self._btn_generate.configure(state='disabled')

        rows_arg = selected_rows if self._row_vars else None

        def worker():
            try:
                out = run_generation(
                    input_file, output_base,
                    progress_callback=lambda m: self._queue.put(('log', m)),
                    selected_rows=rows_arg,
                )
                self._queue.put(('done', out))
            except Exception as exc:
                self._queue.put(('error', str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def _open_output(self):
        if self._output_dir and os.path.isdir(self._output_dir):
            if sys.platform == 'darwin':
                os.system(f'open "{self._output_dir}"')
            else:
                os.startfile(self._output_dir)

    # ── Log helpers ───────────────────────────────────────────────────────────

    def _write_log(self, message):
        self._log.configure(state='normal')
        self._log.insert('end', message + '\n')
        self._log.see('end')
        self._log.configure(state='disabled')

    def _clear_log(self):
        self._log.configure(state='normal')
        self._log.delete('1.0', 'end')
        self._log.configure(state='disabled')

    # ── Fila de progresso (thread-safe) ───────────────────────────────────────

    def _poll_queue(self):
        try:
            while True:
                kind, value = self._queue.get_nowait()
                if kind == 'log':
                    self._write_log(value)
                elif kind == 'done':
                    self._output_dir = value
                    self._write_log('Concluído.')
                    self._btn_generate.configure(state='normal')
                    self._btn_open.pack(pady=(0, 4))
                elif kind == 'error':
                    self._write_log(f'Erro: {value}')
                    self._btn_generate.configure(state='normal')
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)


if __name__ == '__main__':
    app = FioPassApp()
    app.mainloop()
