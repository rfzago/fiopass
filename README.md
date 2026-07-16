# FioPass - Gerador de Formulários FIOTEC

Lê um arquivo de entrada com dados de beneficiários e gera automaticamente um formulário **Solicitação de Passagens e Diárias** (Anexo I FIOTEC) em `.xlsx` para cada registro.

## Requisitos

```bash
pip install openpyxl pyinstaller
```

## Uso

### Interface gráfica
```bash
python3 fiopass_gui.py
```

1. Selecione o arquivo de entrada (`.xls`)
2. Selecione a pasta de saída
3. Clique em **Gerar Formulários**

Os arquivos gerados são salvos em uma subpasta nomeada com a data/hora da execução (`YYYYMMDD_HHMM`) dentro da pasta escolhida, com um arquivo `.xlsx` por beneficiário nomeado pelo CPF.

### Linha de comando
```bash
python3 fiopass.py <arquivo_entrada.xls>
```

Os arquivos gerados são salvos em `YYYYMMDD_HHMM/` na mesma pasta do script.

---

## Distribuição

### macOS — gerar `.app`
```bash
python3 -m PyInstaller -y --windowed \
  --name "FioPass" \
  --add-data "Anexo I_PLANILHA PASSAGENS E DIÁRIAS 2026.xlsx:." \
  --add-data "VERSION:." \
  fiopass_gui.py
```
O `.app` fica em `dist/FioPass.app`. Para distribuir, comprima a pasta `dist/FioPass.app` e envie — o destinatário arrasta para Aplicativos.

> **Importante:** use sempre `python3 -m PyInstaller` em vez do comando `pyinstaller` diretamente. Isso garante que o build usa o mesmo Python do ambiente onde o projeto roda, evitando erros de módulos não encontrados (como `_tkinter`).

### Windows — gerar `.exe`
No Windows, execute o comando abaixo no Prompt de Comando ou PowerShell:
```bash
python -m PyInstaller -y --windowed ^
  --name "FioPass" ^
  --add-data "Anexo I_PLANILHA PASSAGENS E DIÁRIAS 2026.xlsx;." ^
  --add-data "VERSION;." ^
  fiopass_gui.py
```
O `.exe` fica em `dist\FioPass\FioPass.exe`. No Windows o separador do `--add-data` é `;` em vez de `:`.

> **Nota:** o build deve ser executado na plataforma de destino — o `.app` gerado no Mac não roda no Windows e vice-versa.
