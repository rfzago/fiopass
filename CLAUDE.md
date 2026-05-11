# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

`fiopass` is a Python tool (likely a Flask web app) that generates FIOTEC travel request forms — the official "Solicitação de Passagens e Diárias" spreadsheet used by Fiocruz/FIOTEC projects.

**Core workflow:** users upload an input file in the format of `formulario_345_2026-05-08_170531.xls` (which may contain multiple beneficiaries). The program reads each beneficiary record and produces **one output `.xlsx` file per record**, filled out according to the layout of `Anexo I_PLANILHA PASSAGENS E DIÁRIAS 2026.xlsx`.

## Input Format

Input files (e.g. `formulario_345_2026-05-08_170531.xls`) are **HTML documents saved with a `.xls` extension** (Excel-compatible HTML, using `xmlns:x="urn:schemas-microsoft-com:office:excel"`). They are named `formulario_{id}_{YYYY-MM-DD}_{HHMMSS}.xls` and contain one or more beneficiary records. Each record becomes one output file.

## Output Format

One `.xlsx` file is generated per beneficiary, following the layout of `Anexo I_PLANILHA PASSAGENS E DIÁRIAS 2026.xlsx`.

## Data Model

Each generated form can include **multiple beneficiaries**, each with:

**Beneficiary header** (one per person):
- CPF, full name, birth date, job title/function
- Bank: bank name, agency code + DV, account number + DV, savings flag
- Project description, event date range, event location
- Service types requested: `Passagens`, `Diárias`, `Terrestre`, `Aluguel de carro`
- A SHA-1 hash (submission/form identifier)

**Trip segments** (multiple per beneficiary):
- Origin city/state, destination city/state
- City type: `Capital` | `Interior` | `Cidade Média` | `Área remota`
- Departure date + time slot (`Manhã` / `Tarde` / `Noite`)
- Return date + time slot
- Transport type (optional, e.g. `Aéreo`, `Terrestre - Van`, `Fluvial - ...`)
- Observations (free text)

## Reference Files

- `Anexo I_PLANILHA PASSAGENS E DIÁRIAS 2026.xlsx` — official FIOTEC output template. Contains the exact layout, column headers, and boilerplate text that each generated file must match.
- `formulario_345_2026-05-08_170531.xls` — sample input file. Use it to understand the input structure and as a test fixture when parsing uploaded files.
