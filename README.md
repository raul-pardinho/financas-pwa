# Finanças — Raul & Letícia

PWA local-first para acompanhar **Budget, Forecast e Realizado**, com visões de **Competência** e **Caixa**.

## Arquitetura

- Fonte oficial: Google Sheets privado `Finanças Pessoais — Base Mestre`.
- Entrada no app: exportação da planilha como `.xlsx`.
- Persistência: IndexedDB no navegador.
- Hospedagem: GitHub Pages.
- PDF: gerado no navegador.
- Nenhum dado financeiro é armazenado neste repositório.

## Abas esperadas

Obrigatórias: `movimentacoes`, `planejamento`, `parcelamentos`, `controle`.

Opcionais usadas pela V1: `patrimonio`, `pendencias`, `cadastros`, `dicionario`.

Schema esperado: **1.2**.

## Publicação

Em **Settings → Pages**, configure **Deploy from a branch**, branch `main`, pasta `/(root)`.

## Privacidade

Nunca faça commit de XLSX, CSV, PDFs, extratos ou faturas. O `.gitignore` bloqueia esses formatos no uso normal do Git, mas não impede upload manual pelo navegador.
