# Extrator de links e metadados de PDFs 

Script em Python para ler pastas de PDFs de notícias, extrair automaticamente **título, fonte, data de publicação, URL/domínio, autor e tags**, e gerar uma planilha Excel (`.xlsx`) organizada com resumo por pasta.

## Requisitos

1. **Python 3.10+**
2. Bibliotecas Python:
   ```bash
   pip install pymupdf pandas openpyxl pytesseract pillow
   ```
3. **Tesseract OCR** (necessário para PDFs em formato de imagem):
   - **Windows:**
     ```powershell
     winget install UB-Mannheim.TesseractOCR
     ```
   - **Linux:**
     ```bash
     sudo apt install tesseract-ocr tesseract-ocr-por
     ```

## Como usar

Baixar a pasta principal (incluindo todas as subpastas):
```bash
python extrair_links.py "PASTA_COM_PDFS" saida.xlsx
```

Testar com um único arquivo PDF:
```bash
python extrair_links.py noticia.pdf teste.xlsx
```


## O que a planilha gerada contém

- **Aba `Notícias`:**
  - Código, pasta (e subníveis `pasta_1`, `pasta_2`, `pasta_3`), título, fonte, data/hora de publicação e data de impressão.
  - Link clicável da notícia (`url`) e origem da extração (`rodapé do PDF`, `rodapé (OCR)`, `link de compartilhamento`, `link da página`, `domínio do site` ou `imagem (OCR)`).
  - Link direto para abrir o arquivo PDF local (`caminho_pdf`).
  - Coluna `avisos` sinalizando URLs cortadas, duplicadas ou campos que precisam de conferência manual.
- **Aba `Resumo`:**
  - Total de PDFs, quantos têm URL, quantos ficaram sem URL e quantos possuem avisos em cada pasta.

