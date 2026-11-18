# Extrator de Links e Metadados de PDFs

Script em Python para ler pastas de PDFs de notícias, extrair automaticamente **título, fonte, data/hora de publicação e URL**, e gerar uma planilha Excel (`.xlsx`) organizada com resumo por pasta.

## Requisitos

1. **Python 3.10+**
2. **Bibliotecas Python:**
   ```bash
   pip install pymupdf pandas openpyxl pytesseract pillow
   ```
3. **Tesseract OCR** (necessário para PDFs em formato de imagem):
   - **Windows:**
     ```powershell
     winget install UB-Mannheim.TesseractOCR
     ```
     *(Marque o idioma **Portuguese** durante a instalação)*
   - **Linux:**
     ```bash
     sudo apt install tesseract-ocr tesseract-ocr-por
     ```
4. **Ollama (SLM Local)** — **Recomendado** para desempatar links de matérias sem rodapé e preencher metadados faltantes:
   - Instale o [Ollama](https://ollama.com/) (ou via terminal no Windows: `winget install Ollama.Ollama`).
   - Baixe um modelo leve (ex.: `qwen2.5:3b` ou `llama3.2:3b`):
     ```bash
     ollama pull qwen2.5:3b
     ```
   - Certifique-se de que o Ollama está rodando em segundo plano (`http://localhost:11434`) antes de executar o script.

## Como usar

**Com o SLM local ativado (recomendado):**
```bash
python extrair_links.py "PASTA_COM_PDFS" saida.xlsx --modelo-slm qwen2.5:3b
```

**Sem SLM local (apenas regras e OCR):**
```bash
python extrair_links.py "PASTA_COM_PDFS" saida.xlsx
```

**Testar com um único arquivo PDF:**
```bash
python extrair_links.py noticia.pdf teste.xlsx --modelo-slm qwen2.5:3b
```

## O que a planilha gerada contém

- **Aba `Notícias`:**
  - `codigo`, `pasta` (e subníveis `pasta_1`, `pasta_2`, `pasta_3`), `titulo`, `fonte`, `data_publicacao` e `hora_publicacao`.
  - Link clicável da notícia (`url`) e origem da extração (`url_origem`: `rodapé do PDF`, `rodapé + título (...)`, `rodapé (OCR)`, `link de compartilhamento`, `link da página`, `link da página (SLM)`, `imagem (OCR)` ou `home do site`).
- **Aba `Resumo`:**
  - Total de PDFs, quantos têm URL (`com_url`) e quantos ficaram sem URL (`sem_url`) em cada pasta.
