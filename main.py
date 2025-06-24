# =======================================================
# SCRIPT: Processamento de Notas Fiscais em PDF com OCR
# =======================================================
#
# O QUE ESTE SCRIPT FAZ:
# - Lê um arquivo Excel com lista de notas fiscais
# - Procura cada nota PDF nas subpastas definidas
# - Copia e renomeia os PDFs para uma pasta destino
# - Extrai dados dos PDFs e compara com planilha
# - Usa OCR automaticamente se o PDF tiver imagem
# - Gera um log com divergências e erros
#
# REQUISITOS:
# - Python >= 3.10
# - Pacotes: pandas, openpyxl, pymupdf, fpdf, pytesseract, pdf2image, pillow
#
# =======================================================

import os
import sys
import shutil
import pandas as pd
import fitz  # PyMuPDF
import pytesseract
from pdf2image import convert_from_path
from datetime import datetime
from fpdf import FPDF

# Defina o caminho do executável do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ==========================
# CONFIGURAÇÕES (variáveis)
# ==========================

BASE_DIR = r'C:\ProjetoPython'

CAMINHO_EXCEL = os.path.join(BASE_DIR, 'arquivo.xlsx')
PASTA_ORIGEM_NOTAS = os.path.join(BASE_DIR, 'PastasDasNotas')
PASTA_DESTINO_NOTAS = os.path.join(BASE_DIR, 'PastaDestino')
CAMINHO_LOG = os.path.join(BASE_DIR, 'log_erros.txt')
NOME_ABA = 'Notas'

COLUNA_NUMERO_NOTA = 'NumeroNota'
COLUNA_CNPJ = 'CNPJ'
COLUNA_VALOR_TOTAL = 'ValorTotal'
COLUNA_DESCRICAO = 'Descricao'

# ==========================
# FUNÇÕES
# ==========================

def validar_ambiente():
    if sys.version_info < (3, 10):
        print('Python 3.10 ou superior é necessário.')
        sys.exit(1)

def logar(mensagem):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(CAMINHO_LOG, 'a', encoding='utf-8') as log:
        log.write(f'[{timestamp}] {mensagem}\n')

def verificar_caminhos():
    if not os.path.isdir(BASE_DIR):
        os.makedirs(BASE_DIR)
    if not os.path.isfile(CAMINHO_EXCEL):
        criar_planilha_exemplo()
    if not os.path.isdir(PASTA_ORIGEM_NOTAS):
        os.makedirs(PASTA_ORIGEM_NOTAS)
        criar_pdfs_exemplo()
    if not os.path.exists(PASTA_DESTINO_NOTAS):
        os.makedirs(PASTA_DESTINO_NOTAS)

def criar_planilha_exemplo():
    dados = {
        COLUNA_NUMERO_NOTA: ['12345', '67890', '11111'],
        COLUNA_CNPJ: ['12.345.678/0001-99', '98.765.432/0001-11', '11.222.333/0001-44'],
        COLUNA_VALOR_TOTAL: ['1000.00', '2500.50', '750.00'],
        COLUNA_DESCRICAO: ['Produto A', 'Produto B', 'Produto C']
    }
    df = pd.DataFrame(dados)
    df.to_excel(CAMINHO_EXCEL, sheet_name=NOME_ABA, index=False, engine='openpyxl')

def criar_pdfs_exemplo():
    exemplos = [
        ('12345', '12.345.678/0001-99', '1000.00', 'Produto A'),
        ('67890', '98.765.432/0001-11', '2500.50', 'Produto B'),
        ('11111', '11.222.333/0001-44', '750.00', 'Produto C')
    ]
    for nota, cnpj, valor, desc in exemplos:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt=f"NumeroNota: {nota}", ln=True)
        pdf.cell(200, 10, txt=f"CNPJ: {cnpj}", ln=True)
        pdf.cell(200, 10, txt=f"ValorTotal: {valor}", ln=True)
        pdf.cell(200, 10, txt=f"Descricao: {desc}", ln=True)
        caminho_pdf = os.path.join(PASTA_ORIGEM_NOTAS, f'Nota_{nota}.pdf')
        pdf.output(caminho_pdf)

def extrair_texto_pdf(caminho_pdf):
    doc = fitz.open(caminho_pdf)
    texto_completo = ""
    for page in doc:
        texto_pagina = page.get_text()
        if texto_pagina.strip():
            texto_completo += texto_pagina
        else:
            imagens = convert_from_path(caminho_pdf, first_page=page.number+1, last_page=page.number+1)
            for imagem in imagens:
                texto_ocr = pytesseract.image_to_string(imagem, lang='por')
                texto_completo += texto_ocr
    doc.close()
    if not texto_completo.strip():
        raise ValueError('PDF vazio ou ilegível')
    return texto_completo

def procurar_nota(numero_nota):
    for raiz, _, arquivos in os.walk(PASTA_ORIGEM_NOTAS):
        for arquivo in arquivos:
            if str(numero_nota) in arquivo and arquivo.lower().endswith('.pdf'):
                return os.path.join(raiz, arquivo)
    return None

def comparar_campos(numero_nota, esperado, extraido, campo):
    if esperado not in extraido:
        logar(f'Divergência em {campo} para nota {numero_nota}: esperado "{esperado}" não encontrado no PDF')

# ==========================
# EXECUÇÃO PRINCIPAL
# ==========================

def main():
    validar_ambiente()
    verificar_caminhos()

    try:
        df = pd.read_excel(CAMINHO_EXCEL, sheet_name=NOME_ABA, engine='openpyxl')
    except Exception as e:
        print(f'Erro ao ler planilha: {e}')
        logar(f'Erro ao ler planilha: {e}')
        return

    if df.empty:
        logar('Planilha está vazia!')
        print('Planilha está vazia!')
        return

    total_notas = len(df)
    logar(f'Início do processamento de {total_notas} notas')

    for index, row in df.iterrows():
        numero_nota = str(row[COLUNA_NUMERO_NOTA])
        cnpj = str(row[COLUNA_CNPJ])
        valor_total = str(row[COLUNA_VALOR_TOTAL])
        descricao = str(row[COLUNA_DESCRICAO])

        caminho_pdf_origem = procurar_nota(numero_nota)

        if not caminho_pdf_origem:
            logar(f'Nota não encontrada: {numero_nota}')
            continue

        novo_nome = f'Nota_{numero_nota}.pdf'
        caminho_pdf_destino = os.path.join(PASTA_DESTINO_NOTAS, novo_nome)

        try:
            shutil.copy2(caminho_pdf_origem, caminho_pdf_destino)
            logar(f'Nota {numero_nota} copiada para {caminho_pdf_destino}')
        except Exception as e:
            logar(f'Erro ao copiar nota {numero_nota}: {e}')
            continue

        try:
            texto_pdf = extrair_texto_pdf(caminho_pdf_destino)
            comparar_campos(numero_nota, numero_nota, texto_pdf, 'Número da Nota')
            comparar_campos(numero_nota, cnpj, texto_pdf, 'CNPJ')
            comparar_campos(numero_nota, valor_total, texto_pdf, 'Valor Total')
            comparar_campos(numero_nota, descricao, texto_pdf, 'Descrição')
        except Exception as e:
            logar(f'Erro ao processar PDF da nota {numero_nota}: {e}')
            continue

    logar('Fim do processamento')

if __name__ == '__main__':
    main()
