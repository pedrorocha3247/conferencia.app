# -*- coding: utf-8 -*-

import os
import sys
import re
import unicodedata
import io
import fitz  # PyMuPDF
import pandas as pd
from collections import OrderedDict
from flask import Flask, render_template, request, send_file, url_for
from openpyxl.styles import NamedStyle
import logging

# Configuração do Log (inalterada)
handler = logging.StreamHandler(sys.stdout)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
logger.handlers.clear()
logger.addHandler(handler)


# ==== Constantes e Mapeamentos (sem alterações) ====
# ... (código inalterado)

# ==== Funções de Normalização e Extração (sem alterações) ====
# ... (código inalterado)

# ==== Funções de Lógica e Classificação (sem alterações) ====
# ... (código inalterado)

# ==== Funções de Extração de Dados ====

def limpar_rotulo(lbl: str) -> str:
    lbl = re.sub(r"^TAMA\s*[-–—]\s*", "", lbl).strip()
    lbl = re.sub(r"\s+-\s+\d+/\d+$", "", lbl).strip()
    return lbl

# ... (outras funções inalteradas)

def extrair_parcelas(bloco: str):
    """
    Função corrigida para isolar a área de lançamentos e extrair as parcelas de forma mais precisa.
    """
    itens = OrderedDict()
    
    pos_lancamentos = bloco.find("Lançamentos")
    bloco_de_trabalho = bloco[pos_lancamentos + len("Lançamentos"):] if pos_lancamentos != -1 else bloco

    bloco_limpo_linhas = []
    for linha in bloco_de_trabalho.splitlines():
        match = re.search(r'\s{4,}(DÉBITOS DO MÊS ANTERIOR|ENCARGOS POR ATRASO|PAGAMENTO EFETUADO)', linha)
        if match:
            bloco_limpo_linhas.append(linha[:match.start()])
        else:
            bloco_limpo_linhas.append(linha)
    bloco_limpo = "\n".join(bloco_limpo_linhas)

    # Lógica 1: Captura parcelas e valores na mesma linha.
    for m in PADRAO_PARCELA_MESMA_LINHA.finditer(bloco_limpo):
        lbl = limpar_rotulo(m.group(1))
        # ALTERAÇÃO 1: Ignorar o campo DESCONTO
        if 'DESCONTO' in lbl.upper():
            continue
        val = to_float(m.group(2))
        if lbl and lbl not in itens and val is not None:
            itens[lbl] = val

    # Lógica 2: Captura parcelas cujo valor está na linha seguinte.
    linhas = bloco_limpo.splitlines()
    for i, linha in enumerate(linhas):
        linha_limpa = linha.strip()
        if not linha_limpa:
            continue
        
        # ALTERAÇÃO 1 (continuação): Ignorar o campo DESCONTO também nesta lógica
        if 'DESCONTO' in linha_limpa.upper():
            continue

        is_potential_label = (
            any(c.isalpha() for c in linha_limpa) and
            not any(h.upper() in linha_limpa.upper() for h in HEADERS) and
            limpar_rotulo(linha_limpa) not in itens and
            not PADRAO_PARCELA_MESMA_LINHA.match(linha_limpa)
        )

        if is_potential_label:
            j = i + 1
            while j < len(linhas) and not linhas[j].strip():
                j += 1
            if j < len(linhas):
                proxima_linha = linhas[j].strip()
                match_num = PADRAO_NUMERO_PURO.match(proxima_linha)
                if match_num:
                    lbl = limpar_rotulo(linha_limpa)
                    val = to_float(match_num.group(1))
                    if lbl and lbl not in itens and val is not None:
                        itens[lbl] = val
    return itens

# ==== Função de Processamento Principal (sem alterações na lógica interna) ====
# ... (código inalterado)

# --- Início da Aplicação Web Flask ---

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    # ... (código inicial da rota inalterado)
    if file and file.filename.lower().endswith('.pdf'):
        try:
            # ... (código de extração e processamento inalterado)

            df_todas, df_cov, df_div, df_erros = processar_pdf(texto_pdf, modo_separacao, emp_fixo)
            
            # ... (ordenação dos dataframes inalterada)

            output = io.BytesIO()
            
            # ALTERAÇÃO 2: Cria o dicionário para o Excel de forma ordenada
            # e adiciona a aba de erros condicionalmente no final.
            dfs_to_excel = {
                "Divergencias": df_div,
                "Cobertura_Analise": df_cov,
                "Todas_Parcelas_Extraidas": df_todas
            }
            if not df_erros.empty:
                dfs_to_excel["Lotes_Com_Erro"] = df_erros
            
            formatar_excel(output, dfs_to_excel)
            output.seek(0)
            
            # ... (criação do nome do arquivo e salvamento inalterados)

            # ... (criação do HTML e cálculo dos totais inalterados)
            total_erros = len(df_erros)
            
            # ... (log de conclusão inalterado)

            return render_template('results.html',
                                   # ... (outras variáveis inalteradas)
                                   total_erros=total_erros,
                                   erros_html=erros_html)
        
        except Exception:
            # ... (bloco de exceção inalterado)
    
    return "Formato de arquivo inválido. Por favor, envie um PDF.", 400

# ... (restante do código, como a função download_file, inalterado)
