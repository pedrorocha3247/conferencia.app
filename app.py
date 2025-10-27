# -*- coding: utf-8 -*-
import os
import sys
import re
import unicodedata
import io
import fitz  # PyMuPDF
import pandas as pd
from collections import OrderedDict
from flask import Flask, request, send_file, url_for, make_response
import traceback
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import NamedStyle, Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy
import zipfile
import time # Importado para logs

# ==== Constantes e Mapeamentos ====
DASHES = dict.fromkeys(map(ord, "\u2010\u2011\u2012\u2013\u2014\u2015\u2212"), "-")
HEADERS = (
    "Remessa para Conferência", "Página", "Banco", "IMOBILIARIOS", "Débitos do Mês",
    "Vencimento", "Lançamentos", "Programação", "Carta", "DÉBITOS", "ENCARGOS",
    "PAGAMENTO", "TOTAL", "Limite p/", "TOTAL A PAGAR", "PAGAMENTO EFETUADO", "DESCONTO"
)
PADRAO_LOTE = re.compile(r"\b(\d{2,4}\.([A-Z0-9\u0399\u039A]{2})\.\d{1,4})\b")
PADRAO_PARCELA_MESMA_LINHA = re.compile(
    r"^(?!(?:DÉBITOS|ENCARGOS|DESCONTO|PAGAMENTO|TOTAL|Limite p/))\s*"
    r"([A-Za-zÀ-ú][A-Za-zÀ-ú\s\.\-\/\d]+?)\s+([\d.,]+)"
    r"(?=\s{2,}|\t|$)", re.MULTILINE
)
PADRAO_NUMERO_PURO = re.compile(r"^\s*([\d\.,]+)\s*$")
CODIGO_EMP_MAP = {
    '04': 'RSCI', '05': 'RSCIV', '06': 'RSCII', '07': 'RSCV', '08': 'RSCIII',
    '09': 'IATE', '10': 'MARINA', '11': 'NVI', '12': 'NVII',
    '13': 'SBRRI', '14': 'SBRRII', '15': 'SBRRIII'
}
EMP_MAP = {
    "NVI": {"Melhoramentos": 205.61, "Fundo de Transporte": 9.00},
    "NVII": {"Melhoramentos": 245.47, "Fundo de Transporte": 9.00},
    "RSCI": {"Melhoramentos": 250.42, "Fundo de Transporte": 9.00},
    "RSCII": {"Melhoramentos": 240.29, "Fundo de Transporte": 9.00},
    "RSCIII": {"Melhoramentos": 281.44, "Fundo de Transporte": 9.00},
    "RSCIV": {"Melhoramentos": 303.60, "Fundo de Transporte": 9.00},
    "IATE": {"Melhoramentos": 240.00, "Fundo de Transporte": 9.00},
    "MARINA": {"Melhoramentos": 240.00, "Fundo de Transporte": 9.00},
    "SBRRI": {"Melhoramentos": 245.47, "Fundo de Transporte": 13.00},
    "SBRRII": {"Melhoramentos": 245.47, "Fundo de Transporte": 13.00},
    "SBRRIII": {"Melhoramentos": 245.47, "Fundo de Transporte": 13.00},
    "RSCV": {"Melhoramentos": 280.00, "Fundo de Transporte": 9.00},
}
BASE_FIXOS = {
    "Taxa de Conservação": [434.11],
    "Contrib. Social SLIM": [103.00, 309.00],
    "Contribuição ABRASMA - Bronze": [20.00],
    "Contribuição ABRASMA - Prata": [40.00],
    "Contribuição ABRASMA - Ouro": [60.00],
}
# === LÓGICA CCB/REALIZA REIMPLEMENTADA ===
BASE_FIXOS_CCB = {
    "Alienação Fiduciária CCB": [],
    "Financiamento Realiza CCB": [],
    "Encargos Não Pagos CCB": [],
    "Débito por pagamento a menor CCB": [],
    "Crédito por pagamento a maior CCB": [],
    "Negociação Alienação CCB": []
}
# =======================================

app = Flask(__name__)
# Define UPLOAD_FOLDER como um caminho absoluto relativo à raiz do app
UPLOAD_FOLDER_PATH = os.path.join(app.root_path, 'uploads')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER_PATH
# Cria o diretório usando o caminho absoluto
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
print(f"Pasta de Upload configurada em: {app.config['UPLOAD_FOLDER']}") # Log para confirmar

def manual_render_template(template_name, status_code=200, **kwargs):
    template_path = os.path.join(app.root_path, 'templates', template_name)
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        for key, value in kwargs.items():
            placeholder = f"__{key.upper()}__"
            if isinstance(value, str) and ('{' in value and '}' in value):
                 html_content = html_content.replace(f'"{placeholder}"', value)
            else:
                html_content = html_content.replace(placeholder, str(value))

        response = make_response(html_content)
        response.headers['Content-Type'] = 'text/html'
        return response, status_code
    except Exception as e:
        print(f"ERRO CRÍTICO AO RENDERIZAR MANUALMENTE '{template_name}': {e}")
        return f"<h1>Erro 500: Falha Crítica ao Carregar Template</h1><p>O arquivo {template_name} não pôde ser lido. Erro: {e}</p>", 500

def normalizar_texto(s: str) -> str:
    s = s.translate(DASHES).replace("\u00A0", " ")
    s = "".join(ch for ch in s if ch not in "\u200B\u200C\u200D\uFEFF")
    s = unicodedata.normalize("NFKC", s)
    return s

def extrair_texto_pdf(stream_pdf) -> str:
    try:
        doc = fitz.open(stream=stream_pdf, filetype="pdf")
        texto = "\n".join(p.get_text("text", sort=True) for p in doc)
        doc.close()
        return normalizar_texto(texto)
    except Exception as e:
        print(f"Erro ao ler o stream do PDF: {e}")
        return ""

def to_float(s: str):
    try:
        return float(s.replace(".", "").replace(",", ".").strip())
    except (ValueError, TypeError):
        return None

# === LÓGICA CCB/REALIZA REIMPLEMENTADA ===
def fixos_do_emp(emp: str, modo_separacao: str):
    if modo_separacao == 'boleto':
        if emp not in EMP_MAP:
            return BASE_FIXOS
        f = dict(BASE_FIXOS)
        if EMP_MAP.get(emp) and EMP_MAP[emp].get("Melhoramentos") is not None:
            f["Melhoramentos"] = [float(EMP_MAP[emp]["Melhoramentos"])]
        if EMP_MAP.get(emp) and EMP_MAP[emp].get("Fundo de Transporte") is not None:
            f["Fundo de Transporte"] = [float(EMP_MAP[emp]["Fundo de Transporte"])]
        return f
    elif modo_separacao == 'debito_credito':
        # Retorna BASE_FIXOS para débito/crédito (ou ajuste se necessário)
        return BASE_FIXOS
    elif modo_separacao == 'ccb_realiza':
        # Retorna o novo conjunto de parcelas para CCB/Realiza
        return BASE_FIXOS_CCB
    else:
        # Modo desconhecido, retorna vazio para evitar erros
        print(f"[AVISO] Modo de separação desconhecido: {modo_separacao}")
        return {}
# =======================================

def detectar_emp_por_nome_arquivo(path: str):
    nome = os.path.splitext(os.path.basename(path))[0].upper()
    for k in EMP_MAP.keys():
        if nome.endswith("_" + k) or nome.endswith(k):
            return k
    if "SBRR" in nome:
        return "SBRR"
    return None

def detectar_emp_por_lote(lote: str):
    if not lote or "." not in lote:
        return "NAO_CLASSIFICADO"
    prefixo = lote.split('.')[0]
    return CODIGO_EMP_MAP.get(prefixo, "NAO_CLASSIFICADO")

def limpar_rotulo(lbl: str) -> str:
    lbl = re.sub(r"^TAMA\s*[-–—]\s*", "", lbl).strip()
    lbl = re.sub(r"\s+-\s+\d+/\d+$", "", lbl).strip()
    return lbl

def fatiar_blocos(texto: str):
    texto_processado = PADRAO_LOTE.sub(r"\n\1", texto)
    ms = list(PADRAO_LOTE.finditer(texto_processado))
    blocos = []
    for i, m in enumerate(ms):
        ini = m.start()
        fim = ms[i+1].start() if i+1 < len(ms) else len(texto_processado)
        blocos.append((m.group(1), texto_processado[ini:fim]))
    return blocos

def tentar_nome_cliente(bloco: str) -> str:
    linhas = bloco.split('\n')
    if not linhas: return "Nome não localizado"
    
    linhas_para_buscar = [linhas[0]] + linhas[1:5]
    for linha in linhas_para_buscar:
        linha_sem_lote = PADRAO_LOTE.sub('', linha).strip()
        if not linha_sem_lote: continue
        is_valid_name = (
            len(linha_sem_lote) > 5 and ' ' in linha_sem_lote and
            sum(c.isalpha() for c in linha_sem_lote.replace(" ", "")) / len(linha_sem_lote.replace(" ", "")) > 0.7 and
            not any(h.upper() in linha_sem_lote.upper() for h in HEADERS) and
            not re.search(r'\d{2}/\d{2}/\d{4}', linha_sem_lote) and
            not linha_sem_lote.upper().startswith(("TOTAL", "BANCO", "03-"))
        )
        if is_valid_name:
            return linha_sem_lote
    return "Nome não localizado"

def extrair_parcelas(bloco: str):
    itens = OrderedDict()
    pos_lancamentos = bloco.find("Lançamentos")
    bloco_de_trabalho = bloco[pos_lancamentos:] if pos_lancamentos != -1 else bloco

    bloco_limpo_linhas = []
    for linha in bloco_de_trabalho.splitlines():
        match = re.search(r'\s{4,}(DÉBITOS DO MÊS ANTERIOR|ENCARGOS POR ATRASO|PAGAMENTO EFETUADO)', linha)
        linha_processada = linha[:match.start()] if match else linha
        if not any(h in linha_processada for h in ["Lançamentos", "Débitos do Mês"]):
            bloco_limpo_linhas.append(linha_processada)
    bloco_limpo = "\n".join(bloco_limpo_linhas)

    for m in PADRAO_PARCELA_MESMA_LINHA.finditer(bloco_limpo):
        lbl = limpar_rotulo(m.group(1))
        val = to_float(m.group(2))
        if lbl and lbl not in itens and val is not None:
            itens[lbl] = val

    linhas = bloco_limpo.splitlines()
    for i, linha in enumerate(linhas):
        linha_limpa = linha.strip()
        if not linha_limpa: continue
        is_potential_label = (
            any(c.isalpha() for c in linha_limpa) and
            not any(h.upper() in linha_limpa.upper() for h in HEADERS) and
            limpar_rotulo(linha_limpa) not in itens and
            not PADRAO_PARCELA_MESMA_LINHA.match(linha_limpa)
        )
        if is_potential_label:
            j = i + 1
            while j < len(linhas) and not linhas[j].strip(): j += 1
            if j < len(linhas):
                match_num = PADRAO_NUMERO_PURO.match(linhas[j].strip())
                if match_num:
                    lbl = limpar_rotulo(linha_limpa)
                    val = to_float(match_num.group(1))
                    if lbl and lbl not in itens and val is not None: itens[lbl] = val
    return itens

def processar_pdf_validacao(texto_pdf: str, modo_separacao: str, emp_fixo_boleto: str = None):
    blocos = fatiar_blocos(texto_pdf)
    if not blocos: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    linhas_todas, linhas_cov, linhas_div = [], [], []
    for lote, bloco in blocos:
        if modo_separacao == 'boleto':
            emp_atual = detectar_emp_por_lote(lote) if emp_fixo_boleto == "SBRR" else emp_fixo_boleto
        else:
            # Para debito_credito e ccb_realiza, sempre tenta detectar pelo lote
            emp_atual = detectar_emp_por_lote(lote)
        
        cliente = tentar_nome_cliente(bloco)
        itens = extrair_parcelas(bloco)

        # === LÓGICA CCB/REALIZA REIMPLEMENTADA ===
        # Passa o modo_separacao para obter o conjunto correto de parcelas
        VALORES_CORRETOS = fixos_do_emp(emp_atual, modo_separacao)
        # =======================================

        for rot, val in itens.items():
            linhas_todas.append({"Empreendimento": emp_atual, "Lote": lote, "Cliente": cliente, "Parcela": rot, "Valor": val})
        
        cov = {"Empreendimento": emp_atual, "Lote": lote, "Cliente": cliente}
        for k in VALORES_CORRETOS.keys(): cov[k] = None
        for rot, val in itens.items():
            if rot in VALORES_CORRETOS: cov[rot] = val
        
        vistos = [k for k in VALORES_CORRETOS if cov[k] is not None]
        cov["QtdParc_Alvo"] = len(vistos)
        cov["Parc_Alvo"] = ", ".join(vistos)
        linhas_cov.append(cov)

        # A lógica de divergência só se aplica se VALORES_CORRETOS tiver valores definidos (não vazios)
        if modo_separacao != 'ccb_realiza':
            for rot in vistos:
                val = cov[rot]
                if val is None: continue
                permitidos = VALORES_CORRETOS.get(rot, [])
                # Verifica se a lista de permitidos não está vazia antes de comparar
                if permitidos and all(abs(val - v) > 1e-6 for v in permitidos):
                    linhas_div.append({
                        "Empreendimento": emp_atual, "Lote": lote, "Cliente": cliente,
                        "Parcela": rot, "Valor no Documento": float(val),
                        "Valor Correto": " ou ".join(f"{v:.2f}" for v in permitidos)
                    })
        # Para CCB/Realiza, a validação de valor não se aplica da mesma forma (listas vazias em BASE_FIXOS_CCB)
        # então não adicionamos nada a linhas_div nesse caso, a menos que uma regra específica seja definida.

    df_todas = pd.DataFrame(linhas_todas)
    df_cov = pd.DataFrame(linhas_cov)
    df_div = pd.DataFrame(linhas_div)
    
    return df_todas, df_cov, df_div


def processar_comparativo(texto_anterior, texto_atual, modo_separacao, emp_fixo_boleto):
    df_todas_ant_raw, _, _ = processar_pdf_validacao(texto_anterior, modo_separacao, emp_fixo_boleto)
    df_todas_atu_raw, _, _ = processar_pdf_validacao(texto_atual, modo_separacao, emp_fixo_boleto)
    
    df_totais_ant = df_todas_ant_raw[df_todas_ant_raw['Parcela'].str.strip().str.upper() == 'TOTAL A PAGAR'].copy()
    df_totais_ant = df_totais_ant[['Empreendimento', 'Lote', 'Cliente', 'Valor']].rename(columns={'Valor': 'Total Anterior'})

    df_totais_atu = df_todas_atu_raw[df_todas_atu_raw['Parcela'].str.strip().str.upper() == 'TOTAL A PAGAR'].copy()
    df_totais_atu = df_totais_atu[['Empreendimento', 'Lote', 'Cliente', 'Valor']].rename(columns={'Valor': 'Total Atual'})

    parcelas_para_remover = ['TOTAL A PAGAR', 'DESCONTO', 'DÉBITOS DO MÊS']
    df_todas_ant = df_todas_ant_raw[~df_todas_ant_raw['Parcela'].str.strip().str.upper().isin(parcelas_para_remover)].copy()
    df_todas_atu = df_todas_atu_raw[~df_todas_atu_raw['Parcela'].str.strip().str.upper().isin(parcelas_para_remover)].copy()
    
    df_todas_ant = df_todas_ant[~df_todas_ant['Parcela'].str.strip().str.upper().str.startswith('TOTAL BANCO')].copy()
    df_todas_atu = df_todas_atu[~df_todas_atu['Parcela'].str.strip().str.upper().str.startswith('TOTAL BANCO')].copy()

    df_todas_ant.rename(columns={'Valor': 'Valor Anterior'}, inplace=True)
    df_todas_atu.rename(columns={'Valor': 'Valor Atual'}, inplace=True)

    df_comp = pd.merge(df_todas_ant, df_todas_atu, on=['Empreendimento', 'Lote', 'Cliente', 'Parcela'], how='outer')
    lotes_ant = df_todas_ant_raw[['Empreendimento', 'Lote', 'Cliente']].drop_duplicates()
    lotes_atu = df_todas_atu_raw[['Empreendimento', 'Lote', 'Cliente']].drop_duplicates()
    lotes_merged = pd.merge(lotes_ant, lotes_atu, on=['Empreendimento', 'Lote', 'Cliente'], how='outer', indicator=True)
    
    df_adicionados_base = lotes_merged[lotes_merged['_merge'] == 'right_only'][['Empreendimento', 'Lote', 'Cliente']]
    df_removidos_base = lotes_merged[lotes_merged['_merge'] == 'left_only'][['Empreendimento', 'Lote', 'Cliente']]
    
    df_adicionados = pd.merge(df_adicionados_base, df_totais_atu, on=['Empreendimento', 'Lote', 'Cliente'], how='left')
    df_removidos = pd.merge(df_removidos_base, df_totais_ant, on=['Empreendimento', 'Lote', 'Cliente'], how='left')

    df_divergencias = df_comp[(pd.notna(df_comp['Valor Anterior'])) & (pd.notna(df_comp['Valor Atual'])) & (abs(df_comp['Valor Anterior'] - df_comp['Valor Atual']) > 1e-6)].copy()
    df_divergencias['Diferença'] = df_divergencias['Valor Atual'] - df_divergencias['Valor Anterior']
    df_parcelas_novas = df_comp[df_comp['Valor Anterior'].isna() & pd.notna(df_comp['Valor Atual'])][['Empreendimento', 'Lote', 'Cliente', 'Parcela', 'Valor Atual']]
    df_parcelas_removidas = df_comp[df_comp['Valor Atual'].isna() & pd.notna(df_comp['Valor Anterior'])][['Empreendimento', 'Lote', 'Cliente', 'Parcela', 'Valor Anterior']]
    
    total_adicionados_valor = df_adicionados['Total Atual'].sum() if 'Total Atual' in df_adicionados.columns else 0
    total_removidos_valor = df_removidos['Total Anterior'].sum() if 'Total Anterior' in df_removidos.columns else 0
    total_divergencias_valor = df_divergencias['Diferença'].sum() if 'Diferença' in df_divergencias.columns else 0
    total_mes_anterior_valor = df_totais_ant['Total Anterior'].sum() if 'Total Anterior' in df_totais_ant.columns else 0
    total_mes_atual_valor = df_totais_atu['Total Atual'].sum() if 'Total Atual' in df_totais_atu.columns else 0


    resumo_financeiro_data = {
        ' ': ['Lotes Mês Anterior', 'Lotes Mês Atual', 'Lotes Adicionados', 'Lotes Removidos', 'Parcelas com Valor Alterado'],
        'LOTES': [len(lotes_ant), len(lotes_atu), len(df_adicionados), len(df_removidos), df_divergencias['Lote'].nunique() if not df_divergencias.empty else 0], # Conta lotes únicos com divergência
        'TOTAIS': [total_mes_anterior_valor, total_mes_atual_valor, total_adicionados_valor, total_removidos_valor, total_divergencias_valor]
    }
    df_resumo_completo = pd.DataFrame(resumo_financeiro_data)
    
    # === CORREÇÃO DO ERRO DE DIGITAÇÃO AQUI ===
    return df_resumo_completo, df_adicionados, df_removidos, df_divergencias, df_parcelas_novas, df_parcelas_removidas
    # Estava: df_parc_novas, df_parc_removidas
    # =========================================

def formatar_excel(output_stream, dfs: dict):
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
        for sheet_name, df in dfs.items():
            # Garante que mesmo DataFrames vazios sejam tratados
            if df is None:
                continue # Pula se o DataFrame for None
            if isinstance(df, pd.DataFrame) and not df.empty:
                 if sheet_name == "Resumo":
                    df.to_excel(writer, index=False, sheet_name=sheet_name)
                 else:
                    df.to_excel(writer, index=False, sheet_name=sheet_name)
            elif isinstance(df, pd.DataFrame) and df.empty:
                 # Cria uma planilha vazia se o DataFrame estiver vazio mas não for None
                 pd.DataFrame().to_excel(writer, sheet_name=sheet_name, index=False)


        number_style = NamedStyle(name='br_number_style', number_format='#,##0.00')
        integer_style = NamedStyle(name='br_integer_style', number_format='0')

        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            worksheet.sheet_view.showGridLines = False
            
            # Aplica formatação apenas se a planilha tiver dados
            if worksheet.max_row > 1 or worksheet.max_column > 1 : # Verifica se há mais que o cabeçalho
                for column_cells in worksheet.columns:
                    max_length = 0
                    column = column_cells[0].column_letter
                    is_first_row = True
                    for cell in column_cells:
                        if is_first_row: # Pula a formatação do cabeçalho
                            is_first_row = False
                            if cell.value: # Calcula largura baseado no cabeçalho também
                                max_length = max(max_length, len(str(cell.value)))
                            continue

                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))

                        # Aplica estilo numérico apenas se for float ou int (exceto coluna B no Resumo)
                        if isinstance(cell.value, float):
                             cell.style = number_style
                        elif isinstance(cell.value, int):
                             if sheet_name == 'Resumo' and column == 'B': # Coluna LOTES no Resumo
                                 cell.style = integer_style
                             elif column != 'B': # Outras colunas int (se houver) formatar como número
                                 cell.style = number_style # Ou manter como integer_style se preferir

                    adjusted_width = (max_length + 2)
                    worksheet.column_dimensions[column].width = adjusted_width
            else:
                 # Define uma largura mínima para planilhas vazias, se desejar
                 for col_idx in range(1, 5): # Ex: define largura para as 4 primeiras colunas
                      worksheet.column_dimensions[get_column_letter(col_idx)].width = 15

    return output_stream


def normalizar_valor_repasse(valor):
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return round(float(valor), 2)
    s = str(valor).strip().replace("R$", "").replace(" ", "").replace("\xa0", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return round(float(s), 2)
    except:
        return 0.0

def copiar_formatacao(origem, destino):
    if origem and hasattr(origem, 'has_style') and origem.has_style:
        destino.font = copy(origem.font)
        destino.border = copy(origem.border)
        destino.fill = copy(origem.fill)
        destino.number_format = copy(origem.number_format)
        destino.protection = copy(origem.protection)
        destino.alignment = copy(origem.alignment)

def achar_coluna(sheet, nome_coluna):
    if sheet.max_row == 0: # Planilha vazia
        return None
    for cell in sheet[1]: # Itera sobre a primeira linha (cabeçalho)
        if cell.value and str(cell.value).strip().lower() == nome_coluna.lower():
            return cell.column
    return None


def criar_planilha_saida(linhas, ws_diario, incluir_status=False):
    wb_out = Workbook()
    ws_out = wb_out.active

    # Verifica se ws_diario tem pelo menos uma linha para copiar cabeçalho
    if ws_diario.max_row > 0:
        num_cols_header = ws_diario.max_column
        for i, cell in enumerate(ws_diario[1], 1):
            if cell:
                novo = ws_out.cell(row=1, column=i, value=cell.value)
                copiar_formatacao(cell, novo)
                # Tenta copiar a largura da coluna, se existir
                col_letter = get_column_letter(i)
                if col_letter in ws_diario.column_dimensions:
                     ws_out.column_dimensions[col_letter].width = ws_diario.column_dimensions[col_letter].width
            else:
                 ws_out.cell(row=1, column=i, value=None) # Cria célula vazia se original for None
    else:
        num_cols_header = 0 # Define 0 se não há cabeçalho

    col_status = 0
    if incluir_status:
        # Adiciona coluna Status mesmo se não houver cabeçalho original
        col_status = num_cols_header + 1
        ws_out.cell(row=1, column=col_status, value="Status")
        ws_out.column_dimensions[get_column_letter(col_status)].width = 25 # Define uma largura


    linha_out = 2
    for linha_info in linhas:
        linha, status = linha_info

        if linha is None:
            # Escreve apenas o status se a linha for None (item presente só no sistema)
            if incluir_status and col_status > 0:
                ws_out.cell(row=linha_out, column=col_status, value=status)
            linha_out += 1
            continue

        # Processa a linha (seja tupla de valores ou tupla de células)
        for i, cell_data in enumerate(linha, 1):
             try:
                 # Lida com Célula ou Valor diretamente
                 valor = cell_data.value if hasattr(cell_data, "value") else cell_data
                 novo = ws_out.cell(row=linha_out, column=i, value=valor)
                 # Copia formatação apenas se for uma Célula
                 if hasattr(cell_data, "value"):
                     copiar_formatacao(cell_data, novo)
             except Exception as e:
                  print(f"[Aviso] Erro ao processar célula {i} da linha {linha_out}: {e}. Valor: {cell_data}")
                  ws_out.cell(row=linha_out, column=i, value=f"ERRO: {e}") # Marca erro na planilha


        # Adiciona o status na coluna correspondente
        if incluir_status and col_status > 0:
             ws_out.cell(row=linha_out, column=col_status, value=status)

        linha_out += 1

    # Adiciona contagem total no final da planilha de divergentes
    if incluir_status:
         # Garante que a célula exista antes de escrever
         total_cell = ws_out.cell(row=linha_out + 1, column=1)
         total_cell.value = f"Total divergentes: {len(linhas)}"
         # Aplica algum estilo básico, se desejado
         total_cell.font = Font(bold=True)


    stream_out = io.BytesIO()
    wb_out.save(stream_out)
    stream_out.seek(0)
    return stream_out


def processar_repasse(diario_stream, sistema_stream):
    print("📘 [LOG] Início de processar_repasse")
    start_time = time.time()
    
    print("📘 [LOG] Carregando workbook 'Diário'...")
    wb_diario = load_workbook(diario_stream, data_only=True)
    ws_diario = wb_diario.worksheets[0]
    print(f"📗 [LOG] 'Diário' carregado ({ws_diario.max_row} linhas). Tempo: {time.time() - start_time:.2f}s")


    print("📘 [LOG] Carregando workbook 'Sistema'...")
    wb_sistema = load_workbook(sistema_stream, data_only=True)
    ws_sistema = wb_sistema.worksheets[0]
    print(f"📗 [LOG] 'Sistema' carregado ({ws_sistema.max_row} linhas). Tempo: {time.time() - start_time:.2f}s")


    print("📘 [LOG] Achando colunas...")
    col_eq_diario = achar_coluna(ws_diario, "EQL")
    col_parcela_diario = achar_coluna(ws_diario, "Parcela")
    col_principal_diario = 4 # Assumindo coluna D
    col_corrmonet_diario = 9 # Assumindo coluna I

    col_eq_sistema = achar_coluna(ws_sistema, "EQL")
    col_parcela_sistema = achar_coluna(ws_sistema, "Parcela")
    col_valor_sistema = achar_coluna(ws_sistema, "Valor")
    print(f"📗 [LOG] Colunas encontradas: Diário(EQL:{col_eq_diario}, Parc:{col_parcela_diario}), Sistema(EQL:{col_eq_sistema}, Parc:{col_parcela_sistema}, Val:{col_valor_sistema})")


    # Validação crítica das colunas encontradas
    missing_cols = []
    if not col_eq_diario: missing_cols.append("EQL (Diário)")
    if not col_parcela_diario: missing_cols.append("Parcela (Diário)")
    if not col_eq_sistema: missing_cols.append("EQL (Sistema)")
    if not col_parcela_sistema: missing_cols.append("Parcela (Sistema)")
    if not col_valor_sistema: missing_cols.append("Valor (Sistema)")

    if missing_cols:
         error_msg = f"Não foi possível encontrar as seguintes colunas obrigatórias: {', '.join(missing_cols)}. Verifique os nomes nos cabeçalhos das planilhas."
         print(f"📕 [ERRO] {error_msg}")
         raise ValueError(error_msg)


    print("📘 [LOG] Início do Loop 1: Processando 'Diário' (values_only)...")
    valores_diario = {}
    contagem_diario = {}
    linhas_diario_count = 0
    for i, row in enumerate(ws_diario.iter_rows(min_row=2, values_only=True)):
        linhas_diario_count += 1
        if i % 500 == 0 and i > 0:
            print(f"➡️  [LOG] Processando linha {i+2} do Diário (Loop 1)...")

        # Tratamento de erro mais robusto para acesso às colunas
        eql = str(row[col_eq_diario - 1]).strip() if col_eq_diario <= len(row) and row[col_eq_diario - 1] else ""
        parcela = str(row[col_parcela_diario - 1]).strip() if col_parcela_diario <= len(row) and row[col_parcela_diario - 1] else ""
        principal = normalizar_valor_repasse(row[col_principal_diario - 1]) if col_principal_diario <= len(row) else 0.0
        correcao = normalizar_valor_repasse(row[col_corrmonet_diario - 1]) if col_corrmonet_diario <= len(row) else 0.0
        total = round(principal + correcao, 2)

        if eql and parcela:
            chave_completa = (eql, parcela, principal, correcao)
            chave_simples = (eql, parcela)
            contagem_diario[chave_completa] = contagem_diario.get(chave_completa, 0) + 1
            if chave_simples not in valores_diario:
                valores_diario[chave_simples] = total
            # else: # Opcional: Logar se encontrar EQL+Parcela repetido
            #     print(f"[AVISO] Chave EQL+Parcela duplicada no Diário: {chave_simples}. Usando primeiro valor {valores_diario[chave_simples]:.2f}.")

    print(f"📗 [LOG] Fim do Loop 1. Dicionário 'Diário' criado com {linhas_diario_count} linhas processadas. {len(valores_diario)} chaves únicas (EQL+Parcela). Tempo: {time.time() - start_time:.2f}s")


    print("📘 [LOG] Início do Loop 2: Processando 'Sistema'...")
    valores_sistema = {}
    linhas_sistema_count = 0
    for i, row in enumerate(ws_sistema.iter_rows(min_row=2, values_only=True)):
        linhas_sistema_count += 1
        if i % 500 == 0 and i > 0:
            print(f"➡️  [LOG] Processando linha {i+2} do Sistema (Loop 2)...")

        eql = str(row[col_eq_sistema - 1]).strip() if col_eq_sistema <= len(row) and row[col_eq_sistema - 1] else ""
        parcela = str(row[col_parcela_sistema - 1]).strip() if col_parcela_sistema <= len(row) and row[col_parcela_sistema - 1] else ""
        valor = normalizar_valor_repasse(row[col_valor_sistema - 1]) if col_valor_sistema <= len(row) else 0.0


        if eql and parcela:
            chave_simples = (eql, parcela)
            if chave_simples not in valores_sistema:
                valores_sistema[chave_simples] = valor
            # else: # Opcional: Logar se encontrar EQL+Parcela repetido
            #      print(f"[AVISO] Chave EQL+Parcela duplicada no Sistema: {chave_simples}. Usando primeiro valor {valores_sistema[chave_simples]:.2f}.")

    print(f"📗 [LOG] Fim do Loop 2. Dicionário 'Sistema' criado com {linhas_sistema_count} linhas processadas. {len(valores_sistema)} chaves únicas (EQL+Parcela). Tempo: {time.time() - start_time:.2f}s")



    print("📘 [LOG] Início do Loop 3: Comparando 'Diário' (com células) com 'Sistema'...")
    iguais = []
    divergentes = []
    duplicados_vistos = set()
    linhas_diario_loop3 = 0

    # Certifica-se de que há linhas para iterar além do cabeçalho
    if ws_diario.max_row >= 2:
         for row_idx, row_cells in enumerate(ws_diario.iter_rows(min_row=2)):
             linhas_diario_loop3 += 1
             current_row_num = row_idx + 2 # Número real da linha na planilha
             if row_idx % 500 == 0 and row_idx > 0:
                  print(f"➡️  [LOG] Processando linha {current_row_num} do Diário (Loop 3)...")

             # Acessa células com segurança, verificando o índice
             celula_eql = row_cells[col_eq_diario - 1] if col_eq_diario <= len(row_cells) else None
             celula_parcela = row_cells[col_parcela_diario - 1] if col_parcela_diario <= len(row_cells) else None
             celula_principal = row_cells[col_principal_diario - 1] if col_principal_diario <= len(row_cells) else None
             celula_correcao = row_cells[col_corrmonet_diario - 1] if col_corrmonet_diario <= len(row_cells) else None


             eql = str(celula_eql.value).strip() if celula_eql and celula_eql.value is not None else ""
             parcela = str(celula_parcela.value).strip() if celula_parcela and celula_parcela.value is not None else ""

             # Se EQL ou Parcela forem vazios/nulos nesta linha, pula para a próxima
             if not eql or not parcela:
                 print(f"[AVISO] Linha {current_row_num} do Diário ignorada (EQL ou Parcela vazio).")
                 continue

             principal = normalizar_valor_repasse(celula_principal.value if celula_principal else None)
             correcao = normalizar_valor_repasse(celula_correcao.value if celula_correcao else None)

             chave_simples = (eql, parcela)
             chave_completa = (eql, parcela, principal, correcao) # Chave com valores para checar duplicados exatos

             # Verifica duplicidade exata (mesmo EQL, Parcela, Principal e Correção)
             if contagem_diario.get(chave_completa, 0) > 1:
                 # Se já vimos essa combinação exata antes
                 if chave_completa in duplicados_vistos:
                     divergentes.append((row_cells, f"EQL {eql} Parcela {parcela} duplicada no diário (Principal={principal:.2f}, Correção={correcao:.2f})"))
                     continue # Pula para próxima linha do diário
                 else:
                     # Marca como vista, mas processa normalmente na primeira vez que aparece
                     duplicados_vistos.add(chave_completa)
                     # Poderia adicionar um aviso aqui se quisesse, mas continua o fluxo


             # Comparação com dados do sistema
             valor_diario_calculado = round(principal + correcao, 2) # Recalcula para garantir consistência
             valor_sistema = valores_sistema.get(chave_simples) # Busca valor correspondente no sistema


             if valor_sistema is None:
                  divergentes.append((row_cells, f"EQL {eql} Parcela {parcela} não encontrada no sistema"))
             elif abs(valor_diario_calculado - valor_sistema) <= 0.02: # Tolerância de 2 centavos
                  iguais.append((row_cells, "")) # Adiciona a tupla de CÉLULAS
             else:
                  divergentes.append((row_cells, f"Valor diferente (Diário={valor_diario_calculado:.2f} / Sistema={valor_sistema:.2f})"))

    else:
          print("[AVISO] Planilha 'Diário' não contém dados (além do cabeçalho) para processar no Loop 3.")


    print("📘 [LOG] Verificando itens presentes no 'Sistema' mas ausentes no 'Diário'...")
    items_sistema_apenas = 0
    for chave_simples_sistema, valor_sistema in valores_sistema.items():
        if chave_simples_sistema not in valores_diario: # Compara com as chaves simples do diário
            eql, parcela = chave_simples_sistema
            # Adiciona None para 'linha' pois não há linha correspondente no diário
            divergentes.append((None, f"EQL {eql} Parcela {parcela} presente no sistema (Valor={valor_sistema:.2f}), ausente no diário"))
            items_sistema_apenas += 1
    print(f"📗 [LOG] Fim do Loop 3 e verificação de ausentes. {linhas_diario_loop3} linhas do Diário processadas. {items_sistema_apenas} itens encontrados apenas no Sistema. Tempo total: {time.time() - start_time:.2f}s")


    print("📘 [LOG] Criando planilha 'iguais.xlsx'...")
    iguais_stream = criar_planilha_saida(iguais, ws_diario, incluir_status=False)
    print(f"📗 [LOG] Planilha 'iguais.xlsx' criada ({len(iguais)} linhas). Tempo: {time.time() - start_time:.2f}s")
    
    print("📘 [LOG] Criando planilha 'divergentes.xlsx'...")
    divergentes_stream = criar_planilha_saida(divergentes, ws_diario, incluir_status=True)
    print(f"📗 [LOG] Planilha 'divergentes.xlsx' criada ({len(divergentes)} linhas). Tempo: {time.time() - start_time:.2f}s")


    print(f"✅ [LOG] Fim de processar_repasse. Total {len(iguais)} iguais, {len(divergentes)} divergentes.")
    return iguais_stream, divergentes_stream, len(iguais), len(divergentes)

@app.route('/')
def index():
    return manual_render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'pdf_file' not in request.files or request.files['pdf_file'].filename == '':
        return manual_render_template('error.html', status_code=400,
            error_title="Nenhum arquivo enviado", 
            error_message="Você precisa selecionar um arquivo PDF para fazer a análise.")
    
    file = request.files['pdf_file']
    modo_separacao = request.form.get('modo_separacao', 'boleto') # Pega o modo selecionado

    try:
        emp_fixo = None # Para modo boleto
        # Validações específicas do modo antes de processar
        if modo_separacao == 'boleto':
            emp_fixo = detectar_emp_por_nome_arquivo(file.filename)
            if not emp_fixo:
                error_msg = ("Para o modo 'Boleto', o nome do arquivo precisa terminar com um código de empreendimento válido (ex: 'Extrato_RSCI.pdf'). "
                             "Verifique o nome do arquivo ou selecione outro modo de análise.")
                return manual_render_template('error.html', status_code=400,
                    error_title="Empreendimento não identificado (Modo Boleto)", error_message=error_msg)
        
        elif modo_separacao in ['debito_credito', 'ccb_realiza']:
             # Para estes modos, nomes como _RSCI.pdf são inválidos se o modo for selecionado incorretamente
             if detectar_emp_por_nome_arquivo(file.filename) and modo_separacao == 'debito_credito':
                  error_msg = ("Este arquivo parece ser do tipo 'Boleto' (termina com código de empreendimento), mas o modo 'Débito/Crédito' foi selecionado. "
                               "Por favor, use o modo 'Boleto' ou renomeie o arquivo se ele não for específico de um empreendimento.")
                  return manual_render_template('error.html', status_code=400,
                                                error_title="Modo de Análise Incorreto?", error_message=error_msg)
             # Nenhuma validação extra de nome para ccb_realiza por enquanto

        print(f"Iniciando validação para o arquivo '{file.filename}' no modo '{modo_separacao}'...")
        pdf_stream = file.read()
        texto_pdf = extrair_texto_pdf(pdf_stream)
        if not texto_pdf:
            print(f"Falha ao extrair texto do PDF: {file.filename}")
            return manual_render_template('error.html', status_code=500,
                error_title="Erro ao ler o PDF", 
                error_message="Não foi possível extrair o texto do arquivo enviado. Ele pode estar corrompido, ser uma imagem ou estar vazio.")

        print("Texto extraído, processando validação...")
        df_todas_raw, df_cov, df_div = processar_pdf_validacao(texto_pdf, modo_separacao, emp_fixo)
        print(f"Validação concluída. {len(df_cov)} lotes/registros encontrados, {len(df_div)} divergências.")

        
        # Filtra parcelas indesejadas do df_todas ANTES de salvar no Excel
        df_todas_filtrado = df_todas_raw.copy()
        if not df_todas_filtrado.empty:
            parcelas_para_remover = ['TOTAL A PAGAR', 'DESCONTO', 'DÉBITOS DO MÊS ANTERIOR', 'ENCARGOS POR ATRASO', 'PAGAMENTO EFETUADO'] # Incluindo mais itens
            # Usa str.upper().str.strip() para comparação robusta
            df_todas_filtrado = df_todas_filtrado[~df_todas_filtrado['Parcela'].astype(str).str.strip().str.upper().isin(parcelas_para_remover)]
            df_todas_filtrado = df_todas_filtrado[~df_todas_filtrado['Parcela'].astype(str).str.strip().str.upper().str.startswith('TOTAL BANCO')]
        print("Parcelas indesejadas filtradas da aba 'Todas_Parcelas_Extraidas'.")

        
        output = io.BytesIO()
        # Usa o DataFrame filtrado ao gerar o Excel
        dfs_to_excel = {"Divergencias": df_div, "Cobertura_Analise": df_cov, "Todas_Parcelas_Extraidas": df_todas_filtrado}
        print("Gerando arquivo Excel...")
        formatar_excel(output, dfs_to_excel)
        output.seek(0)
        print("Arquivo Excel gerado em memória.")

        # Salva o arquivo Excel fisicamente
        base_name = os.path.splitext(file.filename)[0]
        # Adiciona o modo ao nome do arquivo para clareza
        report_filename = f"relatorio_{modo_separacao}_{base_name}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
        
        try:
            with open(report_path, 'wb') as f:
                f.write(output.getvalue())
            print(f"Relatório salvo em: {report_path}")
        except Exception as e_save:
            print(f"Erro ao salvar o arquivo Excel em {report_path}: {e_save}")
            # Considerar se deve retornar erro ou apenas logar

        nao_classificados = 0
        if not df_cov.empty and 'Empreendimento' in df_cov.columns:
            nao_classificados = df_cov[df_cov['Empreendimento'] == 'NAO_CLASSIFICADO'].shape[0]
            if nao_classificados > 0:
                print(f"[AVISO] {nao_classificados} registros não puderam ser classificados por empreendimento.")

        print("Renderizando página de resultados...")
        return manual_render_template('results.html',
            # Passa os dados para o template
            divergencias_json=df_div.to_json(orient='split', index=False, date_format='iso') if not df_div.empty else 'null',
            total_lotes=len(df_cov),
            total_divergencias=len(df_div),
            nao_classificados=nao_classificados,
            download_url=url_for('download_file', filename=report_filename),
             # Informa o modo usado para exibição
            modo_usado=modo_separacao.replace('_', '/').upper() # Ex: BOLETO, DEBITO/CREDITO, CCB/REALIZA
        )
    
    except Exception as e:
        print(f"📕 [ERRO FATAL] Erro inesperado na rota /upload: {e}")
        traceback.print_exc()
        return manual_render_template('error.html', status_code=500,
            error_title="Erro inesperado no processamento", 
            error_message=f"Ocorreu um erro grave durante a análise do arquivo '{file.filename}'. Detalhes: {e}")


@app.route('/compare', methods=['POST'])
def compare_files():
    if 'pdf_mes_anterior' not in request.files or 'pdf_mes_atual' not in request.files:
        return manual_render_template('error.html', status_code=400,
            error_title="Arquivos faltando", 
            error_message="Ambos os arquivos PDF (mês anterior e atual) são necessários para a comparação.")

    file_ant = request.files['pdf_mes_anterior']
    file_atu = request.files['pdf_mes_atual']
    modo_separacao = request.form.get('modo_separacao_comp', 'boleto') # Pega o modo do form de comparação

    if file_ant.filename == '' or file_atu.filename == '':
        return manual_render_template('error.html', status_code=400,
            error_title="Arquivos faltando", 
            error_message="Selecione os dois arquivos PDF para comparar.")

    # Adiciona verificação de tipo de arquivo (opcional mas recomendado)
    if not file_ant.filename.lower().endswith('.pdf') or not file_atu.filename.lower().endswith('.pdf'):
         return manual_render_template('error.html', status_code=400,
            error_title="Tipo de Arquivo Inválido", 
            error_message="Por favor, envie apenas arquivos no formato PDF para comparação.")


    try:
        emp_fixo_boleto = None # Aplicável apenas ao modo boleto
        # Validações baseadas no modo selecionado
        if modo_separacao == 'boleto':
            emp_ant = detectar_emp_por_nome_arquivo(file_ant.filename)
            emp_atu = detectar_emp_por_nome_arquivo(file_atu.filename)
            if not emp_ant or not emp_atu:
                return manual_render_template('error.html', status_code=400,
                    error_title="Empreendimento não identificado (Modo Boleto)",
                    error_message="Para o modo 'Boleto', o nome de ambos os arquivos PDF precisa terminar com um código de empreendimento válido.")
            if emp_ant != emp_atu:
                return manual_render_template('error.html', status_code=400,
                    error_title="Empreendimentos diferentes (Modo Boleto)",
                    error_message=f"Os arquivos devem ser do mesmo empreendimento para comparação no modo Boleto (Detectado: '{emp_ant}' e '{emp_atu}').")
            emp_fixo_boleto = emp_ant # Define o empreendimento fixo para passar adiante
        
        elif modo_separacao in ['debito_credito', 'ccb_realiza']:
             # Verifica se algum arquivo parece ser de Boleto quando não deveria
             if detectar_emp_por_nome_arquivo(file_ant.filename) or detectar_emp_por_nome_arquivo(file_atu.filename):
                  error_msg = (f"Um dos arquivos parece ser do tipo 'Boleto' (termina com código), mas o modo '{modo_separacao.replace('_','/').upper()}' foi selecionado. "
                               "Use o modo 'Boleto' para esses arquivos ou renomeie-os se a detecção estiver incorreta.")
                  return manual_render_template('error.html', status_code=400,
                                                error_title="Modo de Análise Incorreto?", error_message=error_msg)
        
        print(f"Iniciando comparação modo '{modo_separacao}' entre '{file_ant.filename}' e '{file_atu.filename}'...")
        texto_ant = extrair_texto_pdf(file_ant.read())
        texto_atu = extrair_texto_pdf(file_atu.read())

        if not texto_ant or not texto_atu:
            err_msg = "Não foi possível extrair texto de um ou ambos os PDFs. "
            if not texto_ant and not texto_atu: err_msg += "Ambos os arquivos falharam."
            elif not texto_ant: err_msg += f"Falha ao ler '{file_ant.filename}'."
            else: err_msg += f"Falha ao ler '{file_atu.filename}'."
            err_msg += " Verifique se não estão corrompidos ou se são imagens."
            print(f"[ERRO] {err_msg}")
            return manual_render_template('error.html', status_code=500,
                error_title="Erro ao ler PDF na Comparação", error_message=err_msg)

        print("Textos extraídos. Processando comparação...")
        # Chama processar_comparativo passando o modo
        df_resumo_completo, df_adicionados, df_removidos, df_divergencias, df_parcelas_novas, df_parcelas_removidas = processar_comparativo(
            texto_ant, texto_atu, modo_separacao, emp_fixo_boleto
        )
        print(f"Comparação concluída. Resumo: {len(df_adicionados)} adicionados, {len(df_removidos)} removidos, {len(df_divergencias)} divergências.")


        output = io.BytesIO()
        dfs_to_excel = {
            "Resumo": df_resumo_completo, 
            "Lotes Adicionados": df_adicionados, 
            "Lotes Removidos": df_removidos,
            "Divergências de Valor": df_divergencias, 
            "Parcelas Novas por Lote": df_parcelas_novas,
            "Parcelas Removidas por Lote": df_parcelas_removidas,
        }
        print("Gerando arquivo Excel do comparativo...")
        formatar_excel(output, dfs_to_excel)
        output.seek(0)
        print("Arquivo Excel gerado em memória.")
        
        # Salva o arquivo Excel comparativo
        report_filename = f"comparativo_{modo_separacao}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
        try:
            with open(report_path, 'wb') as f: 
                f.write(output.getvalue())
            print(f"Relatório comparativo salvo em: {report_path}")
        except Exception as e_save:
             print(f"Erro ao salvar o arquivo Excel comparativo em {report_path}: {e_save}")
             # Considerar se retorna erro
        
        # Prepara dados do resumo para o template
        resumo_dict_lotes = {}
        resumo_dict_totais = {}
        if not df_resumo_completo.empty:
             resumo_dict_lotes = pd.Series(df_resumo_completo.set_index(' ')['LOTES']).to_dict()
             resumo_dict_totais = pd.Series(df_resumo_completo.set_index(' ')['TOTAIS']).map('{:,.2f}'.format).to_dict() # Formata totais


        print("Renderizando página de resultados da comparação...")
        return manual_render_template('compare_results.html',
             # Dados para o resumo
             resumo_lotes_mes_anterior=resumo_dict_lotes.get('Lotes Mês Anterior', 0),
             resumo_lotes_mes_atual=resumo_dict_lotes.get('Lotes Mês Atual', 0),
             resumo_lotes_adicionados=resumo_dict_lotes.get('Lotes Adicionados', 0),
             resumo_lotes_removidos=resumo_dict_lotes.get('Lotes Removidos', 0),
             resumo_parcelas_com_valor_alterado=resumo_dict_lotes.get('Parcelas com Valor Alterado', 0), # Este é o número de parcelas, não lotes
             
             # Totais formatados para exibição (opcional)
             total_mes_anterior_str=resumo_dict_totais.get('Lotes Mês Anterior', '0.00'),
             total_mes_atual_str=resumo_dict_totais.get('Lotes Mês Atual', '0.00'),
             total_adicionados_str=resumo_dict_totais.get('Lotes Adicionados', '0.00'),
             total_removidos_str=resumo_dict_totais.get('Lotes Removidos', '0.00'),
             total_diferencas_str=resumo_dict_totais.get('Parcelas com Valor Alterado', '0.00'),

            # Dados para as tabelas (JSON)
            divergencias_json=df_divergencias.to_json(orient='split', index=False, date_format='iso') if not df_divergencias.empty else 'null',
            adicionados_json=df_adicionados.to_json(orient='split', index=False, date_format='iso') if not df_adicionados.empty else 'null',
            removidos_json=df_removidos.to_json(orient='split', index=False, date_format='iso') if not df_removidos.empty else 'null',
            # Adiciona JSON para novas/removidas parcelas se quiser exibi-las também
            # parcelas_novas_json=df_parcelas_novas.to_json(orient='split', index=False, date_format='iso') if not df_parcelas_novas.empty else 'null',
            # parcelas_removidas_json=df_parcelas_removidas.to_json(orient='split', index=False, date_format='iso') if not df_parcelas_removidas.empty else 'null',

            download_url=url_for('download_file', filename=report_filename),
            modo_usado=modo_separacao.replace('_', '/').upper() # Passa o modo usado
        )


    except Exception as e:
        print(f"📕 [ERRO FATAL] Erro inesperado na rota /compare: {e}")
        traceback.print_exc()
        # Passa o nome do erro específico para a página de erro
        error_details = f"{type(e).__name__}: {e}"
        return manual_render_template('error.html', status_code=500,
            error_title="Erro inesperado na comparação", 
            error_message=f"Ocorreu um erro grave durante a comparação dos arquivos. Detalhes: {error_details}")


@app.route('/repasse', methods=['POST'])
def repasse_file():
    print("\n--- RECEIVED REQUEST /repasse ---")
    start_time_route = time.time()
    
    if 'diario_file' not in request.files or 'sistema_file' not in request.files:
        print("📕 [ERRO] Arquivos 'diario' ou 'sistema' faltando no request.")
        return manual_render_template('error.html', status_code=400,
            error_title="Arquivos faltando", 
            error_message="Você precisa enviar os dois arquivos Excel (Diário e Sistema) para a conciliação.")

    file_diario = request.files['diario_file']
    file_sistema = request.files['sistema_file']

    if file_diario.filename == '' or file_sistema.filename == '':
        print("📕 [ERRO] Nomes dos arquivos Excel estão vazios.")
        return manual_render_template('error.html', status_code=400,
            error_title="Arquivos faltando", 
            error_message="Selecione os dois arquivos Excel para conciliar.")

    # Validação do tipo de arquivo Excel
    allowed_extensions = {'.xlsx', '.xlsm'} # Adicione outras se necessário
    diario_ext = os.path.splitext(file_diario.filename)[1].lower()
    sistema_ext = os.path.splitext(file_sistema.filename)[1].lower()
    if diario_ext not in allowed_extensions or sistema_ext not in allowed_extensions:
         print(f"📕 [ERRO] Extensões de arquivo inválidas: {diario_ext}, {sistema_ext}")
         return manual_render_template('error.html', status_code=400,
             error_title="Tipo de Arquivo Inválido", 
             error_message=f"Por favor, envie apenas arquivos Excel ({', '.join(allowed_extensions)}). Recebido: {file_diario.filename}, {file_sistema.filename}")

    
    print(f"📘 [LOG] Recebidos Excel: {file_diario.filename}, {file_sistema.filename}")
    
    try:
        diario_stream = io.BytesIO(file_diario.read())
        sistema_stream = io.BytesIO(file_sistema.read())
        print(f"📘 [LOG] Arquivos Excel lidos em memória. Tempo: {time.time() - start_time_route:.2f}s")

        iguais_stream, divergentes_stream, count_iguais, count_divergentes = processar_repasse(diario_stream, sistema_stream)
        
        print(f"📘 [LOG] Processamento de repasse concluído. Criando ZIP... Tempo total rota: {time.time() - start_time_route:.2f}s")
        zip_stream = io.BytesIO()
        # Usar nomes mais descritivos e seguros
        timestamp_str = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
        safe_diario_name = re.sub(r'[^\w\-.]', '_', os.path.splitext(file_diario.filename)[0])
        safe_sistema_name = re.sub(r'[^\w\-.]', '_', os.path.splitext(file_sistema.filename)[0])
        
        zip_filename_iguais = f"iguais_{safe_diario_name}_vs_{safe_sistema_name}_{timestamp_str}.xlsx"
        zip_filename_divergentes = f"divergentes_{safe_diario_name}_vs_{safe_sistema_name}_{timestamp_str}.xlsx"

        with zipfile.ZipFile(zip_stream, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(zip_filename_iguais, iguais_stream.getvalue())
            zf.writestr(zip_filename_divergentes, divergentes_stream.getvalue())
        zip_stream.seek(0)
        print(f"📗 [LOG] ZIP criado com arquivos: {zip_filename_iguais}, {zip_filename_divergentes}. Tempo rota: {time.time() - start_time_route:.2f}s")
        
        # Nome do arquivo ZIP para download
        report_filename = f"repasse_conciliado_{timestamp_str}.zip"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
        try:
            with open(report_path, 'wb') as f: 
                f.write(zip_stream.getvalue())
            print(f"📗 [LOG] Arquivo ZIP salvo em {report_path}. Tempo rota: {time.time() - start_time_route:.2f}s")
        except Exception as e_save:
             print(f"📕 [ERRO] Erro ao salvar o arquivo ZIP em {report_path}: {e_save}")
             # Considerar retornar erro ao usuário
             raise e_save # Re-levanta o erro para ser pego pelo bloco except externo


        print("✅ [LOG] Enviando resposta para 'repasse_results.html'")
        return manual_render_template('repasse_results.html',
            count_iguais=count_iguais,
            count_divergentes=count_divergentes,
            download_url=url_for('download_file', filename=report_filename)
        )

    except ValueError as ve: # Captura erros de coluna não encontrada especificamente
         print(f"📕 [ERRO VALIDAÇÃO] {ve}")
         traceback.print_exc()
         return manual_render_template('error.html', status_code=400, # Bad request por causa dos arquivos
             error_title="Erro na Conciliação - Colunas Não Encontradas", 
             error_message=f"Verifique os nomes das colunas nas planilhas. Detalhes: {ve}")
    except Exception as e: # Captura outros erros inesperados
        print(f"📕 [ERRO FATAL] Erro inesperado na rota /repasse: {e}")
        traceback.print_exc()
        error_details = f"{type(e).__name__}: {e}"
        return manual_render_template('error.html', status_code=500,
            error_title="Erro inesperado na conciliação", 
            error_message=f"Ocorreu um erro grave durante a análise. Detalhes: {error_details}")


@app.route('/download/<filename>')
def download_file(filename):
     # Segurança: Verifica se o nome do arquivo é seguro e se ele existe
     safe_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
     if not os.path.normpath(safe_path).startswith(os.path.abspath(app.config['UPLOAD_FOLDER'])):
         print(f" Tentativa de acesso a caminho inválido: {filename}")
         return "Acesso negado.", 403
     
     if not os.path.exists(safe_path):
          print(f" Arquivo não encontrado para download: {filename}")
          return "Arquivo não encontrado.", 404
          
     print(f"Enviando arquivo para download: {filename}")
     # O as_attachment=True força o download
     return send_file(safe_path, as_attachment=True)


if __name__ == '__main__':
    print("Iniciando servidor Flask local...")
    # Usa a porta definida pelo ambiente (Render) ou 8080 por padrão
    port = int(os.environ.get('PORT', 8080))
    # debug=True é útil localmente, mas deve ser False em produção (Render define automaticamente)
    # use_reloader=False e threaded=True foram para debug local, pode remover se não precisar mais
    # host='0.0.0.0' é necessário para o Render acessar o app dentro do container
    print(f"Executando em http://0.0.0.0:{port} (debug={'True' if os.environ.get('FLASK_DEBUG') == '1' else 'False'})")
    app.run(debug=(os.environ.get('FLASK_DEBUG') == '1'), host='0.0.0.0', port=port)

