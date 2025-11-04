# -*- coding: utf-8 -*-
import os
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
BASE_FIXOS_CCB = {
    "Alienação Fiduciária CCB": [],
    "Financiamento Realiza CCB": [],
    "Encargos Não Pagos CCB": [],
    "Débito por pagamento a menor CCB": [],
    "Crédito por pagamento a maior CCB": [],
    "Negociação Alienação CCB": []
}

app = Flask(__name__)
# Define UPLOAD_FOLDER como um caminho absoluto relativo à raiz do app
UPLOAD_FOLDER_PATH = os.path.join(app.root_path, 'uploads')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER_PATH
# Cria o diretório usando o caminho absoluto
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
print(f"Pasta de Upload configurada em: {app.config['UPLOAD_FOLDER']}")

def manual_render_template(template_name, status_code=200, **kwargs):
    template_path = os.path.join(app.root_path, 'templates', template_name)
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            html_content = f.read()

        for key, value in kwargs.items():
            placeholder = f"__{key.upper()}__"
            # Trata json strings e outros tipos
            if isinstance(value, str) and value.startswith('{') and value.endswith('}'):
                 # Substitui placeholder entre aspas por valor json sem aspas
                 html_content = html_content.replace(f'"{placeholder}"', value)
            else:
                 # Substituição normal para outros tipos
                 html_content = html_content.replace(placeholder, str(value))

        response = make_response(html_content)
        response.headers['Content-Type'] = 'text/html'
        return response, status_code
    except Exception as e:
        print(f"ERRO CRÍTICO AO RENDERIZAR MANUALMENTE '{template_name}': {e}")
        # Retorna uma página de erro mais informativa
        error_html = f"""
        <!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Erro 500</title></head>
        <body><h1>Erro 500: Falha Crítica ao Carregar Template</h1>
        <p>Ocorreu um erro interno ao tentar carregar ou processar o template <strong>{template_name}</strong>.</p>
        <p><strong>Detalhes do Erro:</strong> {e}</p>
        <p>Verifique se o arquivo existe no caminho esperado ({template_path}) e se o conteúdo é válido.</p>
        </body></html>
        """
        return make_response(error_html, 500)


def normalizar_texto(s: str) -> str:
    s = s.translate(DASHES).replace("\u00A0", " ")
    s = "".join(ch for ch in s if ch not in "\u200B\u200C\u200D\uFEFF")
    s = unicodedata.normalize("NFKC", s)
    return s

def extrair_texto_pdf(stream_pdf) -> str:
    texto_completo = ""
    try:
        with fitz.open(stream=stream_pdf, filetype="pdf") as doc:
             for page_num in range(len(doc)):
                 page = doc.load_page(page_num)
                 texto_pagina = page.get_text("text", sort=True)
                 texto_completo += texto_pagina + "\n"
        return normalizar_texto(texto_completo)
    except Exception as e:
        print(f"Erro detalhado ao ler o stream do PDF: {type(e).__name__} - {e}")
        traceback.print_exc()
        return ""

def to_float(s: str):
    if s is None:
        return None
    try:
        cleaned_s = str(s).strip().replace("R$", "").replace(".", "").replace(",", ".")
        return float(cleaned_s)
    except (ValueError, TypeError) as e:
        return None

def fixos_do_emp(emp: str, modo_separacao: str):
    if modo_separacao == 'boleto':
        if emp not in EMP_MAP:
            return BASE_FIXOS
        f = dict(BASE_FIXOS)
        if EMP_MAP.get(emp):
            if "Melhoramentos" in EMP_MAP[emp]:
                f["Melhoramentos"] = [float(EMP_MAP[emp]["Melhoramentos"])]
            if "Fundo de Transporte" in EMP_MAP[emp]:
                f["Fundo de Transporte"] = [float(EMP_MAP[emp]["Fundo de Transporte"])]
        return f
    elif modo_separacao == 'debito_credito':
        return BASE_FIXOS
    elif modo_separacao == 'ccb_realiza':
        return BASE_FIXOS_CCB
    else:
        print(f"[AVISO] Modo de separação desconhecido '{modo_separacao}' em fixos_do_emp.")
        return {}

def detectar_emp_por_nome_arquivo(path: str):
    if not path: return None
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
    if not isinstance(lbl, str): return ""
    lbl = re.sub(r"^TAMA\s*[-–—]\s*", "", lbl, flags=re.IGNORECASE).strip()
    lbl = re.sub(r"\s+-\s+\d+/\d+$", "", lbl).strip()
    lbl = re.sub(r'\s{2,}', ' ', lbl).strip()
    return lbl

def fatiar_blocos(texto: str):
    texto_processado = PADRAO_LOTE.sub(r"\n\1", texto)
    matches = list(PADRAO_LOTE.finditer(texto_processado))
    blocos = []
    for i, match in enumerate(matches):
        lote_atual = match.group(1)
        inicio_bloco = match.start()
        fim_bloco = matches[i+1].start() if i+1 < len(matches) else len(texto_processado)
        texto_bloco = texto_processado[inicio_bloco:fim_bloco].strip()
        if texto_bloco:
             blocos.append((lote_atual, texto_bloco))
    if not blocos:
         print("[AVISO] Nenhum bloco de lote encontrado no PDF.")
    return blocos

def tentar_nome_cliente(bloco: str) -> str:
    linhas = bloco.split('\n')
    if not linhas: return "Nome não localizado"

    linhas_para_buscar = linhas[:6]
    nome_candidato = "Nome não localizado"

    for linha in linhas_para_buscar:
        linha_sem_lote = PADRAO_LOTE.sub('', linha).strip()
        if not linha_sem_lote: continue

        is_valid_name = (
            len(linha_sem_lote) > 5 and
            ' ' in linha_sem_lote and
            sum(c.isalpha() for c in linha_sem_lote.replace(" ", "")) / len(linha_sem_lote.replace(" ", "")) > 0.7 and
            not any(h.upper() in linha_sem_lote.upper() for h in HEADERS if h) and
            not re.search(r'\d{2}/\d{2}/\d{4}', linha_sem_lote) and
            not re.match(r'^[\d.,\s]+$', linha_sem_lote) and
            not linha_sem_lote.upper().startswith(("TOTAL", "BANCO", "03-", "LIMITE P/", "PÁGINA"))
        )

        if is_valid_name:
            nome_candidato = linha_sem_lote
            break

    return nome_candidato.strip()

def extrair_parcelas(bloco: str):
    itens = OrderedDict()
    pos_lancamentos = bloco.find("Lançamentos")
    bloco_de_trabalho = bloco[pos_lancamentos:] if pos_lancamentos != -1 else bloco

    bloco_limpo_linhas = []
    linhas_originais = bloco_de_trabalho.splitlines()
    ignorar_proxima_linha_se_numero = False

    for i, linha in enumerate(linhas_originais):
        match_total_direita = re.search(r'\s{4,}(DÉBITOS DO MÊS ANTERIOR|ENCARGOS POR ATRASO|PAGAMENTO EFETUADO)\s+[\d.,]+$', linha)
        linha_processada = linha[:match_total_direita.start()] if match_total_direita else linha
        linha_processada = linha_processada.strip()

        if not linha_processada or any(h.strip().upper() == linha_processada.upper() for h in ["Lançamentos", "Débitos do Mês"]):
            continue

        if ignorar_proxima_linha_se_numero:
             ignorar_proxima_linha_se_numero = False
             continue

        match_mesma_linha = PADRAO_PARCELA_MESMA_LINHA.match(linha_processada)
        if match_mesma_linha:
            lbl = limpar_rotulo(match_mesma_linha.group(1))
            val = to_float(match_mesma_linha.group(2))
            if lbl and lbl not in itens and val is not None:
                itens[lbl] = val
                continue

        is_potential_label = (
            any(c.isalpha() for c in linha_processada) and
            limpar_rotulo(linha_processada) not in itens
        )

        if is_potential_label:
            j = i + 1
            while j < len(linhas_originais) and not linhas_originais[j].strip():
                j += 1
            if j < len(linhas_originais):
                 linha_seguinte_limpa = linhas_originais[j].strip()
                 match_num_puro = PADRAO_NUMERO_PURO.match(linha_seguinte_limpa)
                 if match_num_puro:
                      lbl = limpar_rotulo(linha_processada)
                      val = to_float(match_num_puro.group(1))
                      if lbl and lbl not in itens and val is not None:
                           itens[lbl] = val
                           ignorar_proxima_linha_se_numero = True
                           continue
    return itens

def processar_pdf_validacao(texto_pdf: str, modo_separacao: str, emp_fixo_boleto: str = None):
    """Processa o texto do PDF para validação."""
    blocos = fatiar_blocos(texto_pdf)
    if not blocos: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    linhas_todas, linhas_cov, linhas_div = [], [], []
    for lote, bloco in blocos:
        if modo_separacao == 'boleto':
            emp_atual = detectar_emp_por_lote(lote) if emp_fixo_boleto == "SBRR" else emp_fixo_boleto
        else:
            emp_atual = detectar_emp_por_lote(lote)

        cliente = tentar_nome_cliente(bloco)
        itens = extrair_parcelas(bloco)
        VALORES_CORRETOS = fixos_do_emp(emp_atual, modo_separacao)

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

        if modo_separacao != 'ccb_realiza':
            for rot in vistos:
                val = cov[rot]
                if val is None: continue
                permitidos = VALORES_CORRETOS.get(rot, [])
                if permitidos and all(abs(val - v) > 1e-6 for v in permitidos):
                    linhas_div.append({
                        "Empreendimento": emp_atual, "Lote": lote, "Cliente": cliente,
                        "Parcela": rot, "Valor no Documento": float(val),
                        "Valor Correto": " ou ".join(f"{v:.2f}" for v in permitidos)
                    })

    df_todas = pd.DataFrame(linhas_todas)
    df_cov = pd.DataFrame(linhas_cov)
    df_div = pd.DataFrame(linhas_div)

    return df_todas, df_cov, df_div

def processar_comparativo(texto_anterior, texto_atual, modo_separacao, emp_fixo_boleto):
    """Compara os dados extraídos de dois PDFs."""
    df_todas_ant_raw, _, _ = processar_pdf_validacao(texto_anterior, modo_separacao, emp_fixo_boleto)
    df_todas_atu_raw, _, _ = processar_pdf_validacao(texto_atual, modo_separacao, emp_fixo_boleto)

    df_totais_ant = df_todas_ant_raw[df_todas_ant_raw['Parcela'].astype(str).str.strip().str.upper() == 'TOTAL A PAGAR'].copy()
    df_totais_ant = df_totais_ant[['Empreendimento', 'Lote', 'Cliente', 'Valor']].rename(columns={'Valor': 'Total Anterior'})

    df_totais_atu = df_todas_atu_raw[df_todas_atu_raw['Parcela'].astype(str).str.strip().str.upper() == 'TOTAL A PAGAR'].copy()
    df_totais_atu = df_totais_atu[['Empreendimento', 'Lote', 'Cliente', 'Valor']].rename(columns={'Valor': 'Total Atual'})

    parcelas_para_remover = ['TOTAL A PAGAR', 'DESCONTO', 'DÉBITOS DO MÊS', 'DÉBITOS DO MÊS ANTERIOR', 'ENCARGOS POR ATRASO', 'PAGAMENTO EFETUADO']
    df_todas_ant = df_todas_ant_raw[~df_todas_ant_raw['Parcela'].astype(str).str.strip().str.upper().isin(parcelas_para_remover)].copy()
    df_todas_atu = df_todas_atu_raw[~df_todas_atu_raw['Parcela'].astype(str).str.strip().str.upper().isin(parcelas_para_remover)].copy()

    df_todas_ant = df_todas_ant[~df_todas_ant['Parcela'].astype(str).str.strip().str.upper().str.startswith('TOTAL BANCO')].copy()
    df_todas_atu = df_todas_atu[~df_todas_atu['Parcela'].astype(str).str.strip().str.upper().str.startswith('TOTAL BANCO')].copy()

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

    df_divergencias = df_comp[
        (pd.notna(df_comp['Valor Anterior'])) &
        (pd.notna(df_comp['Valor Atual'])) &
        (abs(df_comp['Valor Anterior'] - df_comp['Valor Atual']) > 0.025)
    ].copy()
    if not df_divergencias.empty:
         df_divergencias['Diferença'] = df_divergencias['Valor Atual'] - df_divergencias['Valor Anterior']

    df_parcelas_novas = df_comp[df_comp['Valor Anterior'].isna() & pd.notna(df_comp['Valor Atual'])][['Empreendimento', 'Lote', 'Cliente', 'Parcela', 'Valor Atual']].copy()
    df_parcelas_removidas = df_comp[df_comp['Valor Atual'].isna() & pd.notna(df_comp['Valor Anterior'])][['Empreendimento', 'Lote', 'Cliente', 'Parcela', 'Valor Anterior']].copy()

    total_adicionados_valor = df_adicionados['Total Atual'].sum() if 'Total Atual' in df_adicionados.columns else 0
    total_removidos_valor = df_removidos['Total Anterior'].sum() if 'Total Anterior' in df_removidos.columns else 0
    total_divergencias_valor = df_divergencias['Diferença'].sum() if 'Diferença' in df_divergencias.columns else 0
    total_mes_anterior_valor = df_totais_ant['Total Anterior'].sum() if 'Total Anterior' in df_totais_ant.columns else 0
    total_mes_atual_valor = df_totais_atu['Total Atual'].sum() if 'Total Atual' in df_totais_atu.columns else 0

    resumo_financeiro_data = {
        ' ': ['Lotes Mês Anterior', 'Lotes Mês Atual', 'Lotes Adicionados', 'Lotes Removidos', 'Parcelas com Valor Alterado'],
        'LOTES': [len(lotes_ant), len(lotes_atu), len(df_adicionados), len(df_removidos), df_divergencias['Lote'].nunique() if not df_divergencias.empty else 0],
        'TOTAIS': [total_mes_anterior_valor, total_mes_atual_valor, total_adicionados_valor, total_removidos_valor, total_divergencias_valor]
    }
    df_resumo_completo = pd.DataFrame(resumo_financeiro_data)

    return df_resumo_completo, df_adicionados, df_removidos, df_divergencias, df_parcelas_novas, df_parcelas_removidas


def formatar_excel(output_stream, dfs: dict):
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, NamedStyle
    from openpyxl.utils import get_column_letter
    import pandas as pd

    with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
        for sheet_name, df in dfs.items():
            if df is None:
                continue
            if isinstance(df, pd.DataFrame):
                 df.to_excel(writer, index=False, sheet_name=sheet_name)
            else:
                 print(f"[AVISO] Tentando salvar algo que não é DataFrame na planilha '{sheet_name}': {type(df)}")
                 pd.DataFrame([{"Erro": f"Dados inválidos para {sheet_name}"}]).to_excel(writer, index=False, sheet_name=sheet_name)

        number_style = NamedStyle(name='br_number_style', number_format='#,##0.00')
        integer_style = NamedStyle(name='br_integer_style', number_format='0')

        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            worksheet.sheet_view.showGridLines = False

            if worksheet.max_row > 0 and worksheet.max_column > 0:
                for col_idx, column_cells in enumerate(worksheet.columns, 1):
                    max_length = 0
                    column = get_column_letter(col_idx)
                    is_first_row = True
                    for cell in column_cells:
                        if cell.value:
                             try:
                                 if isinstance(cell.value, (int, float)) and cell.number_format != 'General':
                                     formatted_value = f"{cell.value:{cell.number_format.replace('#,##','').replace('0.00','.2f')}}"
                                     max_length = max(max_length, len(formatted_value))
                                 else:
                                      max_length = max(max_length, len(str(cell.value)))
                             except:
                                  max_length = max(max_length, len(str(cell.value)))

                        if not is_first_row:
                             if isinstance(cell.value, float):
                                 cell.style = number_style
                             elif isinstance(cell.value, int):
                                 if sheet_name == 'Resumo' and column == 'B':
                                     cell.style = integer_style
                                 elif column != 'B':
                                     cell.style = number_style
                        is_first_row = False

                    adjusted_width = (max_length + 2) * 1.15
                    worksheet.column_dimensions[column].width = min(adjusted_width, 60)

                # ===> ADICIONA O AUTOFILTRO <===
                worksheet.auto_filter.ref = worksheet.dimensions
                print(f"[LOG] Autofilter aplicado à planilha '{sheet_name}'. Ref: {worksheet.dimensions}")
                # ==============================
    return output_stream


def normalizar_valor_repasse(valor):
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return round(float(valor), 2)
    s = str(valor).strip().replace("R$", "").replace(" ", "").replace("\xa0", "")
    if "," in s and "." in s: # Formato 1.234,56
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s: # Formato 1234,56
        s = s.replace(",", ".")
    # Assume 1234.56 ou 1234
    try:
        return round(float(s), 2)
    except ValueError:
        print(f"[AVISO] Falha ao normalizar valor repasse: '{valor}' -> '{s}'")
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
    if sheet.max_row == 0: return None
    for cell in sheet[1]:
        if cell.value and str(cell.value).strip().lower() == nome_coluna.lower():
            return cell.column
    return None

def criar_planilha_saida(linhas, ws_diario, incluir_status=False):
    wb_out = Workbook()
    ws_out = wb_out.active

    if ws_diario.max_row > 0:
        num_cols_header = ws_diario.max_column
        for i, cell in enumerate(ws_diario[1], 1):
            if cell:
                novo = ws_out.cell(row=1, column=i, value=cell.value)
                copiar_formatacao(cell, novo)
                col_letter = get_column_letter(i)
                if col_letter in ws_diario.column_dimensions:
                     ws_out.column_dimensions[col_letter].width = ws_diario.column_dimensions[col_letter].width
                else: ws_out.column_dimensions[col_letter].width = 15
            else:
                 ws_out.cell(row=1, column=i, value=None)
    else:
        num_cols_header = 0

    col_status = 0
    if incluir_status:
        col_status = num_cols_header + 1
        cell_status_header = ws_out.cell(row=1, column=col_status, value="Status")
        cell_status_header.font = Font(bold=True)
        ws_out.column_dimensions[get_column_letter(col_status)].width = 30

    linha_out = 2
    for linha_info in linhas:
        linha, status = linha_info
        if linha is None:
            if incluir_status and col_status > 0:
                ws_out.cell(row=linha_out, column=col_status, value=status)
            linha_out += 1
            continue

        for i, cell_data in enumerate(linha, 1):
             try:
                 valor = cell_data.value if hasattr(cell_data, "value") else cell_data
                 novo = ws_out.cell(row=linha_out, column=i, value=valor)
                 if hasattr(cell_data, "value"):
                     copiar_formatacao(cell_data, novo)
             except Exception as e:
                  print(f"[Aviso] Erro ao processar célula {i} da linha {linha_out}: {e}. Valor: {cell_data}")
                  ws_out.cell(row=linha_out, column=i, value=f"ERRO: {e}")

        if incluir_status and col_status > 0:
             ws_out.cell(row=linha_out, column=col_status, value=status)
        linha_out += 1

    if incluir_status and len(linhas) > 0:
         total_cell = ws_out.cell(row=linha_out + 1, column=1)
         total_cell.value = f"Total divergentes/não encontrados: {len(linhas)}"
         total_cell.font = Font(bold=True)

    if ws_out.max_row > 0 and ws_out.max_column > 0:
        ws_out.auto_filter.ref = ws_out.calculate_dimension()
        print(f"[LOG] Autofilter aplicado à planilha de saída do repasse. Ref: {ws_out.auto_filter.ref}")

    stream_out = io.BytesIO()
    wb_out.save(stream_out)
    stream_out.seek(0)
    return stream_out

def salvar_stream_em_arquivo(stream, caminho):
    """Salva BytesIO ou bytes em arquivo binário."""
    try:
        with open(caminho, "wb") as f:
            if hasattr(stream, "getvalue"): # É BytesIO
                f.write(stream.getvalue())
            elif isinstance(stream, (bytes, bytearray)): # Já são bytes
                f.write(stream)
            else:
                 raise TypeError(f"Tipo inesperado para salvar: {type(stream)}")
        print(f"Stream salvo com sucesso em: {caminho}")
    except Exception as e:
        print(f"📕 [ERRO] Falha ao salvar stream em '{caminho}': {e}")
        raise

def processar_repasse(diario_stream, sistema_stream):
    """Lógica original de conciliação PickMoney (Diário vs Sistema)."""
    print("📘 [LOG] Início de processar_repasse (PickMoney)")
    start_time = time.time()

    print("📘 [LOG] Carregando workbook 'Diário'...")
    wb_diario = load_workbook(diario_stream, data_only=True)
    ws_diario = wb_diario.worksheets[0]
    print(f"📗 [LOG] 'Diário' carregado ({ws_diario.max_row} linhas). Tempo: {time.time() - start_time:.2f}s")

    print("📘 [LOG] Carregando workbook 'Sistema'...")
    wb_sistema = load_workbook(sistema_stream, data_only=True)
    ws_sistema = wb_sistema.worksheets[0]
    print(f"📗 [LOG] 'Sistema' carregado ({ws_sistema.max_row} linhas). Tempo: {time.time() - start_time:.2f}s")

    print("📘 [LOG] Achando colunas (PickMoney)...")
    col_eq_diario = achar_coluna(ws_diario, "EQL")
    col_parcela_diario = achar_coluna(ws_diario, "Parcela")
    col_principal_diario = 4
    col_corrmonet_diario = 9

    col_eq_sistema = achar_coluna(ws_sistema, "EQL")
    col_parcela_sistema = achar_coluna(ws_sistema, "Parcela")
    col_valor_sistema = achar_coluna(ws_sistema, "Valor")
    print(f"📗 [LOG] Colunas encontradas: Diário(EQL:{col_eq_diario}, Parc:{col_parcela_diario}), Sistema(EQL:{col_eq_sistema}, Parc:{col_parcela_sistema}, Val:{col_valor_sistema})")

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

    print(f"📗 [LOG] Fim do Loop 1. 'Diário' processado ({linhas_diario_count} linhas). {len(valores_diario)} chaves únicas. Tempo: {time.time() - start_time:.2f}s")

    print("📘 [LOG] Início do Loop 2: Processando 'Sistema'...")
    valores_sistema = {}
    linhas_sistema_count = 0
    for i, row in enumerate(ws_sistema.iter_rows(min_row=2, values_only=True)):
        linhas_sistema_count += 1

        eql = str(row[col_eq_sistema - 1]).strip() if col_eq_sistema <= len(row) and row[col_eq_sistema - 1] else ""
        parcela = str(row[col_parcela_sistema - 1]).strip() if col_parcela_sistema <= len(row) and row[col_parcela_sistema - 1] else ""
        valor = normalizar_valor_repasse(row[col_valor_sistema - 1]) if col_valor_sistema <= len(row) else 0.0

        if eql and parcela:
            chave_simples = (eql, parcela)
            if chave_simples not in valores_sistema:
                valores_sistema[chave_simples] = valor

    print(f"📗 [LOG] Fim do Loop 2. 'Sistema' processado ({linhas_sistema_count} linhas). {len(valores_sistema)} chaves únicas. Tempo: {time.time() - start_time:.2f}s")

    print("📘 [LOG] Início do Loop 3: Comparando 'Diário' com 'Sistema'...")
    iguais = []
    divergentes = []
    nao_encontrados_diario = []
    duplicados_vistos = set()
    linhas_diario_loop3 = 0

    if ws_diario.max_row >= 2:
         for row_idx, row_cells in enumerate(ws_diario.iter_rows(min_row=2)):
             linhas_diario_loop3 += 1
             current_row_num = row_idx + 2

             celula_eql = row_cells[col_eq_diario - 1] if col_eq_diario <= len(row_cells) else None
             celula_parcela = row_cells[col_parcela_diario - 1] if col_parcela_diario <= len(row_cells) else None
             celula_principal = row_cells[col_principal_diario - 1] if col_principal_diario <= len(row_cells) else None
             celula_correcao = row_cells[col_corrmonet_diario - 1] if col_corrmonet_diario <= len(row_cells) else None

             eql = str(celula_eql.value).strip() if celula_eql and celula_eql.value is not None else ""
             parcela = str(celula_parcela.value).strip() if celula_parcela and celula_parcela.value is not None else ""

             if not eql or not parcela: continue

             principal = normalizar_valor_repasse(celula_principal.value if celula_principal else None)
             correcao = normalizar_valor_repasse(celula_correcao.value if celula_correcao else None)
             valor_diario_calculado = round(principal + correcao, 2)

             chave_simples = (eql, parcela)
             chave_completa = (eql, parcela, principal, correcao)

             if contagem_diario.get(chave_completa, 0) > 1:
                 if chave_completa in duplicados_vistos:
                     divergentes.append((row_cells, f"Duplicado no Diário (EQL {eql}, P {parcela}, V {valor_diario_calculado:.2f})"))
                     continue
                 else:
                     duplicados_vistos.add(chave_completa)

             valor_sistema = valores_sistema.get(chave_simples)

             if valor_sistema is None:
                 nao_encontrados_diario.append((row_cells, f"Não encontrado no Sistema (Valor Diário={valor_diario_calculado:.2f})"))
             elif abs(valor_diario_calculado - valor_sistema) <= 0.02:
                 iguais.append((row_cells, ""))
             else:
                 divergentes.append((row_cells, f"Valor diferente (Diário={valor_diario_calculado:.2f} / Sistema={valor_sistema:.2f})"))
    else:
        print("[AVISO] Planilha 'Diário' sem dados para Loop 3.")

    print("📘 [LOG] Verificando itens do 'Sistema' ausentes no 'Diário'...")
    nao_encontrados_sistema_formatado = []
    items_sistema_apenas = 0
    for chave_simples_sistema, valor_sistema in valores_sistema.items():
        if chave_simples_sistema not in valores_diario:
            eql, parcela = chave_simples_sistema
            status_msg = f"Presente no Sistema (EQL {eql}, P {parcela}, Valor={valor_sistema:.2f}), Ausente no Diário"
            nao_encontrados_sistema_formatado.append((None, status_msg))
            items_sistema_apenas += 1

    print(f"📗 [LOG] Fim Comparação. {linhas_diario_loop3} linhas Diário. {len(nao_encontrados_diario)} não encontradas. {items_sistema_apenas} só no Sistema. Tempo: {time.time() - start_time:.2f}s")

    print("📘 [LOG] Criando planilhas de saída...")
    iguais_stream = criar_planilha_saida(iguais, ws_diario, incluir_status=False)
    divergentes_stream = criar_planilha_saida(divergentes, ws_diario, incluir_status=True)
    nao_encontrados_combinados = nao_encontrados_diario + nao_encontrados_sistema_formatado
    nao_encontrados_stream = criar_planilha_saida(nao_encontrados_combinados, ws_diario, incluir_status=True)

    timestamp_str = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
    pasta_saida = os.path.join(app.config['UPLOAD_FOLDER'], f"repasse_pickmoney_{timestamp_str}")
    os.makedirs(pasta_saida, exist_ok=True)
    print(f"Pasta de saída criada: {pasta_saida}")

    try:
        salvar_stream_em_arquivo(iguais_stream, os.path.join(pasta_saida, "iguais.xlsx"))
        salvar_stream_em_arquivo(divergentes_stream, os.path.join(pasta_saida, "divergentes.xlsx"))
        salvar_stream_em_arquivo(nao_encontrados_stream, os.path.join(pasta_saida, "nao_encontrados.xlsx"))
        print(f"📗 [LOG] Arquivos Excel (PickMoney) salvos na pasta: {pasta_saida}")
    except Exception as e_save:
         print(f"📕 [ERRO] Falha ao salvar arquivos Excel (PickMoney) na pasta {pasta_saida}: {e_save}")
         raise

    count_nao_encontrados = len(nao_encontrados_combinados)
    print(f"✅ [LOG] Fim de processar_repasse (PickMoney). Totais: Iguais={len(iguais)}, Divergentes={len(divergentes)}, Não Encontrados={count_nao_encontrados}. Tempo total: {time.time() - start_time:.2f}s")
    return pasta_saida, len(iguais), len(divergentes), count_nao_encontrados


# =======================================================
# === FUNÇÃO ABRASMA CORRIGIDA ===
# =======================================================
def processar_repasse_abrasma(anterior_stream, complementar_stream):
    """Lógica de conciliação ABRASMA (Anterior vs Complementar) usando colunas EQL, Parc, Total Recebido."""
    print("📘 [LOG] Início de processar_repasse_abrasma")
    start_time = time.time()

    print("📘 [LOG] Carregando workbook 'Planilha Anterior'...")
    wb_ant = load_workbook(anterior_stream, data_only=True)
    ws_ant = wb_ant.worksheets[0]
    print(f"📗 [LOG] 'Anterior' carregada ({ws_ant.max_row} linhas). Tempo: {time.time() - start_time:.2f}s")

    print("📘 [LOG] Carregando workbook 'Planilha Complementar'...")
    wb_comp = load_workbook(complementar_stream, data_only=True)
    ws_comp = wb_comp.worksheets[0]
    print(f"📗 [LOG] 'Complementar' carregada ({ws_comp.max_row} linhas). Tempo: {time.time() - start_time:.2f}s")

    print("📘 [LOG] Achando colunas (ABRASMA)...")
    col_eql_ant = achar_coluna(ws_ant, "EQL")
    col_parc_ant = achar_coluna(ws_ant, "Parc")
    col_total_ant = achar_coluna(ws_ant, "Total Recebido")

    col_eql_comp = achar_coluna(ws_comp, "EQL")
    col_parc_comp = achar_coluna(ws_comp, "Parc")
    col_total_comp = achar_coluna(ws_comp, "Total Recebido")

    print(f"📗 [LOG] Colunas encontradas: Anterior(EQL:{col_eql_ant}, Parc:{col_parc_ant}, Total:{col_total_ant}), Complementar(EQL:{col_eql_comp}, Parc:{col_parc_comp}, Total:{col_total_comp})")

    missing_cols = []
    if not col_eql_ant: missing_cols.append("EQL (Anterior)")
    if not col_parc_ant: missing_cols.append("Parc (Anterior)")
    if not col_total_ant: missing_cols.append("Total Recebido (Anterior)")
    if not col_eql_comp: missing_cols.append("EQL (Complementar)")
    if not col_parc_comp: missing_cols.append("Parc (Complementar)")
    if not col_total_comp: missing_cols.append("Total Recebido (Complementar)")

    if missing_cols:
         error_msg = f"Não foi possível encontrar as seguintes colunas obrigatórias: {', '.join(missing_cols)}. Verifique os nomes nos cabeçalhos."
         print(f"📕 [ERRO] {error_msg}")
         raise ValueError(error_msg)

    print("📘 [LOG] Loop 1 (ABRASMA): Processando 'Anterior' (values_only)...")
    valores_ant = {}
    contagem_ant = {}
    linhas_ant_count = 0
    for i, row in enumerate(ws_ant.iter_rows(min_row=2, values_only=True)):
        linhas_ant_count += 1
        
        eql = str(row[col_eql_ant - 1]).strip() if col_eql_ant <= len(row) and row[col_eql_ant - 1] else ""
        parc = str(row[col_parc_ant - 1]).strip() if col_parc_ant <= len(row) and row[col_parc_ant - 1] else ""
        total = normalizar_valor_repasse(row[col_total_ant - 1]) if col_total_ant <= len(row) else 0.0

        if eql and parc:
            chave_completa = (eql, parc, total)
            chave_simples = (eql, parc)
            contagem_ant[chave_completa] = contagem_ant.get(chave_completa, 0) + 1
            if chave_simples not in valores_ant:
                valores_ant[chave_simples] = total

    print(f"📗 [LOG] Fim Loop 1. 'Anterior' processada ({linhas_ant_count} linhas). {len(valores_ant)} chaves. Tempo: {time.time() - start_time:.2f}s")

    print("📘 [LOG] Loop 2 (ABRASMA): Processando 'Complementar'...")
    valores_comp = {}
    linhas_comp_count = 0
    for i, row in enumerate(ws_comp.iter_rows(min_row=2, values_only=True)):
        linhas_comp_count += 1

        eql = str(row[col_eql_comp - 1]).strip() if col_eql_comp <= len(row) and row[col_eql_comp - 1] else ""
        parc = str(row[col_parc_comp - 1]).strip() if col_parc_comp <= len(row) and row[col_parc_comp - 1] else ""
        total = normalizar_valor_repasse(row[col_total_comp - 1]) if col_total_comp <= len(row) else 0.0

        if eql and parc:
            chave_simples = (eql, parc)
            if chave_simples not in valores_comp:
                valores_comp[chave_simples] = total

    print(f"📗 [LOG] Fim Loop 2. 'Complementar' processada ({linhas_comp_count} linhas). {len(valores_comp)} chaves. Tempo: {time.time() - start_time:.2f}s")

    print("📘 [LOG] Loop 3 (ABRASMA): Comparando 'Anterior' com 'Complementar'...")
    iguais = []
    divergentes = []
    nao_encontrados_anterior = []
    duplicados_vistos = set()
    linhas_ant_loop3 = 0

    if ws_ant.max_row >= 2:
         for row_idx, row_cells in enumerate(ws_ant.iter_rows(min_row=2)):
             linhas_ant_loop3 += 1
             
             celula_eql = row_cells[col_eql_ant - 1] if col_eql_ant <= len(row_cells) else None
             celula_parc = row_cells[col_parc_ant - 1] if col_parc_ant <= len(row_cells) else None
             celula_total = row_cells[col_total_ant - 1] if col_total_ant <= len(row_cells) else None

             eql = str(celula_eql.value).strip() if celula_eql and celula_eql.value is not None else ""
             parc = str(celula_parc.value).strip() if celula_parc and celula_parc.value is not None else ""

             if not eql or not parc: continue

             total_ant = normalizar_valor_repasse(celula_total.value if celula_total else None)
             chave_simples = (eql, parc)
             chave_completa = (eql, parc, total_ant)

             if contagem_ant.get(chave_completa, 0) > 1:
                 if chave_completa in duplicados_vistos:
                     divergentes.append((row_cells, f"Duplicado na 'Anterior' (EQL {eql}, P {parc}, Total {total_ant:.2f})"))
                     continue
                 else:
                     duplicados_vistos.add(chave_completa)

             valor_comp = valores_comp.get(chave_simples)

             if valor_comp is None:
                 nao_encontrados_anterior.append((row_cells, f"Não encontrado na 'Complementar' (Valor Anterior={total_ant:.2f})"))
             elif abs(total_ant - valor_comp) <= 0.02:
                 iguais.append((row_cells, ""))
             else:
                 divergentes.append((row_cells, f"Valor diferente (Anterior={total_ant:.2f} / Complementar={valor_comp:.2f})"))
    else:
        print("[AVISO] Planilha 'Anterior' sem dados para Loop 3.")


    # --- INÍCIO DA CORREÇÃO ---
    print("📘 [LOG] Verificando itens da 'Complementar' ausentes na 'Anterior'...")
    nao_encontrados_comp_formatado = [] # Lista de tuplas (row_cells, status)
    items_comp_apenas = 0
    
    # Otimização: Achar as chaves que precisamos (Complementar - Anterior)
    chaves_comp_apenas = valores_comp.keys() - valores_ant.keys()
    
    if chaves_comp_apenas:
        print(f"📘 [LOG] {len(chaves_comp_apenas)} itens exclusivos da 'Complementar' encontrados. Re-iterando 'Complementar' para buscar dados da linha...")
        
        # Precisamos re-iterar a planilha 'Complementar' para obter os objetos 'row_cells'
        for row_cells_comp in ws_comp.iter_rows(min_row=2):
            # Extrai a chave (EQL, Parc) desta linha
            celula_eql = row_cells_comp[col_eql_comp - 1] if col_eql_comp <= len(row_cells_comp) else None
            celula_parc = row_cells_comp[col_parc_comp - 1] if col_parc_comp <= len(row_cells_comp) else None
            
            eql = str(celula_eql.value).strip() if celula_eql and celula_eql.value is not None else ""
            parc = str(celula_parc.value).strip() if celula_parc and celula_parc.value is not None else ""

            chave_simples = (eql, parc)

            # Se a chave desta linha é uma das que procuramos
            if chave_simples in chaves_comp_apenas:
                # Pega o valor total para incluir no status
                celula_total = row_cells_comp[col_total_comp - 1] if col_total_comp <= len(row_cells_comp) else None
                total_comp = normalizar_valor_repasse(celula_total.value if celula_total else None)
                
                status_msg = f"Presente na 'Complementar' (Valor={total_comp:.2f}), Ausente na 'Anterior'"
                
                # Adiciona a TUPLA DE CÉLULAS (row_cells_comp) e a mensagem de status
                nao_encontrados_comp_formatado.append((row_cells_comp, status_msg))
                
                # Otimização: remove a chave do set para não procurar mais
                chaves_comp_apenas.remove(chave_simples) 
                
                # Otimização: se já achamos todas, para de iterar a planilha
                if not chaves_comp_apenas:
                    break 
    
    items_comp_apenas = len(nao_encontrados_comp_formatado)
    # --- FIM DA CORREÇÃO ---

    print(f"📗 [LOG] Fim Comparação ABRASMA. {linhas_ant_loop3} linhas 'Anterior'. {len(nao_encontrados_anterior)} não encontradas. {items_comp_apenas} só na 'Complementar'. Tempo: {time.time() - start_time:.2f}s")

    print("📘 [LOG] Criando planilhas de saída (ABRASMA)...")
    iguais_stream = criar_planilha_saida(iguais, ws_ant, incluir_status=False)
    divergentes_stream = criar_planilha_saida(divergentes, ws_ant, incluir_status=True)
    # Combina as duas listas de "não encontrados"
    nao_encontrados_combinados = nao_encontrados_anterior + nao_encontrados_comp_formatado
    nao_encontrados_stream = criar_planilha_saida(nao_encontrados_combinados, ws_ant, incluir_status=True)

    timestamp_str = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
    pasta_saida = os.path.join(app.config['UPLOAD_FOLDER'], f"repasse_abrasma_{timestamp_str}")
    os.makedirs(pasta_saida, exist_ok=True)
    print(f"Pasta de saída criada: {pasta_saida}")

    try:
        salvar_stream_em_arquivo(iguais_stream, os.path.join(pasta_saida, "iguais.xlsx"))
        salvar_stream_em_arquivo(divergentes_stream, os.path.join(pasta_saida, "divergentes.xlsx"))
        salvar_stream_em_arquivo(nao_encontrados_stream, os.path.join(pasta_saida, "nao_encontrados.xlsx"))
        print(f"📗 [LOG] Arquivos Excel (ABRASMA) salvos na pasta: {pasta_saida}")
    except Exception as e_save:
         print(f"📕 [ERRO] Falha ao salvar arquivos Excel (ABRASMA) na pasta {pasta_saida}: {e_save}")
         raise

    count_nao_encontrados = len(nao_encontrados_combinados)
    print(f"✅ [LOG] Fim de processar_repasse (ABRASMA). Totais: Iguais={len(iguais)}, Divergentes={len(divergentes)}, Não Encontrados={count_nao_encontrados}. Tempo total: {time.time() - start_time:.2f}s")
    return pasta_saida, len(iguais), len(divergentes), count_nao_encontrados


@app.route('/repasse', methods=['POST'])
def repasse_file():
    """Rota para a conciliação PickMoney (Diário vs Sistema)"""
    print("\n--- RECEIVED REQUEST /repasse (PickMoney) ---")
    start_time_route = time.time()

    if 'diario_file' not in request.files or 'sistema_file' not in request.files:
        print("📕 [ERRO] Arquivos 'diario_file' ou 'sistema_file' faltando.")
        return manual_render_template('error.html', status_code=400,
            error_title="Arquivos faltando",
            error_message="Você precisa enviar os arquivos 'Diário' e 'Sistema' para a conciliação PickMoney.")

    file_diario = request.files['diario_file']
    file_sistema = request.files['sistema_file']

    if file_diario.filename == '' or file_sistema.filename == '':
        print("📕 [ERRO] Nomes dos arquivos Excel (PickMoney) estão vazios.")
        return manual_render_template('error.html', status_code=400,
            error_title="Arquivos faltando",
            error_message="Selecione os dois arquivos Excel (Diário e Sistema) para conciliar.")

    allowed_extensions = {'.xlsx', '.xlsm'}
    diario_ext = os.path.splitext(file_diario.filename)[1].lower()
    sistema_ext = os.path.splitext(file_sistema.filename)[1].lower()
    if diario_ext not in allowed_extensions or sistema_ext not in allowed_extensions:
         print(f"📕 [ERRO] Extensões de arquivo inválidas: {diario_ext}, {sistema_ext}")
         return manual_render_template('error.html', status_code=400,
             error_title="Tipo de Arquivo Inválido",
             error_message=f"Por favor, envie apenas arquivos Excel ({', '.join(allowed_extensions)}).")

    print(f"📘 [LOG] Recebidos (PickMoney): {file_diario.filename}, {file_sistema.filename}")

    try:
        diario_stream = io.BytesIO(file_diario.read())
        sistema_stream = io.BytesIO(file_sistema.read())
        print(f"📘 [LOG] Arquivos Excel (PickMoney) lidos em memória. Tempo: {time.time() - start_time_route:.2f}s")

        pasta_saida, count_iguais, count_divergentes, count_nao_encontrados = processar_repasse(diario_stream, sistema_stream)

        print(f"📘 [LOG] Processamento (PickMoney) concluído. Criando ZIP da pasta '{pasta_saida}'...")
        zip_stream = io.BytesIO()
        timestamp_str = os.path.basename(pasta_saida).replace('repasse_pickmoney_', '')

        zip_arcname_iguais = "iguais.xlsx"
        zip_arcname_divergentes = "divergentes.xlsx"
        zip_arcname_nao_encontrados = "nao_encontrados.xlsx"

        with zipfile.ZipFile(zip_stream, 'w', zipfile.ZIP_DEFLATED) as zf:
            path_iguais = os.path.join(pasta_saida, "iguais.xlsx")
            path_divergentes = os.path.join(pasta_saida, "divergentes.xlsx")
            path_nao_encontrados = os.path.join(pasta_saida, "nao_encontrados.xlsx")

            if os.path.exists(path_iguais): zf.write(path_iguais, arcname=zip_arcname_iguais)
            if os.path.exists(path_divergentes): zf.write(path_divergentes, arcname=zip_arcname_divergentes)
            if os.path.exists(path_nao_encontrados): zf.write(path_nao_encontrados, arcname=zip_arcname_nao_encontrados)

        zip_stream.seek(0)
        print(f"📗 [LOG] ZIP (PickMoney) criado em memória.")

        report_filename = f"repasse_pickmoney_conciliado_{timestamp_str}.zip"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
        try:
            with open(report_path, 'wb') as f:
                f.write(zip_stream.getvalue())
            print(f"📗 [LOG] Arquivo ZIP (PickMoney) salvo para download em {report_path}.")
        except Exception as e_save:
             print(f"📕 [ERRO] Erro ao salvar o arquivo ZIP (PickMoney) em {report_path}: {e_save}")
             raise e_save

        print("✅ [LOG] Enviando resposta (PickMoney) para 'repasse_results.html'")
        return manual_render_template('repasse_results.html',
            count_iguais=count_iguais,
            count_divergentes=count_divergentes,
            count_nao_encontrados=count_nao_encontrados,
            download_url=url_for('download_file', filename=report_filename)
        )

    except ValueError as ve:
         print(f"📕 [ERRO VALIDAÇÃO] {ve}")
         traceback.print_exc()
         return manual_render_template('error.html', status_code=400,
             error_title="Erro na Conciliação (PickMoney) - Colunas Não Encontradas",
             error_message=f"Verifique os nomes das colunas nas planilhas. Detalhes: {ve}")
    except Exception as e:
        print(f"📕 [ERRO FATAL] Erro inesperado na rota /repasse (PickMoney): {e}")
        traceback.print_exc()
        error_details = f"{type(e).__name__}: {e}"
        return manual_render_template('error.html', status_code=500,
            error_title="Erro inesperado na conciliação (PickMoney)",
            error_message=f"Ocorreu um erro grave durante a análise. Detalhes: {error_details}")


@app.route('/repasse_abrasma', methods=['POST'])
def repasse_abrasma_file():
    """Rota para a conciliação ABRASMA (Anterior vs Complementar)"""
    print("\n--- RECEIVED REQUEST /repasse_abrasma ---")
    start_time_route = time.time()

    if 'anterior_file' not in request.files or 'complementar_file' not in request.files:
        print("📕 [ERRO] Arquivos 'anterior_file' ou 'complementar_file' faltando.")
        return manual_render_template('error.html', status_code=400,
            error_title="Arquivos faltando",
            error_message="Você precisa enviar a 'Planilha Anterior' e a 'Planilha Complementar' para a conciliação ABRASMA.")

    file_ant = request.files['anterior_file']
    file_comp = request.files['complementar_file']

    if file_ant.filename == '' or file_comp.filename == '':
        print("📕 [ERRO] Nomes dos arquivos Excel (ABRASMA) estão vazios.")
        return manual_render_template('error.html', status_code=400,
            error_title="Arquivos faltando",
            error_message="Selecione os dois arquivos Excel (Anterior e Complementar) para conciliar.")

    allowed_extensions = {'.xlsx', '.xlsm'}
    ant_ext = os.path.splitext(file_ant.filename)[1].lower()
    comp_ext = os.path.splitext(file_comp.filename)[1].lower()
    if ant_ext not in allowed_extensions or comp_ext not in allowed_extensions:
         print(f"📕 [ERRO] Extensões de arquivo inválidas: {ant_ext}, {comp_ext}")
         return manual_render_template('error.html', status_code=400,
             error_title="Tipo de Arquivo Inválido",
             error_message=f"Por favor, envie apenas arquivos Excel ({', '.join(allowed_extensions)}).")

    print(f"📘 [LOG] Recebidos (ABRASMA): {file_ant.filename}, {file_comp.filename}")

    try:
        anterior_stream = io.BytesIO(file_ant.read())
        complementar_stream = io.BytesIO(file_comp.read())
        print(f"📘 [LOG] Arquivos Excel (ABRASMA) lidos em memória. Tempo: {time.time() - start_time_route:.2f}s")

        # Chama a NOVA função de processamento ABRASMA
        pasta_saida, count_iguais, count_divergentes, count_nao_encontrados = processar_repasse_abrasma(anterior_stream, complementar_stream)

        print(f"📘 [LOG] Processamento (ABRASMA) concluído. Criando ZIP da pasta '{pasta_saida}'...")
        zip_stream = io.BytesIO()
        timestamp_str = os.path.basename(pasta_saida).replace('repasse_abrasma_', '')

        zip_arcname_iguais = "iguais.xlsx"
        zip_arcname_divergentes = "divergentes.xlsx"
        zip_arcname_nao_encontrados = "nao_encontrados.xlsx"

        with zipfile.ZipFile(zip_stream, 'w', zipfile.ZIP_DEFLATED) as zf:
            path_iguais = os.path.join(pasta_saida, "iguais.xlsx")
            path_divergentes = os.path.join(pasta_saida, "divergentes.xlsx")
            path_nao_encontrados = os.path.join(pasta_saida, "nao_encontrados.xlsx")

            if os.path.exists(path_iguais): zf.write(path_iguais, arcname=zip_arcname_iguais)
            if os.path.exists(path_divergentes): zf.write(path_divergentes, arcname=zip_arcname_divergentes)
            if os.path.exists(path_nao_encontrados): zf.write(path_nao_encontrados, arcname=zip_arcname_nao_encontrados)

        zip_stream.seek(0)
        print(f"📗 [LOG] ZIP (ABRASMA) criado em memória.")

        report_filename = f"repasse_abrasma_conciliado_{timestamp_str}.zip"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
        try:
            with open(report_path, 'wb') as f:
                f.write(zip_stream.getvalue())
            print(f"📗 [LOG] Arquivo ZIP (ABRASMA) salvo para download em {report_path}.")
        except Exception as e_save:
             print(f"📕 [ERRO] Erro ao salvar o arquivo ZIP (ABRASMA) em {report_path}: {e_save}")
             raise e_save

        print("✅ [LOG] Enviando resposta (ABRASMA) para 'repasse_results.html'")
        return manual_render_template('repasse_results.html',
            count_iguais=count_iguais,
            count_divergentes=count_divergentes,
            count_nao_encontrados=count_nao_encontrados,
            download_url=url_for('download_file', filename=report_filename)
        )

    except ValueError as ve:
         print(f"📕 [ERRO VALIDAÇÃO ABRASMA] {ve}")
         traceback.print_exc()
         return manual_render_template('error.html', status_code=400,
             error_title="Erro na Conciliação (ABRASMA) - Colunas Não Encontradas",
             error_message=f"Verifique os nomes das colunas (EQL, Parc, Total Recebido). Detalhes: {ve}")
    except Exception as e:
        print(f"📕 [ERRO FATAL] Erro inesperado na rota /repasse_abrasma: {e}")
        traceback.print_exc()
        error_details = f"{type(e).__name__}: {e}"
        return manual_render_template('error.html', status_code=500,
            error_title="Erro inesperado na conciliação (ABRASMA)",
            error_message=f"Ocorreu um erro grave durante a análise. Detalhes: {error_details}")


@app.route('/download/<filename>')
def download_file(filename):
     safe_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
     normalized_safe_path = os.path.normpath(safe_path)
     normalized_upload_folder = os.path.normpath(app.config['UPLOAD_FOLDER'])

     # Adiciona 'os.sep' para garantir que não pegue pastas com nome parecido
     if not normalized_safe_path.startswith(normalized_upload_folder + os.sep) and normalized_safe_path != normalized_upload_folder :
         print(f" Tentativa de acesso a caminho inválido: {filename} (Normalizado: {normalized_safe_path} vs Base: {normalized_upload_folder})")
         return "Acesso negado.", 403

     if not os.path.exists(safe_path):
          print(f" Arquivo não encontrado para download: {filename}")
          return "Arquivo não encontrado.", 404

     print(f"Enviando arquivo para download: {filename}")
     return send_file(safe_path, as_attachment=True)


if __name__ == '__main__':
    print("Iniciando servidor Flask local...")
    port = int(os.environ.get('PORT', 8080))
    debug_mode = os.environ.get('FLASK_DEBUG') == '1'
    print(f"Executando em http://0.0.0.0:{port} (debug={debug_mode})")
    app.run(debug=debug_mode, host='0.0.0.0', port=port)
