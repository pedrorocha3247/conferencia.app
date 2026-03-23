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
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
import time

# ==== Constantes e Mapeamentos ====
DASHES = dict.fromkeys(map(ord, "\u2010\u2011\u2012\u2013\u2014\u2015\u2212"), "-")
HEADERS = (
    "Remessa para Conferência", "Página", "Banco", "IMOBILIARIOS", "Débitos do Mês",
    "Vencimento", "Lançamentos", "Programação", "Carta", "DÉBITOS", "ENCARGOS",
    "PAGAMENTO", "TOTAL", "Limite p/", "TOTAL A PAGAR", "PAGAMENTO EFETUADO", "DESCONTO"
)
PADRAO_LOTE = re.compile(r"\b(\d{2,4}\.([A-Z0-9\u0399\u039A]{2})\.\d{1,4})\b")
PADRAO_PARCELA_MESMA_LINHA = re.compile(
    r"^(?!(?:DÉBITOS|ENCARGOS|DESCONTO|PAGAMENTO|TOTAL(?!\s*A PAGAR)|Limite p/))\s*"
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
    "RSCIV": {"Melhoramentos": 324.20, "Fundo de Transporte": 9.00},
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
UPLOAD_FOLDER_PATH = os.path.join(app.root_path, 'uploads')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER_PATH
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def manual_render_template(template_name, status_code=200, **kwargs):
    template_path = os.path.join(app.root_path, 'templates', template_name)
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        for key, value in kwargs.items():
            placeholder = f"__{key.upper()}__"
            if isinstance(value, str) and value.startswith('{') and value.endswith('}'):
                 html_content = html_content.replace(f'"{placeholder}"', value)
            else:
                 html_content = html_content.replace(placeholder, str(value))
        response = make_response(html_content)
        response.headers['Content-Type'] = 'text/html'
        return response, status_code
    except Exception as e:
        error_html = f"<html><body><h1>Erro 500: Template não encontrado</h1><p>{e}</p></body></html>"
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
             for page in doc:
                 texto_pagina = page.get_text("text", sort=True)
                 texto_completo += texto_pagina + "\n"
        return normalizar_texto(texto_completo)
    except Exception as e:
        traceback.print_exc()
        return ""

def normalizar_valor(valor):
    if valor is None: return 0.0
    if isinstance(valor, (int, float)): return round(float(valor), 2)
    s_norm = str(valor).strip().replace("R$", "").replace(" ", "").replace("\xa0", "")
    has_comma = "," in s_norm
    has_dot = "." in s_norm
    if has_comma and has_dot:
        if s_norm.rfind(',') > s_norm.rfind('.'):
             s_norm = s_norm.replace(".", "").replace(",", ".")
        else:
             s_norm = s_norm.replace(",", "")
    elif has_comma:
        s_norm = s_norm.replace(",", ".")
    elif has_dot and s_norm.count('.') > 1:
            parts = s_norm.split('.')
            s_norm = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return round(float(s_norm), 2)
    except:
        return 0.0

def fixos_do_emp(emp: str, modo_separacao: str):
    if modo_separacao == 'boleto':
        if emp not in EMP_MAP: return BASE_FIXOS
        f = dict(BASE_FIXOS)
        if EMP_MAP.get(emp):
            if "Melhoramentos" in EMP_MAP[emp]: f["Melhoramentos"] = [float(EMP_MAP[emp]["Melhoramentos"])]
            if "Fundo de Transporte" in EMP_MAP[emp]: f["Fundo de Transporte"] = [float(EMP_MAP[emp]["Fundo de Transporte"])]
        return f
    elif modo_separacao == 'ccb_realiza': return BASE_FIXOS_CCB
    return BASE_FIXOS

def detectar_emp_por_nome_arquivo(path: str):
    if not path: return None
    nome = os.path.splitext(os.path.basename(path))[0].upper()
    for k in EMP_MAP.keys():
        if nome.endswith("_" + k) or nome.endswith(k): return k
    if "SBRR" in nome: return "SBRR"
    return None

def detectar_emp_por_lote(lote: str):
    if not lote or "." not in lote: return "NAO_CLASSIFICADO"
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
        if texto_bloco: blocos.append((lote_atual, texto_bloco))
    return blocos

def tentar_nome_cliente(bloco: str) -> str:
    linhas = bloco.split('\n')
    if not linhas: return "Nome não localizado"
    for linha in linhas[:6]:
        linha_sem_lote = PADRAO_LOTE.sub('', linha).strip()
        if not linha_sem_lote: continue
        if len(linha_sem_lote) > 5 and ' ' in linha_sem_lote and not any(h.upper() in linha_sem_lote.upper() for h in HEADERS if h):
            return linha_sem_lote
    return "Nome não localizado"

def extrair_parcelas(bloco: str):
    itens = OrderedDict()
    pos_lancamentos = bloco.find("Lançamentos")
    bloco_de_trabalho = bloco[pos_lancamentos:] if pos_lancamentos != -1 else bloco
    linhas_originais = bloco_de_trabalho.splitlines()
    ignorar_proxima = False
    for i, linha in enumerate(linhas_originais):
        if ignorar_proxima:
            ignorar_proxima = False
            continue
        linha_limpa = linha.strip()
        if not linha_limpa or any(h.strip().upper() == linha_limpa.upper() for h in ["Lançamentos", "Débitos do Mês"]):
            continue
        match_mesma = PADRAO_PARCELA_MESMA_LINHA.match(linha_limpa)
        if match_mesma:
            itens[limpar_rotulo(match_mesma.group(1))] = normalizar_valor(match_mesma.group(2))
            continue
        if any(c.isalpha() for c in linha_limpa):
            j = i + 1
            while j < len(linhas_originais) and not linhas_originais[j].strip(): j += 1
            if j < len(linhas_originais):
                match_num = PADRAO_NUMERO_PURO.match(linhas_originais[j].strip())
                if match_num:
                    itens[limpar_rotulo(linha_limpa)] = normalizar_valor(match_num.group(1))
                    ignorar_proxima = True
    return itens

def processar_pdf_validacao(texto_pdf: str, modo_separacao: str, emp_fixo_boleto: str = None):
    blocos = fatiar_blocos(texto_pdf)
    if not blocos: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    linhas_todas, linhas_cov, linhas_div = [], [], []
    for lote, bloco in blocos:
        emp_atual = (detectar_emp_por_lote(lote) if emp_fixo_boleto == "SBRR" else emp_fixo_boleto) if modo_separacao == 'boleto' else detectar_emp_por_lote(lote)
        cliente, itens = tentar_nome_cliente(bloco), extrair_parcelas(bloco)
        VALORES_CORRETOS = fixos_do_emp(emp_atual, modo_separacao)
        for rot, val in itens.items():
            linhas_todas.append({"Empreendimento": emp_atual, "Lote": lote, "Cliente": cliente, "Parcela": rot, "Valor": val})
        cov = {"Empreendimento": emp_atual, "Lote": lote, "Cliente": cliente}
        for k in VALORES_CORRETOS.keys(): cov[k] = itens.get(k)
        vistos = [k for k in VALORES_CORRETOS if cov[k] is not None]
        cov["QtdParc_Alvo"], cov["Parc_Alvo"] = len(vistos), ", ".join(vistos)
        linhas_cov.append(cov)
        if modo_separacao != 'ccb_realiza':
            for rot in vistos:
                val = cov[rot]
                permitidos = VALORES_CORRETOS.get(rot, [])
                if permitidos and all(abs(val - v) > 1e-6 for v in permitidos):
                    linhas_div.append({"Empreendimento": emp_atual, "Lote": lote, "Cliente": cliente, "Parcela": rot, "Valor no Documento": float(val), "Valor Correto": " ou ".join(f"{v:.2f}" for v in permitidos)})
    return pd.DataFrame(linhas_todas), pd.DataFrame(linhas_cov), pd.DataFrame(linhas_div)

def processar_comparativo(texto_anterior, texto_atual, modo_separacao, emp_fixo_boleto):
    df_ant_raw, _, _ = processar_pdf_validacao(texto_anterior, modo_separacao, emp_fixo_boleto)
    df_atu_raw, _, _ = processar_pdf_validacao(texto_atual, modo_separacao, emp_fixo_boleto)
    
    def get_totais(df):
        return df[df['Parcela'].str.strip().str.upper() == 'TOTAL A PAGAR'][['Empreendimento', 'Lote', 'Cliente', 'Valor']]

    df_totais_ant = get_totais(df_ant_raw).rename(columns={'Valor': 'Total Anterior'})
    df_totais_atu = get_totais(df_atu_raw).rename(columns={'Valor': 'Total Atual'})

    p_rem = ['TOTAL A PAGAR', 'DESCONTO', 'DÉBITOS DO MÊS', 'DÉBITOS DO MÊS ANTERIOR', 'ENCARGOS POR ATRASO', 'PAGAMENTO EFETUADO']
    df_ant = df_ant_raw[~df_ant_raw['Parcela'].str.strip().str.upper().isin(p_rem)].copy()
    df_atu = df_atu_raw[~df_atu_raw['Parcela'].str.strip().str.upper().isin(p_rem)].copy()

    df_comp = pd.merge(df_ant.rename(columns={'Valor': 'Valor Anterior'}), df_atu.rename(columns={'Valor': 'Valor Atual'}), on=['Empreendimento', 'Lote', 'Cliente', 'Parcela'], how='outer')
    
    lotes_ant = df_ant_raw[['Empreendimento', 'Lote', 'Cliente']].drop_duplicates()
    lotes_atu = df_atu_raw[['Empreendimento', 'Lote', 'Cliente']].drop_duplicates()
    lotes_m = pd.merge(lotes_ant, lotes_atu, on=['Empreendimento', 'Lote', 'Cliente'], how='outer', indicator=True)

    df_adic = pd.merge(lotes_m[lotes_m['_merge'] == 'right_only'].drop('_merge', axis=1), df_totais_atu, how='left')
    df_remov = pd.merge(lotes_m[lotes_m['_merge'] == 'left_only'].drop('_merge', axis=1), df_totais_ant, how='left')
    
    df_div = df_comp[(df_comp['Valor Anterior'].notna()) & (df_comp['Valor Atual'].notna()) & (abs(df_comp['Valor Anterior'] - df_comp['Valor Atual']) > 0.025)].copy()
    if not df_div.empty: df_div['Diferença'] = df_div['Valor Atual'] - df_div['Valor Anterior']

    df_resumo = pd.DataFrame({
        ' ': ['Lotes Mês Anterior', 'Lotes Mês Atual', 'Lotes Adicionados', 'Lotes Removidos', 'Parcelas com Valor Alterado'],
        'LOTES': [len(lotes_ant), len(lotes_atu), len(df_adic), len(df_remov), df_div['Lote'].nunique() if not df_div.empty else 0],
        'TOTAIS': [df_totais_ant['Total Anterior'].sum(), df_totais_atu['Total Atual'].sum(), df_adic['Total Atual'].sum(), df_remov['Total Anterior'].sum(), df_div['Diferença'].sum() if not df_div.empty else 0]
    })
    return df_resumo, df_adic, df_remov, df_div, df_comp[df_comp['Valor Anterior'].isna()], df_comp[df_comp['Valor Atual'].isna()]

def formatar_excel(output_stream, dfs: dict):
    with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
        for sheet_name, df in dfs.items():
            if df is not None: df.to_excel(writer, index=False, sheet_name=sheet_name)
        num_s = NamedStyle(name='br_num', number_format='#,##0.00')
        for sheet in writer.sheets.values():
            sheet.sheet_view.showGridLines = False
            for col in sheet.columns:
                max_l = max([len(str(cell.value)) for cell in col if cell.value] + [10])
                sheet.column_dimensions[get_column_letter(col[0].column)].width = min(max_l + 2, 50)
                for cell in col[1:]:
                    if isinstance(cell.value, (int, float)): cell.style = num_s
            sheet.auto_filter.ref = sheet.dimensions

# ==== ROTAS FLASK ====

@app.route('/')
def index(): return manual_render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files.get('pdf_file')
    if not file or file.filename == '': return manual_render_template('error.html', error_title="Arquivo ausente", error_message="Selecione um PDF.")
    modo = request.form.get('modo_separacao', 'boleto')
    emp_f = detectar_emp_por_nome_arquivo(file.filename) if modo == 'boleto' else None
    if modo == 'boleto' and not emp_f: return manual_render_template('error.html', error_title="Erro de Nome", error_message="Arquivo deve conter código do empreendimento.")
    
    pdf_text = extrair_texto_pdf(file.read())
    df_t, df_c, df_d = processar_pdf_validacao(pdf_text, modo, emp_f)
    
    output = io.BytesIO()
    formatar_excel(output, {"Divergencias": df_d, "Cobertura": df_c, "Todas_Parcelas": df_t})
    output.seek(0)
    
    fname = f"relatorio_{modo}_{int(time.time())}.xlsx"
    with open(os.path.join(app.config['UPLOAD_FOLDER'], fname), 'wb') as f: f.write(output.getvalue())
    
    return manual_render_template('results.html', total_lotes=len(df_c), total_divergencias=len(df_d), download_url=url_for('download_file', filename=fname), modo_usado=modo.upper())

@app.route('/compare', methods=['POST'])
def compare_files():
    f_ant = request.files.get('pdf_mes_anterior')
    f_atu = request.files.get('pdf_mes_atual')
    if not f_ant or not f_atu: return manual_render_template('error.html', error_message="Envie os dois PDFs.")
    modo = request.form.get('modo_separacao_comp', 'boleto')
    
    df_r, df_a, df_rm, df_dv, df_n, df_rx = processar_comparativo(f_ant.read(), f_atu.read(), modo, detectar_emp_por_nome_arquivo(f_ant.filename))
    
    output = io.BytesIO()
    formatar_excel(output, {"Resumo": df_r, "Adicionados": df_a, "Removidos": df_rm, "Divergencias": df_dv})
    output.seek(0)
    
    fname = f"comparativo_{int(time.time())}.xlsx"
    with open(os.path.join(app.config['UPLOAD_FOLDER'], fname), 'wb') as f: f.write(output.getvalue())
    
    return manual_render_template('compare_results.html', download_url=url_for('download_file', filename=fname), modo_usado=modo.upper())

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))
