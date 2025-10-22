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

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

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

def fixos_do_emp(emp: str):
    if emp not in EMP_MAP:
        return BASE_FIXOS
    f = dict(BASE_FIXOS)
    if EMP_MAP.get(emp) and EMP_MAP[emp].get("Melhoramentos") is not None:
        f["Melhoramentos"] = [float(EMP_MAP[emp]["Melhoramentos"])]
    if EMP_MAP.get(emp) and EMP_MAP[emp].get("Fundo de Transporte") is not None:
        f["Fundo de Transporte"] = [float(EMP_MAP[emp]["Fundo de Transporte"])]
    return f

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
            emp_atual = detectar_emp_por_lote(lote)
        cliente = tentar_nome_cliente(bloco)
        itens = extrair_parcelas(bloco)
        VALORES_CORRETOS = fixos_do_emp(emp_atual)
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
        for rot in vistos:
            val = cov[rot]
            if val is None: continue
            permitidos = VALORES_CORRETOS.get(rot, [])
            if all(abs(val - v) > 1e-6 for v in permitidos):
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
    
    total_adicionados_valor = df_adicionados['Total Atual'].sum()
    total_removidos_valor = df_removidos['Total Anterior'].sum()
    total_divergencias_valor = df_divergencias['Diferença'].sum()
    total_mes_anterior_valor = df_totais_ant['Total Anterior'].sum()
    total_mes_atual_valor = df_totais_atu['Total Atual'].sum()

    resumo_financeiro_data = {
        ' ': ['Lotes Mês Anterior', 'Lotes Mês Atual', 'Lotes Adicionados', 'Lotes Removidos', 'Parcelas com Valor Alterado'],
        'LOTES': [len(lotes_ant), len(lotes_atu), len(df_adicionados), len(df_removidos), len(df_divergencias)],
        'TOTAIS': [total_mes_anterior_valor, total_mes_atual_valor, total_adicionados_valor, total_removidos_valor, total_divergencias_valor]
    }
    df_resumo_completo = pd.DataFrame(resumo_financeiro_data)
    
    return df_resumo_completo, df_adicionados, df_removidos, df_divergencias, df_parcelas_novas, df_parcelas_removidas

def formatar_excel(output_stream, dfs: dict):
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
        for sheet_name, df in dfs.items():
            if not df.empty:
                if sheet_name == "Resumo":
                    df.to_excel(writer, index=False, sheet_name=sheet_name)
                else:
                    df.to_excel(writer, index=False, sheet_name=sheet_name)
        
        number_style = NamedStyle(name='br_number_style', number_format='#,##0.00')
        integer_style = NamedStyle(name='br_integer_style', number_format='0')

        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            worksheet.sheet_view.showGridLines = False
            
            for column_cells in worksheet.columns:
                max_length = 0
                column = column_cells[0].column_letter
                for cell in column_cells:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                    
                    if isinstance(cell.value, float):
                        cell.style = number_style
                    elif isinstance(cell.value, int):
                         if sheet_name == 'Resumo' and column == 'B':
                            cell.style = integer_style
                         elif column != 'B':
                            cell.style = number_style

                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column].width = adjusted_width
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
    if origem and origem.has_style:
        destino.font = copy(origem.font)
        destino.border = copy(origem.border)
        destino.fill = copy(origem.fill)
        destino.number_format = copy(origem.number_format)
        destino.protection = copy(origem.protection)
        destino.alignment = copy(origem.alignment)

def achar_coluna(sheet, nome_coluna):
    for cell in sheet[1]:
        if str(cell.value).strip().lower() == nome_coluna.lower():
            return cell.column
    return None

# ===================================================================
# FUNÇÃO CORRIGIDA
# ===================================================================
def criar_planilha_saida(linhas, ws_diario, incluir_status=False):
    wb_out = Workbook()
    ws_out = wb_out.active

    # Copia o cabeçalho
    for i, cell in enumerate(ws_diario[1], 1):
        novo = ws_out.cell(row=1, column=i, value=cell.value)
        if cell:
            copiar_formatacao(cell, novo)
        ws_out.column_dimensions[openpyxl.utils.get_column_letter(i)].width = ws_diario.column_dimensions[
            openpyxl.utils.get_column_letter(i)
        ].width

    col_status = 0
    if incluir_status:
        col_status = len(list(ws_diario[1])) + 1
        ws_out.cell(row=1, column=col_status, value="Status")

    linha_out = 2
    for linha_info in linhas:
        # CORREÇÃO 1: 'linhas' sempre contém 2-tuplas (dados, status)
        # Removemos o if/else problemático daqui.
        linha, status = linha_info 

        if linha is None:
            if incluir_status:
                ws_out.cell(row=linha_out, column=col_status, value=status)
            linha_out += 1
            continue

        # CORREÇÃO 2: Loop robusto sugerido por você
        for i, cell in enumerate(linha, 1):
            try:
                # Se 'cell' for um obj Cell, usa .value. Se for um valor puro (str, int), usa 'cell'
                valor = cell.value if hasattr(cell, "value") else cell
                novo = ws_out.cell(row=linha_out, column=i, value=valor)
                # Só copia formatação se 'cell' for um obj Cell
                if hasattr(cell, "value"):
                    copiar_formatacao(cell, novo)
            except Exception as e:
                print(f"[Aviso] Erro ao copiar célula {i} da linha {linha_out}: {e}")

        if incluir_status:
            ws_out.cell(row=linha_out, column=col_status, value=status)

        linha_out += 1

    if incluir_status:
        ws_out.cell(row=linha_out + 1, column=1, value=f"Total divergentes: {len(linhas)}")

    stream_out = io.BytesIO()
    wb_out.save(stream_out)
    stream_out.seek(0)
    return stream_out
# ===================================================================
# FIM DA FUNÇÃO CORRIGIDA
# ===================================================================

def processar_repasse(diario_stream, sistema_stream):
    wb_diario = load_workbook(diario_stream, data_only=True)
    ws_diario = wb_diario.worksheets[0]

    wb_sistema = load_workbook(sistema_stream, data_only=True)
    ws_sistema = wb_sistema.worksheets[0]

    col_eq_diario = achar_coluna(ws_diario, "EQL")
    col_parcela_diario = achar_coluna(ws_diario, "Parcela")
    col_principal_diario = 4
    col_corrmonet_diario = 9

    col_eq_sistema = achar_coluna(ws_sistema, "EQL")
    col_parcela_sistema = achar_coluna(ws_sistema, "Parcela")
    col_valor_sistema = achar_coluna(ws_sistema, "Valor")

    if not all([col_eq_diario, col_parcela_diario, col_principal_diario, col_corrmonet_diario, col_eq_sistema, col_parcela_sistema, col_valor_sistema]):
        raise ValueError("Não foi possível encontrar todas as colunas necessárias (EQL, Parcela, Valor, etc.)")

    valores_diario = {}
    contagem_diario = {}
    for row in ws_diario.iter_rows(min_row=2, values_only=True):
        eql = str(row[col_eq_diario - 1]).strip() if row[col_eq_diario - 1] else ""
        parcela = str(row[col_parcela_diario - 1]).strip() if row[col_parcela_diario - 1] else ""
        principal = normalizar_valor_repasse(row[col_principal_diario - 1]) if len(row) >= col_principal_diario else 0.0
        correcao = normalizar_valor_repasse(row[col_corrmonet_diario - 1]) if len(row) >= col_corrmonet_diario else 0.0
        total = round(principal + correcao, 2)

        if eql and parcela:
            chave_completa = (eql, parcela, principal, correcao)
            chave_simples = (eql, parcela)
            contagem_diario[chave_completa] = contagem_diario.get(chave_completa, 0) + 1
            if chave_simples not in valores_diario:
                valores_diario[chave_simples] = total

    valores_sistema = {}
    for row in ws_sistema.iter_rows(min_row=2, values_only=True):
        eql = str(row[col_eq_sistema - 1]).strip() if row[col_eq_sistema - 1] else ""
        parcela = str(row[col_parcela_sistema - 1]).strip() if row[col_parcela_sistema - 1] else ""
        valor = normalizar_valor_repasse(row[col_valor_sistema - 1])

        if eql and parcela:
            chave_simples = (eql, parcela)
            if chave_simples not in valores_sistema:
                valores_sistema[chave_simples] = valor

    iguais = []
    divergentes = []
    duplicados_vistos = set()

    # Itera sobre as CÉLULAS (objetos) para poder copiar a formatação
    for row in ws_diario.iter_rows(min_row=2):
        celula_eql = row[col_eq_diario - 1]
        celula_parcela = row[col_parcela_diario - 1]
        celula_principal = row[col_principal_diario - 1]
        celula_correcao = row[col_corrmonet_diario - 1]

        eql = str(celula_eql.value).strip() if celula_eql and celula_eql.value else ""
        parcela = str(celula_parcela.value).strip() if celula_parcela and celula_parcela.value else ""
        principal = normalizar_valor_repasse(celula_principal.value if celula_principal else None)
        correcao = normalizar_valor_repasse(celula_correcao.value if celula_correcao else None)
        
        chave_simples = (eql, parcela)
        chave_completa = (eql, parcela, principal, correcao)

        if not eql or not parcela:
            continue

        if contagem_diario.get(chave_completa, 0) > 1:
            if chave_completa not in duplicados_vistos:
                duplicados_vistos.add(chave_completa)
            else:
                # 'row' é uma tupla de Células
                divergentes.append((row, f"EQL {eql} Parcela {parcela} duplicada no diário (Principal={principal}, Correção={correcao})"))
                continue

        valor_diario = valores_diario.get(chave_simples, 0.0)
        valor_sistema = valores_sistema.get(chave_simples)

        if valor_sistema is None:
            # 'row' é uma tupla de Células
            divergentes.append((row, f"EQL {eql} Parcela {parcela} não encontrada no sistema"))
        elif abs(valor_diario - valor_sistema) <= 0.02:
            # 'row' é uma tupla de Células
            iguais.append((row, ""))
        else:
            # 'row' é uma tupla de Células
            divergentes.append((row, f"Valor diferente (Diário={valor_diario:.2f} / Sistema={valor_sistema:.2f})"))

    for chave in valores_sistema:
        if chave not in valores_diario:
            eql, parcela = chave
            # 'None' aqui é o 'linha' (não há linha no diário)
            divergentes.append((None, f"EQL {eql} Parcela {parcela} presente no sistema, ausente no diário"))

    iguais_stream = criar_planilha_saida(iguais, ws_diario, incluir_status=False)
    divergentes_stream = criar_planilha_saida(divergentes, ws_diario, incluir_status=True)

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
    modo_separacao = request.form.get('modo_separacao', 'boleto')

    try:
        emp_fixo = None
        if modo_separacao == 'boleto':
            emp_fixo = detectar_emp_por_nome_arquivo(file.filename)
            if not emp_fixo:
                error_msg = ("Para o modo 'Boleto', o nome do arquivo precisa terminar com um código de empreendimento (ex: 'Extrato_IATE.pdf'). "
                             "Você pode ter selecionado o modo de análise errado para este tipo de arquivo.")
                return manual_render_template('error.html', status_code=400,
                    error_title="Empreendimento não identificado", error_message=error_msg)
        
        elif modo_separacao == 'debito_credito':
            if detectar_emp_por_nome_arquivo(file.filename):
                error_msg = ("Este arquivo é do tipo 'Boleto', mas o modo 'Débito/Crédito' foi selecionado. "
                             "Por favor, volte e selecione o modo de análise correto para este arquivo.")
                return manual_render_template('error.html', status_code=400,
                                              error_title="Modo de Análise Incorreto",
                                              error_message=error_msg)

        pdf_stream = file.read()
        texto_pdf = extrair_texto_pdf(pdf_stream)
        if not texto_pdf:
            return manual_render_template('error.html', status_code=500,
                error_title="Erro ao ler o PDF", 
                error_message="Não foi possível extrair o texto do arquivo enviado. Ele pode estar corrompido ou ser uma imagem.")

        df_todas_raw, df_cov, df_div = processar_pdf_validacao(texto_pdf, modo_separacao, emp_fixo)
        
        df_todas = df_todas_raw.copy()
        if not df_todas.empty:
            parcelas_para_remover = ['TOTAL A PAGAR', 'DESCONTO', 'DÉBITOS DO MÊS']
            df_todas = df_todas[~df_todas['Parcela'].str.strip().str.upper().isin(parcelas_para_remover)]
            df_todas = df_todas[~df_todas['Parcela'].str.strip().str.upper().str.startswith('TOTAL BANCO')]
        
        output = io.BytesIO()
        dfs_to_excel = {"Divergencias": df_div, "Cobertura_Analise": df_cov, "Todas_Parcelas_Extraidas": df_todas}
        formatar_excel(output, dfs_to_excel)
        output.seek(0)

        base_name = os.path.splitext(file.filename)[0]
        report_filename = f"relatorio_{base_name}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
        
        with open(report_path, 'wb') as f: f.write(output.getvalue())

        nao_classificados = 0
        if not df_cov.empty and 'Empreendimento' in df_cov.columns:
            nao_classificados = df_cov[df_cov['Empreendimento'] == 'NAO_CLASSIFICADO'].shape[0]

        return manual_render_template('results.html',
            divergencias_json=df_div.to_json(orient='split', index=False) if not df_div.empty else 'null',
            total_lotes=len(df_cov),
            total_divergencias=len(df_div),
            nao_classificados=nao_classificados,
            download_url=url_for('download_file', filename=report_filename),
            modo_usado=modo_separacao.replace('_', '/')
        )
    
    except Exception as e:
        traceback.print_exc()
        return manual_render_template('error.html', status_code=500,
            error_title="Erro inesperado no processamento", 
            error_message=f"Ocorreu um erro grave durante a análise do arquivo. Detalhes: {e}")

@app.route('/compare', methods=['POST'])
def compare_files():
    if 'pdf_mes_anterior' not in request.files or 'pdf_mes_atual' not in request.files:
        return manual_render_template('error.html', status_code=400,
            error_title="Arquivos faltando", 
            error_message="Ambos os arquivos (mês anterior e atual) são necessários para a comparação.")

    file_ant = request.files['pdf_mes_anterior']
    file_atu = request.files['pdf_mes_atual']
    modo_separacao = request.form.get('modo_separacao_comp', 'boleto')

    if file_ant.filename == '' or file_atu.filename == '':
        return manual_render_template('error.html', status_code=400,
            error_title="Arquivos faltando", 
            error_message="Selecione os dois arquivos para comparar.")

    try:
        emp_fixo_boleto = None
        if modo_separacao == 'boleto':
            emp_ant = detectar_emp_por_nome_arquivo(file_ant.filename)
            emp_atu = detectar_emp_por_nome_arquivo(file_atu.filename)
            if not emp_ant or not emp_atu:
                return manual_render_template('error.html', status_code=400,
                    error_title="Empreendimento não identificado",
                    error_message="Para o modo 'Boleto', o nome de ambos os arquivos precisa terminar com um código de empreendimento.")
            if emp_ant != emp_atu:
                return manual_render_template('error.html', status_code=400,
                    error_title="Empreendimentos diferentes",
                    error_message="Para o modo 'Boleto', os arquivos devem ser do mesmo empreendimento.")
            emp_fixo_boleto = emp_ant
        elif modo_separacao == 'debito_credito':
            if detectar_emp_por_nome_arquivo(file_ant.filename) or detectar_emp_por_nome_arquivo(file_atu.filename):
                error_msg = ("Um dos arquivos parece ser do tipo 'Boleto', mas o modo 'Débito/Crédito' foi selecionado. "
                             "Por favor, volte e selecione o modo de análise correto.")
                return manual_render_template('error.html', status_code=400,
                                              error_title="Modo de Análise Incorreto",
                                              error_message=error_msg)

        texto_ant = extrair_texto_pdf(file_ant.read())
        texto_atu = extrair_texto_pdf(file_atu.read())

        if not texto_ant or not texto_atu:
            return manual_render_template('error.html', status_code=500,
                error_title="Erro ao ler o PDF",
                error_message="Não foi possível extrair texto de um dos PDFs. Ele pode estar corrompido ou ser uma imagem.")

        df_resumo_completo, df_adicionados, df_removidos, df_divergencias, df_parcelas_novas, df_parcelas_removidas = processar_comparativo(
            texto_ant, texto_atu, modo_separacao, emp_fixo_boleto
        )

        output = io.BytesIO()
        dfs_to_excel = {
            "Resumo": df_resumo_completo, "Lotes Adicionados": df_adicionados, "Lotes Removidos": df_removidos,
            "Divergências de Valor": df_divergencias, "Parcelas Novas por Lote": df_parcelas_novas,
            "Parcelas Removidas por Lote": df_parcelas_removidas,
        }
        formatar_excel(output, dfs_to_excel)
        output.seek(0)
        
        report_filename = f"comparativo_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
        with open(report_path, 'wb') as f: f.write(output.getvalue())
        
        resumo_dict = pd.Series(df_resumo_completo.set_index(' ')['LOTES']).to_dict()

        return manual_render_template('compare_results.html',
            resumo_lotes_mes_anterior=resumo_dict.get('Lotes Mês Anterior', 0),
            resumo_lotes_mes_atual=resumo_dict.get('Lotes Mês Atual', 0),
            resumo_lotes_adicionados=resumo_dict.get('Lotes Adicionados', 0),
            resumo_lotes_removidos=resumo_dict.get('Lotes Removidos', 0),
            resumo_parcelas_com_valor_alterado=resumo_dict.get('Parcelas com Valor Alterado', 0),
            divergencias_json=df_divergencias.to_json(orient='split', index=False) if not df_divergencias.empty else 'null',
            adicionados_json=df_adicionados.to_json(orient='split', index=False) if not df_adicionados.empty else 'null',
            removidos_json=df_removidos.to_json(orient='split', index=False) if not df_removidos.empty else 'null',
            download_url=url_for('download_file', filename=report_filename)
        )

    except Exception as e:
        traceback.print_exc()
        return manual_render_template('error.html', status_code=500,
            error_title="Erro inesperado na comparação", 
            error_message=f"Ocorreu um erro grave durante a comparação dos arquivos. Detalhes: {e}")

@app.route('/repasse', methods=['POST'])
def repasse_file():
    if 'diario_file' not in request.files or 'sistema_file' not in request.files:
        return manual_render_template('error.html', status_code=400,
            error_title="Arquivos faltando", 
            error_message="Você precisa enviar os dois arquivos (Diário e Sistema) para a conciliação.")

    file_diario = request.files['diario_file']
    file_sistema = request.files['sistema_file']

    if file_diario.filename == '' or file_sistema.filename == '':
        return manual_render_template('error.html', status_code=400,
            error_title="Arquivos faltando", 
            error_message="Selecione os dois arquivos para conciliar.")
    
    try:
        diario_stream = io.BytesIO(file_diario.read())
        sistema_stream = io.BytesIO(file_sistema.read())

        iguais_stream, divergentes_stream, count_iguais, count_divergentes = processar_repasse(diario_stream, sistema_stream)

        zip_stream = io.BytesIO()
        with zipfile.ZipFile(zip_stream, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr('iguais.xlsx', iguais_stream.getvalue())
            zf.writestr('divergentes.xlsx', divergentes_stream.getvalue())
        zip_stream.seek(0)
        
        report_filename = f"repasse_conciliado_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.zip"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
        with open(report_path, 'wb') as f: f.write(zip_stream.getvalue())

        return manual_render_template('repasse_results.html',
            count_iguais=count_iguais,
            count_divergentes=count_divergentes,
            download_url=url_for('download_file', filename=report_filename)
        )

    except Exception as e:
        traceback.print_exc()
        # Captura o erro 'tuple' object... se ele ocorrer aqui
        if "'tuple' object has no attribute 'value'" in str(e):
             error_message = f"Ocorreu um erro grave durante a análise. Detalhes: {e}. Isso indica que uma linha na planilha 'Diário' pode estar em um formato inesperado. Verifique a função 'criar_planilha_saida' no app.py."
        else:
             error_message = f"Ocorreu um erro grave durante a análise. Detalhes: {e}"
        
        return manual_render_template('error.html', status_code=500,
            error_title="Erro inesperado na conciliação", 
            error_message=error_message)

@app.route('/download/<filename>')
def download_file(filename):
    path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    return send_file(path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=int(os.environ.get('PORT', 8080)))
