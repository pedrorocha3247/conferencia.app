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
import traceback

# ==== Constantes e Mapeamentos ====

DASHES = dict.fromkeys(map(ord, "\u2010\u2011\u2012\u2013\u2014\u2015\u2212"), "-")
HEADERS = (
    "Remessa para Conferência", "Página", "Banco", "IMOBILIARIOS", "Débitos do Mês",
    "Vencimento", "Lançamentos", "Programação", "Carta", "DÉBITOS", "ENCARGOS",
    "PAGAMENTO", "TOTAL", "Limite p/", "TOTAL A PAGAR", "PAGAMENTO EFETUADO", "DESCONTO"
)

PADRAO_LOTE = re.compile(r"\b(\d{2,4}\.(?:[A-Z\u0399\u039A]{2}|\d{2})\.\d{1,4})\b")

PADRAO_PARCELA_MESMA_LINHA = re.compile(
    r"^(?!(?:DÉBITOS|ENCARGOS|DESCONTO|PAGAMENTO|TOTAL|Limite p/))\s*"
    r"([A-Za-zÀ-ú][A-Za-zÀ-ú\s\.\-\/]+?)\s+([\d.,]+)"
    r"(?=\s{2,}|\t|$)", re.MULTILINE
)
PADRAO_NUMERO_PURO = re.compile(r"^\s*([\d\.,]+)\s*$")

CODIGO_EMP_MAP = {
    '04': 'RSCI', '05': 'RSCIV', '06': 'RSCII', '07': 'TSCV', '08': 'RSCIII',
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

# ==== Funções de Normalização e Extração ====

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

# ==== Funções de Lógica e Classificação ====

def fixos_do_emp(emp: str):
    if emp not in EMP_MAP:
        return BASE_FIXOS
    f = dict(BASE_FIXOS)
    f["Melhoramentos"] = [float(EMP_MAP[emp]["Melhoramentos"])]
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

# ==== Funções de Extração de Dados ====

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

# ==== Funções de Processamento ====

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
    df_todas_ant, _, _ = processar_pdf_validacao(texto_anterior, modo_separacao, emp_fixo_boleto)
    df_todas_atu, _, _ = processar_pdf_validacao(texto_atual, modo_separacao, emp_fixo_boleto)
    
    df_todas_ant.rename(columns={'Valor': 'Valor Anterior'}, inplace=True)
    df_todas_atu.rename(columns={'Valor': 'Valor Atual'}, inplace=True)

    df_comp = pd.merge(
        df_todas_ant, df_todas_atu,
        on=['Empreendimento', 'Lote', 'Cliente', 'Parcela'],
        how='outer'
    )

    lotes_ant = df_todas_ant[['Empreendimento', 'Lote', 'Cliente']].drop_duplicates()
    lotes_atu = df_todas_atu[['Empreendimento', 'Lote', 'Cliente']].drop_duplicates()
    
    lotes_merged = pd.merge(lotes_ant, lotes_atu, on=['Empreendimento', 'Lote', 'Cliente'], how='outer', indicator=True)
    
    df_adicionados = lotes_merged[lotes_merged['_merge'] == 'right_only'][['Empreendimento', 'Lote', 'Cliente']]
    df_removidos = lotes_merged[lotes_merged['_merge'] == 'left_only'][['Empreendimento', 'Lote', 'Cliente']]

    df_divergencias = df_comp[
        (pd.notna(df_comp['Valor Anterior'])) & 
        (pd.notna(df_comp['Valor Atual'])) &
        (abs(df_comp['Valor Anterior'] - df_comp['Valor Atual']) > 1e-6)
    ].copy()
    df_divergencias['Diferença'] = df_divergencias['Valor Atual'] - df_divergencias['Valor Anterior']

    df_parcelas_novas = df_comp[df_comp['Valor Anterior'].isna() & pd.notna(df_comp['Valor Atual'])][['Empreendimento', 'Lote', 'Cliente', 'Parcela', 'Valor Atual']]
    df_parcelas_removidas = df_comp[df_comp['Valor Atual'].isna() & pd.notna(df_comp['Valor Anterior'])][['Empreendimento', 'Lote', 'Cliente', 'Parcela', 'Valor Anterior']]
    
    parcelas_para_remover = ['TOTAL A PAGAR', 'DESCONTO', 'DÉBITOS DO MÊS']
    
    df_divergencias = df_divergencias[~df_divergencias['Parcela'].str.strip().str.upper().isin(parcelas_para_remover)]
    df_parcelas_novas = df_parcelas_novas[~df_parcelas_novas['Parcela'].str.strip().str.upper().isin(parcelas_para_remover)]
    df_parcelas_removidas = df_parcelas_removidas[~df_parcelas_removidas['Parcela'].str.strip().str.upper().isin(parcelas_para_remover)]
    
    resumo = {
        "Lotes Mês Anterior": len(lotes_ant),
        "Lotes Mês Atual": len(lotes_atu),
        "Lotes Adicionados": len(df_adicionados),
        "Lotes Removidos": len(df_removidos),
        "Parcelas com Valor Alterado": len(df_divergencias)
    }
    df_resumo = pd.DataFrame([resumo])

    return df_resumo, df_adicionados, df_removidos, df_divergencias, df_parcelas_novas, df_parcelas_removidas

# ==== Funções de Formatação do Excel ====

def formatar_excel(output_stream, dfs: dict):
    with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
        for sheet_name, df in dfs.items():
            if not df.empty:
                df.to_excel(writer, index=False, sheet_name=sheet_name)

        number_style = NamedStyle(name='br_number_style', number_format='#,##0.00')
        
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column_cells in worksheet.columns:
                max_length = 0
                column = column_cells[0].column_letter
                for cell in column_cells:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                    if isinstance(cell.value, float):
                        cell.style = number_style
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column].width = adjusted_width
    return output_stream

# --- Início da Aplicação Web Flask ---

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'pdf_file' not in request.files or request.files['pdf_file'].filename == '':
        return "Nenhum arquivo enviado.", 400
    file = request.files['pdf_file']
    modo_separacao = request.form.get('modo_separacao', 'boleto')

    try:
        emp_fixo = None
        if modo_separacao == 'boleto':
            emp_fixo = detectar_emp_por_nome_arquivo(file.filename)
            if not emp_fixo:
                return f"<h1>Erro: Empreendimento não identificado.</h1><p>Para o modo 'Boleto', o nome do arquivo precisa terminar com um dos códigos (ex: 'Extrato_IATE.pdf').</p>", 400

        pdf_stream = file.read()
        texto_pdf = extrair_texto_pdf(pdf_stream)
        if not texto_pdf: return "Não foi possível extrair texto do PDF.", 500

        df_todas, df_cov, df_div = processar_pdf_validacao(texto_pdf, modo_separacao, emp_fixo)
        
        output = io.BytesIO()
        dfs_to_excel = {
            "Divergencias": df_div,
            "Cobertura_Analise": df_cov,
            "Todas_Parcelas_Extraidas": df_todas
        }
        formatar_excel(output, dfs_to_excel)
        output.seek(0)

        base_name = os.path.splitext(file.filename)[0]
        report_filename = f"relatorio_{base_name}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
        
        with open(report_path, 'wb') as f: f.write(output.getvalue())

        return render_template('results.html',
                               divergencias_json=df_div.to_json(orient='split', index=False) if not df_div.empty else 'null',
                               total_lotes=len(df_cov),
                               total_divergencias=len(df_div),
                               nao_classificados=len(df_cov[df_cov['Empreendimento'] == 'NAO_CLASSIFICADO']) if not df_cov.empty else 0,
                               download_url=url_for('download_file', filename=report_filename),
                               modo_usado=modo_separacao)
    
    except Exception as e:
        traceback.print_exc()
        return f"Ocorreu um erro inesperado durante o processamento: {e}", 500

@app.route('/compare', methods=['POST'])
def compare_files():
    if 'pdf_mes_anterior' not in request.files or 'pdf_mes_atual' not in request.files:
        return "Ambos os arquivos (mês anterior e atual) são necessários.", 400

    file_ant = request.files['pdf_mes_anterior']
    file_atu = request.files['pdf_mes_atual']
    modo_separacao = request.form.get('modo_separacao_comp', 'boleto')

    if file_ant.filename == '' or file_atu.filename == '':
        return "Selecione os dois arquivos para comparar.", 400

    try:
        emp_fixo_boleto = None
        if modo_separacao == 'boleto':
            emp_ant = detectar_emp_por_nome_arquivo(file_ant.filename)
            emp_atu = detectar_emp_por_nome_arquivo(file_atu.filename)
            if not emp_ant or not emp_atu:
                return "Modo Boleto: Empreendimento não identificado em um dos arquivos.", 400
            if emp_ant != emp_atu:
                return "Modo Boleto: Os arquivos devem ser do mesmo empreendimento.", 400
            emp_fixo_boleto = emp_ant

        texto_ant = extrair_texto_pdf(file_ant.read())
        texto_atu = extrair_texto_pdf(file_atu.read())

        if not texto_ant or not texto_atu:
            return "Não foi possível extrair texto de um dos PDFs.", 500

        df_resumo, df_adicionados, df_removidos, df_divergencias, df_parc_novas, df_parc_removidas = processar_comparativo(
            texto_ant, texto_atu, modo_separacao, emp_fixo_boleto
        )

        output = io.BytesIO()
        dfs_to_excel = {
            "Resumo": df_resumo,
            "Lotes Adicionados": df_adicionados,
            "Lotes Removidos": df_removidos,
            "Divergências de Valor": df_divergencias,
            "Parcelas Novas por Lote": df_parc_novas,
            "Parcelas Removidas por Lote": df_parc_removidas,
        }
        formatar_excel(output, dfs_to_excel)
        output.seek(0)
        
        report_filename = f"comparativo_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
        with open(report_path, 'wb') as f: f.write(output.getvalue())
        
        return render_template('compare_results.html',
                               resumo=df_resumo.to_dict('records')[0],
                               divergencias_json=df_divergencias.to_json(orient='split', index=False) if not df_divergencias.empty else 'null',
                               adicionados_json=df_adicionados.to_json(orient='split', index=False) if not df_adicionados.empty else 'null',
                               removidos_json=df_removidos.to_json(orient='split', index=False) if not df_removidos.empty else 'null',
                               download_url=url_for('download_file', filename=report_filename))

    except Exception as e:
        traceback.print_exc()
        return f"Ocorreu um erro inesperado durante a comparação: {e}", 500

@app.route('/download/<filename>')
def download_file(filename):
    path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    return send_file(path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=int(os.environ.get('PORT', 8080)))

