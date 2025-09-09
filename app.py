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
from openpyxl.styles import NamedStyle, Alignment

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
    "TSCV": {"Melhoramentos": 0.00, "Fundo de Transporte": 9.00},
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
    if not linhas:
        return "Nome não localizado"

    primeira_linha = linhas[0].strip()
    match_lote = PADRAO_LOTE.match(primeira_linha)
    if match_lote:
        nome_candidato = primeira_linha[match_lote.end():].strip()
        if len(nome_candidato) > 4 and ' ' in nome_candidato and not any(h.upper() in nome_candidato.upper() for h in HEADERS):
            return nome_candidato

    for linha in linhas[1:5]:
        linha_limpa = linha.strip()
        is_valid_name = (
            len(linha_limpa) > 5 and ' ' in linha_limpa and
            sum(c.isalpha() for c in linha_limpa) / len(linha_limpa.replace(" ", "")) > 0.8 and
            not any(h.upper() in linha_limpa.upper() for h in HEADERS) and
            not PADRAO_LOTE.match(linha_limpa) and
            not re.search(r'\d{2}/\d{2}/\d{4}', linha_limpa) and
            not linha_limpa.upper().startswith("TOTAL")
        )
        if is_valid_name:
            return linha_limpa
    return "Nome não localizado"

def extrair_parcelas(bloco: str):
    """
    Função corrigida para isolar a área de lançamentos e extrair as parcelas de forma mais precisa.
    """
    itens = OrderedDict()
    
    # Isola o bloco de texto que contém as parcelas, começando após "Lançamentos".
    pos_lancamentos = bloco.find("Lançamentos")
    bloco_de_trabalho = bloco[pos_lancamentos + len("Lançamentos"):] if pos_lancamentos != -1 else bloco

    # Pré-processamento para remover colunas de resumo financeiro que aparecem na mesma linha.
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
        val = to_float(m.group(2))
        if lbl and lbl not in itens and val is not None:
            itens[lbl] = val

    # Lógica 2: Captura parcelas cujo valor está na linha seguinte.
    linhas = bloco_limpo.splitlines()
    for i, linha in enumerate(linhas):
        linha_limpa = linha.strip()
        if not linha_limpa:
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

# ==== Função de Processamento Principal ====

def processar_pdf(texto_pdf: str, modo_separacao: str, emp_fixo: str = None):
    blocos = fatiar_blocos(texto_pdf)
    if not blocos:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    linhas_todas, linhas_cov, linhas_div = [], [], []
    for lote, bloco in blocos:
        emp_atual = emp_fixo if modo_separacao == 'boleto' else detectar_emp_por_lote(lote)
        cliente = tentar_nome_cliente(bloco)
        itens = extrair_parcelas(bloco)
        
        VALORES_CORRETOS = fixos_do_emp(emp_atual)

        for rot, val in itens.items():
            linhas_todas.append({"Empreendimento": emp_atual, "Lote": lote, "Cliente": cliente, "Parcela": rot, "Valor": val})

        cov = {"Empreendimento": emp_atual, "Lote": lote, "Cliente": cliente}
        for k in VALORES_CORRETOS.keys(): cov[k] = None
        
        for rot, val in itens.items():
            if rot in VALORES_CORRETOS:
                cov[rot] = val
        
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

    vistos = set()
    linhas_div_dedup = []
    for r in linhas_div:
        chave = (r["Lote"], r["Parcela"])
        if chave not in vistos:
            linhas_div_dedup.append(r)
            vistos.add(chave)

    df_todas = pd.DataFrame(linhas_todas)
    df_cov = pd.DataFrame(linhas_cov)
    df_div = pd.DataFrame(linhas_div_dedup)
    
    return df_todas, df_cov, df_div

# ==== Funções de Formatação do Excel ====

def formatar_excel(output_stream, dfs: dict):
    with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
        for sheet_name, df in dfs.items():
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
                    if isinstance(cell.value, (int, float)):
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
    if 'pdf_file' not in request.files: return "Nenhum arquivo enviado.", 400
    file = request.files['pdf_file']
    modo_separacao = request.form.get('modo_separacao', 'boleto')

    if file.filename == '': return "Nenhum arquivo selecionado.", 400

    if file and file.filename.lower().endswith('.pdf'):
        try:
            emp_fixo = None
            if modo_separacao == 'boleto':
                emp_fixo = detectar_emp_por_nome_arquivo(file.filename)
                if not emp_fixo:
                    return f"""<h1>Erro: Empreendimento não identificado (Modo Boleto)</h1>
                           <p>O nome do arquivo <strong>'{file.filename}'</strong> não corresponde a um empreendimento mapeado.</p>
                           <p>Para o modo 'Boleto', o nome do arquivo precisa terminar com um dos códigos (ex: 'Extrato_IATE.pdf').</p>
                           <a href="/">Voltar</a>""", 400

            pdf_stream = file.read()
            texto_pdf = extrair_texto_pdf(pdf_stream)
            if not texto_pdf: return "Não foi possível extrair texto do PDF.", 500

            df_todas, df_cov, df_div = processar_pdf(texto_pdf, modo_separacao, emp_fixo)
            
            if not df_div.empty: df_div = df_div.sort_values(by=['Empreendimento', 'Lote'])
            if not df_cov.empty: df_cov = df_cov.sort_values(by=['Empreendimento', 'Lote'])
            if not df_todas.empty: df_todas = df_todas.sort_values(by=['Empreendimento', 'Lote'])

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

            div_html = df_div.to_html(classes='table table-striped table-hover', index=False, border=0) if not df_div.empty else "<p>Nenhuma divergência encontrada.</p>"
            
            total_lotes = len(df_cov)
            total_divergencias = len(df_div)
            nao_classificados = len(df_cov[df_cov['Empreendimento'] == 'NAO_CLASSIFICADO']) if not df_cov.empty else 0

            return render_template('results.html',
                                   table=div_html,
                                   total_lotes=total_lotes,
                                   total_divergencias=total_divergencias,
                                   nao_classificados=nao_classificados,
                                   download_url=url_for('download_file', filename=report_filename),
                                   modo_usado=modo_separacao)
        
        except Exception as e:
            print(f"Ocorreu um erro no processamento: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc()
            return f"Ocorreu um erro inesperado durante o processamento. Verifique os logs do servidor.", 500
    
    return "Formato de arquivo inválido. Por favor, envie um PDF.", 400

@app.route('/download/<filename>')
def download_file(filename):
    path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    return send_file(path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)

