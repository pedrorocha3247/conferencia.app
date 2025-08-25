# app.py
import os, sys, re, unicodedata, io
import fitz  # PyMuPDF
import pandas as pd
from collections import OrderedDict
from flask import Flask, render_template, request, send_file, url_for

# --- Início da Lógica de Análise (Atualizada) ---

# ==== Normalização ====
DASHES = dict.fromkeys(map(ord, "\u2010\u2011\u2012\u2013\u2014\u2015\u2212"), "-")
def normalizar_texto(s: str) -> str:
    s = s.translate(DASHES).replace("\u00A0", " ")
    s = "".join(ch for ch in s if ch not in "\u200B\u200C\u200D\uFEFF")
    s = unicodedata.normalize("NFKC", s)
    return s

def extrair_texto_pdf(stream_pdf) -> str:
    # Esta função lê o PDF a partir do upload da web (stream)
    try:
        doc = fitz.open(stream=stream_pdf, filetype="pdf")
        texto = "\n".join(p.get_text("text") for p in doc)
        doc.close()
        return normalizar_texto(texto)
    except Exception as e:
        print(f"Erro ao ler o stream do PDF: {e}")
        return ""

def to_float(s: str):
    try: return float(s.replace(".","").replace(",", ".").strip())
    except: return None

# ==== Regras comuns ====
HEADERS = (
    "Remessa para Conferência","Página","Banco","IMOBILIARIOS","Débitos do Mês",
    "Vencimento","Lançamentos","Programação","Carta","DÉBITOS","ENCARGOS",
    "PAGAMENTO","TOTAL","Limite p/","TOTAL A PAGAR","PAGAMENTO EFETUADO","DESCONTO"
)

PADRAO_LOTE = re.compile(r"\b(\d{2,4}\.[A-Z]{2}\.\d{1,4})\b")
PADRAO_PARCELA_MESMA_LINHA = re.compile(
    r"^(?!(?:DÉBITOS|ENCARGOS|DESCONTO|PAGAMENTO|TOTAL|Limite p/))\s*"
    r"([A-Za-zÀ-ú][A-Za-zÀ-ú\s\.\-\/]+?)\s+([\d.,]+)"
    r"(?=\s{2,}|\t|$)",
    re.MULTILINE
)
PADRAO_NUMERO_PURO = re.compile(r"^\s*([\d\.,]+)\s*$")

# ==== NOVO: Mapa dos empreendimentos e valores dinâmicos ====
EMP_MAP = {
    "NVI":    {"Melhoramentos": 205.61, "Fundo de Transporte": 9.00},
    "NVII":   {"Melhoramentos": 245.47, "Fundo de Transporte": 9.00},
    "RSCI":   {"Melhoramentos": 250.42, "Fundo de Transporte": 9.00},
    "RSCII":  {"Melhoramentos": 240.29, "Fundo de Transporte": 9.00},
    "RSCIII": {"Melhoramentos": 281.44, "Fundo de Transporte": 9.00},
    "RSCIV":  {"Melhoramentos": 303.60, "Fundo de Transporte": 9.00},
    "IATE":   {"Melhoramentos": 240.00, "Fundo de Transporte": 9.00},  # RSCXIII (parte 1)
    "MARINA": {"Melhoramentos": 240.00, "Fundo de Transporte": 9.00},  # RSCXIII (parte 2)
    "SBRR":   {"Melhoramentos": 245.47, "Fundo de Transporte": 13.00}, # único FT=13,00
    "TSCV":   {"Melhoramentos": 0.00,   "Fundo de Transporte": 9.00},
}
BASE_FIXOS = {
    "Taxa de Conservação": [434.11],
    "Contrib. Social SLIM": [103.00, 309.00],
    "Contribuição ABRASMA - Bronze": [20.00],
    "Contribuição ABRASMA - Prata": [40.00],
    "Contribuição ABRASMA - Ouro": [60.00],
}
def fixos_do_emp(emp: str):
    f = dict(BASE_FIXOS)
    f["Melhoramentos"] = [float(EMP_MAP[emp]["Melhoramentos"])]
    f["Fundo de Transporte"] = [float(EMP_MAP[emp]["Fundo de Transporte"])]
    return f

def detectar_emp_por_nome_arquivo(path: str):
    nome = os.path.splitext(os.path.basename(path))[0].upper()
    for k in EMP_MAP.keys():
        if nome.endswith("_"+k) or nome.endswith(k):
            return k
    for k in EMP_MAP.keys():
        if f"_{k}." in (os.path.basename(path).upper()+"."):
            return k
    return None

# ==== Funções de extração (atualizadas) ====
def limpar_rotulo(lbl: str) -> str:
    lbl = re.sub(r"^TAMA\s*[-–—]\s*", "", lbl).strip()
    lbl = re.sub(r"\s+-\s+\d+/\d+$", "", lbl).strip()
    return lbl

def fatiar_blocos(texto: str):
    ms = list(PADRAO_LOTE.finditer(texto))
    blocos = []
    for i, m in enumerate(ms):
        ini = m.start()
        fim = ms[i+1].start() if i+1 < len(ms) else len(texto)
        blocos.append((m.group(1), texto[ini:fim]))
    return blocos

def tentar_nome_cliente(bloco: str) -> str:
    for linha in bloco.splitlines()[1:12]:
        L = linha.strip()
        if len(L) < 4: continue
        if any(h in L for h in HEADERS): continue
        if " " in L and sum(c.isalpha() for c in L) >= 5:
            return L
    return "Nome não localizado"

def extrair_parcelas(bloco: str):
    itens = OrderedDict()
    # 1) mesma linha
    for m in PADRAO_PARCELA_MESMA_LINHA.finditer(bloco):
        lbl = limpar_rotulo(m.group(1)); val = to_float(m.group(2))
        if lbl not in itens and val is not None:
            itens[lbl] = val
    # 2) rótulo na linha i + número puro na linha i+1
    linhas = bloco.splitlines()
    i = 0
    while i < len(linhas):
        L = linhas[i].strip()
        if L and not any(h in L for h in HEADERS):
            tem_letras = any(c.isalpha() for c in L)
            if tem_letras and not PADRAO_NUMERO_PURO.match(L):
                j = i + 1
                while j < len(linhas) and not linhas[j].strip(): j += 1
                if j < len(linhas):
                    m2 = PADRAO_NUMERO_PURO.match(linhas[j].strip())
                    if m2:
                        lbl = limpar_rotulo(L); val = to_float(m2.group(1))
                        if lbl not in itens and val is not None:
                            itens[lbl] = val
                        i = j
        i += 1
    return itens

# ==== Função de Processamento Principal (Atualizada) ====
def processar_pdf(texto_pdf: str, emp: str):
    VALORES_CORRETOS = fixos_do_emp(emp)
    texto_pdf = texto_pdf.replace("Total Geral..: 357.917,14", "")
    blocos = fatiar_blocos(texto_pdf)
    
    linhas_todas, linhas_cov, linhas_div = [], [], []

    for lote, bloco in blocos:
        cliente = tentar_nome_cliente(bloco)
        itens = extrair_parcelas(bloco)

        for rot, val in itens.items():
            linhas_todas.append({"Lote": lote, "Cliente": cliente, "Parcela": rot, "Valor": val})

        cov = {"Lote": lote, "Cliente": cliente}
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
            permitidos = VALORES_CORRETOS[rot]
            if all(abs(val - v) > 1e-6 for v in permitidos):
                linhas_div.append({
                    "Lote": lote, "Cliente": cliente, "Parcela": rot,
                    "Valor no Documento": float(val),
                    "Valor Correto": " ou ".join(f"{v:.2f}" for v in permitidos)
                })

    vistos = set(); linhas_div_dedup = []
    for r in linhas_div:
        chave = (r["Lote"], r["Parcela"])
        if chave not in vistos:
            linhas_div_dedup.append(r); vistos.add(chave)

    df_todas = pd.DataFrame(linhas_todas)
    df_cov   = pd.DataFrame(linhas_cov)
    df_div   = pd.DataFrame(linhas_div_dedup)
    return df_todas, df_cov, df_div

# --- Início da Aplicação Web Flask ---
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    """Renderiza a página inicial de upload."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Processa o arquivo PDF enviado."""
    if 'pdf_file' not in request.files:
        return "Nenhum arquivo enviado.", 400
    
    file = request.files['pdf_file']
    if file.filename == '':
        return "Nenhum arquivo selecionado.", 400

    if file and file.filename.lower().endswith('.pdf'):
        try:
            # 1. Detectar o empreendimento pelo nome do arquivo
            emp = detectar_emp_por_nome_arquivo(file.filename)
            
            # 2. Se não encontrar, retorna uma mensagem de erro amigável
            if not emp:
                return f"""
                    <h1>Erro: Empreendimento não identificado</h1>
                    <p>O nome do arquivo <strong>'{file.filename}'</strong> não corresponde a nenhum empreendimento mapeado.</p>
                    <p>O nome do arquivo precisa terminar com um dos códigos (ex: 'Extrato_IATE.pdf', 'Conferencia_RSCI.pdf').</p>
                    <a href="/">Voltar</a>
                """, 400

            pdf_stream = file.read()
            texto_pdf = extrair_texto_pdf(pdf_stream)
            if not texto_pdf:
                return "Não foi possível extrair texto do PDF.", 500

            # 3. Passar o 'emp' para a função de processamento
            df_todas, df_cov, df_div = processar_pdf(texto_pdf, emp)

            # 4. Geração do relatório em Excel (continua igual)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_div.sort_values(["Parcela", "Lote"]).to_excel(writer, index=False, sheet_name="Divergencias")
                df_cov.sort_values(["Lote"]).to_excel(writer, index=False, sheet_name="Cobertura")
                df_todas.sort_values(["Lote", "Parcela"]).to_excel(writer, index=False, sheet_name="TodasParcelas")
            
            report_filename = f"relatorio_{emp}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
            with open(report_path, 'wb') as f:
                f.write(output.getvalue())

            div_html = df_div.to_html(classes='table table-striped table-hover', index=False, border=0)

            # 5. Renderizar a página de resultados, agora com o nome do empreendimento
            return render_template('results.html', 
                                   table=div_html, 
                                   total_lotes=len(df_cov),
                                   total_divergencias=len(df_div),
                                   download_url=url_for('download_file', filename=report_filename),
                                   emp_detectado=emp)
        
        except Exception as e:
            # Imprime o erro no console do Render para depuração
            print(f"Ocorreu um erro: {e}", file=sys.stderr)
            return f"Ocorreu um erro inesperado durante o processamento. Verifique os logs do servidor.", 500

    return "Formato de arquivo inválido. Por favor, envie um PDF.", 400

@app.route('/download/<filename>')
def download_file(filename):
    """Serve o arquivo de relatório para download."""
    path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    return send_file(path, as_attachment=True)

if __name__ == '__main__':
    # Usar host='0.0.0.0' para tornar acessível na sua rede local
    # A porta é definida pelo Render, não precisa especificar aqui para produção
    app.run(debug=True, host='0.0.0.0')
