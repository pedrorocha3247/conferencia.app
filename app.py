# app.py
import sys, re, unicodedata, os, io
import fitz  # PyMuPDF
import pandas as pd
from collections import OrderedDict
from flask import Flask, render_template, request, send_file, url_for

# --- Início da sua lógica original (copiada do seu script) ---
DASHES = dict.fromkeys(map(ord, "\u2010\u2011\u2012\u2013\u2014\u2015\u2212"), "-")

def normalizar_texto(s: str) -> str:
    s = s.translate(DASHES).replace("\u00A0", " ")
    s = "".join(ch for ch in s if ch not in "\u200B\u200C\u200D\uFEFF")
    s = unicodedata.normalize("NFKC", s)
    return s

def extrair_texto_pdf(stream_pdf) -> str:
    try:
        # Modificado para ler de um stream de bytes em vez de um caminho de arquivo
        doc = fitz.open(stream=stream_pdf, filetype="pdf")
        texto = "\n".join(page.get_text("text") for page in doc)
        doc.close()
        return normalizar_texto(texto)
    except Exception as e:
        print(f"Erro ao ler o stream do PDF: {e}")
        return ""

def to_float(s: str):
    try:
        return float(s.replace(".", "").replace(",", ".").strip())
    except:
        return None

VALORES_CORRETOS = {
    "Taxa de Conservação": [434.11],
    "Melhoramentos": [240.00],
    "Contrib. Social SLIM": [103.00, 309.00],
    "Fundo de Transporte": [9.00],
    "Contribuição ABRASMA - Bronze": [20.00],
    "Contribuição ABRASMA - Prata": [40.00],
    "Contribuição ABRASMA - Ouro": [60.00],
}

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

def limpar_rotulo(lbl: str) -> str:
    lbl = re.sub(r"^TAMA\s*[-–—]\s*", "", lbl).strip()
    lbl = re.sub(r"\s+-\s+\d+/\d+$", "", lbl).strip()
    return lbl

def fatiar_blocos_por_lote(texto: str):
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

def extrair_parcelas(block: str):
    itens = OrderedDict()
    for m in PADRAO_PARCELA_MESMA_LINHA.finditer(block):
        lbl = limpar_rotulo(m.group(1))
        val = to_float(m.group(2))
        if lbl not in itens and val is not None:
            itens[lbl] = val
    linhas = block.splitlines()
    i = 0
    while i < len(linhas):
        L = linhas[i].strip()
        if L and not any(h in L for h in HEADERS):
            tem_letras = any(c.isalpha() for c in L)
            if tem_letras and not PADRAO_NUMERO_PURO.match(L):
                j = i + 1
                while j < len(linhas) and not linhas[j].strip():
                    j += 1
                if j < len(linhas):
                    m2 = PADRAO_NUMERO_PURO.match(linhas[j].strip())
                    if m2:
                        lbl = limpar_rotulo(L)
                        val = to_float(m2.group(1))
                        if lbl not in itens and val is not None:
                            itens[lbl] = val
                        i = j
        i += 1
    return list(itens.items())

def processar(texto_pdf: str):
    texto_pdf = texto_pdf.replace("Total Geral..: 357.917,14", "")
    blocos = fatiar_blocos_por_lote(texto_pdf)
    linhas_todas, linhas_cov, linhas_div = [], [], []
    for lote, bloco in blocos:
        cliente = tentar_nome_cliente(bloco)
        pares = extrair_parcelas(bloco)
        for rot, val in pares:
            linhas_todas.append({"Lote": lote, "Cliente": cliente, "Parcela": rot, "Valor": val})
        cov_row = {"Lote": lote, "Cliente": cliente}
        for alvo in VALORES_CORRETOS.keys():
            cov_row[alvo] = None
        for rot, val in pares:
            if rot in VALORES_CORRETOS:
                cov_row[rot] = val
        vistos = [k for k in VALORES_CORRETOS.keys() if cov_row[k] is not None]
        cov_row["QtdParc_Alvo"] = len(vistos)
        cov_row["Parc_Alvo"] = ", ".join(vistos)
        linhas_cov.append(cov_row)
        for rot in vistos:
            val = cov_row[rot]
            if val is None: continue
            permitidos = VALORES_CORRETOS[rot]
            if all(abs(val - v) > 1e-6 for v in permitidos):
                linhas_div.append({
                    "Lote": lote, "Cliente": cliente, "Parcela": rot,
                    "Valor no Documento": float(val),
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
    return df_todas, df_cov, df_div, len(blocos)
# --- Fim da sua lógica original ---


# --- Início da Aplicação Web Flask ---
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads' # Opcional: para salvar arquivos temporariamente
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
            # Lê o conteúdo do arquivo em memória
            pdf_stream = file.read()
            
            # Processa o PDF
            texto_pdf = extrair_texto_pdf(pdf_stream)
            if not texto_pdf:
                return "Não foi possível extrair texto do PDF.", 500

            df_todas, df_cov, df_div, total_lotes = processar(texto_pdf)

            # Prepara o arquivo Excel em memória para download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_div.sort_values(["Parcela", "Lote"]).to_excel(writer, index=False, sheet_name="Divergencias")
                df_cov.sort_values(["Lote"]).to_excel(writer, index=False, sheet_name="Cobertura")
                df_todas.sort_values(["Lote", "Parcela"]).to_excel(writer, index=False, sheet_name="TodasParcelas")
            
            # Salva temporariamente o buffer para criar um link de download
            report_filename = f"relatorio_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
            with open(report_path, 'wb') as f:
                f.write(output.getvalue())

            # Converte a tabela de divergências para HTML para exibição na página
            div_html = df_div.to_html(classes='table table-striped table-hover', index=False, border=0)

            return render_template('results.html', 
                                   table=div_html, 
                                   total_lotes=total_lotes,
                                   total_divergencias=len(df_div),
                                   download_url=url_for('download_file', filename=report_filename))
        
        except Exception as e:
            return f"Ocorreu um erro durante o processamento: {e}", 500

    return "Formato de arquivo inválido. Por favor, envie um PDF.", 400

@app.route('/download/<filename>')
def download_file(filename):
    """Serve o arquivo de relatório para download."""
    path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    return send_file(path, as_attachment=True)

if __name__ == '__main__':
    # Usar host='0.0.0.0' para tornar acessível na sua rede local
    app.run(debug=True, host='0.0.0.0', port=5000)