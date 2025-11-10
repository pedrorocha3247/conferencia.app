# -*- coding: utf-8 -*-
import os
import re
import unicodedata
import io
import fitz  # PyMuPDF
import pandas as pd
from collections import OrderedDict, Counter # Importa o Counter
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
    r"^(?!(?:DÉBITOS|ENCARGOS|DESCONTO|PAGAMENTO|TOTAL(?!\s*A PAGAR)|Limite p/))\s*" # <-- MUDANÇA AQUI
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
    s = s.translate(DASHES).replace("\u00A0", " ") # Substitui hífens e nbsp
    s = "".join(ch for ch in s if ch not in "\u200B\u200C\u200D\uFEFF") # Remove ZWSP e similares
    s = unicodedata.normalize("NFKC", s) # Normaliza caracteres unicode
    return s

def extrair_texto_pdf(stream_pdf) -> str:
    texto_completo = ""
    try:
        # Abre o PDF a partir do stream de bytes
        with fitz.open(stream=stream_pdf, filetype="pdf") as doc:
             # Itera sobre cada página e extrai o texto, mantendo a ordem
             for page_num in range(len(doc)):
                 page = doc.load_page(page_num)
                 # get_text("text", sort=True) tenta ordenar o texto como lido visualmente
                 texto_pagina = page.get_text("text", sort=True)
                 texto_completo += texto_pagina + "\n" # Adiciona nova linha entre páginas
        return normalizar_texto(texto_completo)
    except Exception as e:
        print(f"Erro detalhado ao ler o stream do PDF: {type(e).__name__} - {e}")
        traceback.print_exc() # Imprime o stack trace completo no log
        return "" # Retorna string vazia em caso de erro

# =======================================================
# === FUNÇÃO DE CONVERSÃO DE VALOR UNIFICADA E CORRIGIDA ===
# =======================================================
def normalizar_valor(valor):
    """Converte string para float, lidando com formatos , e . como decimal."""
    if valor is None:
        return 0.0 # Retorna 0.0 para None para simplificar somas
    if isinstance(valor, (int, float)):
        return round(float(valor), 2)
    
    s_norm = str(valor).strip().replace("R$", "").replace(" ", "").replace("\xa0", "")
    
    has_comma = "," in s_norm
    has_dot = "." in s_norm

    if has_comma and has_dot:
        # Formato 1.234,56 (vírgula é decimal)
        if s_norm.rfind(',') > s_norm.rfind('.'):
             s_norm = s_norm.replace(".", "").replace(",", ".")
        # Formato 1,234.56 (ponto é decimal)
        else:
             s_norm = s_norm.replace(",", "")
    elif has_comma:
        # Formato 1234,56 (vírgula é decimal)
        s_norm = s_norm.replace(",", ".")
    elif has_dot:
        # Formato 1234.56 (ponto é decimal)
        # OU formato 8.054.23 (pontos são milhares E decimal) - PDF INCONSISTENTE
        if s_norm.count('.') > 1:
            # Remove todos os pontos, exceto o último
            parts = s_norm.split('.')
            s_norm = "".join(parts[:-1]) + "." + parts[-1]
    
    try:
        return round(float(s_norm), 2)
    except (ValueError, TypeError):
        print(f"[AVISO] Falha ao normalizar valor: '{valor}' -> '{s_norm}'")
        return 0.0 # Retorna 0.0 em caso de falha de conversão
# =======================================================
# === FIM DA FUNÇÃO UNIFICADA ===
# =======================================================

def fixos_do_emp(emp: str, modo_separacao: str):
    """Retorna o dicionário de parcelas fixas esperadas com base no empreendimento e modo."""
    if modo_separacao == 'boleto':
        if emp not in EMP_MAP:
            return BASE_FIXOS # Retorna base se o empreendimento não tiver mapa específico
        f = dict(BASE_FIXOS) # Cria cópia da base
        # Adiciona/sobrescreve valores específicos do empreendimento
        if EMP_MAP.get(emp):
            if "Melhoramentos" in EMP_MAP[emp]:
                f["Melhoramentos"] = [float(EMP_MAP[emp]["Melhoramentos"])]
            if "Fundo de Transporte" in EMP_MAP[emp]:
                f["Fundo de Transporte"] = [float(EMP_MAP[emp]["Fundo de Transporte"])]
        return f
    elif modo_separacao == 'debito_credito':
        # Assume que Débito/Crédito usa as mesmas parcelas base que Boleto
        # Se for diferente, crie um BASE_FIXOS_DEBITO_CREDITO
        return BASE_FIXOS
    elif modo_separacao == 'ccb_realiza':
        # Retorna o dicionário específico para CCB/Realiza (sem valores fixos pré-definidos)
        return BASE_FIXOS_CCB
    else:
        print(f"[AVISO] Modo de separação desconhecido '{modo_separacao}' em fixos_do_emp.")
        return {} # Retorna dicionário vazio para evitar erros

def detectar_emp_por_nome_arquivo(path: str):
    """Tenta detectar o código do empreendimento pelo sufixo no nome do arquivo."""
    if not path: return None
    nome = os.path.splitext(os.path.basename(path))[0].upper()
    # Verifica se termina com _CODIGO ou apenas CODIGO
    for k in EMP_MAP.keys():
        if nome.endswith("_" + k) or nome.endswith(k):
            return k
    # Caso especial para SBRR (se contiver no nome, mas não como sufixo exato)
    if "SBRR" in nome:
        return "SBRR" # Pode precisar de ajuste se houver outros com SBRR
    return None

def detectar_emp_por_lote(lote: str):
    """Detecta o empreendimento com base no prefixo do código do lote."""
    if not lote or "." not in lote:
        return "NAO_CLASSIFICADO"
    prefixo = lote.split('.')[0]
    # Retorna o código do mapa ou "NAO_CLASSIFICADO" se não encontrar
    return CODIGO_EMP_MAP.get(prefixo, "NAO_CLASSIFICADO")

def limpar_rotulo(lbl: str) -> str:
    """Remove prefixos e sufixos comuns dos rótulos das parcelas."""
    if not isinstance(lbl, str): return "" # Garante que é string
    lbl = re.sub(r"^TAMA\s*[-–—]\s*", "", lbl, flags=re.IGNORECASE).strip() # Remove prefixo TAMA
    lbl = re.sub(r"\s+-\s+\d+/\d+$", "", lbl).strip() # Remove sufixo de parcela N/M
    lbl = re.sub(r'\s{2,}', ' ', lbl).strip() # Remove espaços múltiplos
    return lbl

def fatiar_blocos(texto: str):
    """Divide o texto do PDF em blocos, cada um começando com um código de lote."""
    # Adiciona uma quebra de linha antes de cada padrão de lote para facilitar a divisão
    texto_processado = PADRAO_LOTE.sub(r"\n\1", texto)
    # Encontra todas as ocorrências do padrão de lote
    matches = list(PADRAO_LOTE.finditer(texto_processado))
    blocos = []
    # Itera sobre as correspondências para extrair o texto entre elas
    for i, match in enumerate(matches):
        lote_atual = match.group(1)
        inicio_bloco = match.start()
        # Fim do bloco é o início do próximo lote, ou o final do texto se for o último
        fim_bloco = matches[i+1].start() if i+1 < len(matches) else len(texto_processado)
        # Extrai o texto do bloco
        texto_bloco = texto_processado[inicio_bloco:fim_bloco].strip()
        if texto_bloco: # Adiciona apenas se o bloco não estiver vazio
             blocos.append((lote_atual, texto_bloco))
    if not blocos:
         print("[AVISO] Nenhum bloco de lote encontrado no PDF.")
    return blocos

def tentar_nome_cliente(bloco: str) -> str:
    """Tenta extrair o nome do cliente das primeiras linhas do bloco."""
    linhas = bloco.split('\n')
    if not linhas: return "Nome não localizado"

    # Considera as primeiras 5-6 linhas como candidatas
    linhas_para_buscar = linhas[:6]
    nome_candidato = "Nome não localizado"

    for linha in linhas_para_buscar:
        # Remove o código do lote da linha e espaços extras
        linha_sem_lote = PADRAO_LOTE.sub('', linha).strip()
        if not linha_sem_lote: continue # Pula linhas vazias após remover lote

        # Heurísticas mais refinadas para identificar um nome:
        is_valid_name = (
            len(linha_sem_lote) > 5 and # Pelo menos 6 caracteres
            ' ' in linha_sem_lote and # Deve conter espaço (nome composto)
            sum(c.isalpha() for c in linha_sem_lote.replace(" ", "")) / len(linha_sem_lote.replace(" ", "")) > 0.7 and # Maioria letras
            not any(h.upper() in linha_sem_lote.upper() for h in HEADERS if h) and # Não contém cabeçalhos
            not re.search(r'\d{2}/\d{2}/\d{4}', linha_sem_lote) and # Não é data
            not re.match(r'^[\d.,\s]+$', linha_sem_lote) and # Não é apenas número
            not linha_sem_lote.upper().startswith(("TOTAL", "BANCO", "03-", "LIMITE P/", "PÁGINA")) # Não começa com termos comuns
        )

        if is_valid_name:
            # Assume que a primeira linha válida encontrada é o nome
            nome_candidato = linha_sem_lote
            break # Para após encontrar o primeiro candidato válido

    return nome_candidato.strip()

def extrair_parcelas(bloco: str):
    """Extrai os nomes e valores das parcelas dentro de um bloco de texto."""
    itens = OrderedDict()
    # Tenta focar na seção "Lançamentos", se existir
    pos_lancamentos = bloco.find("Lançamentos")
    bloco_de_trabalho = bloco[pos_lancamentos:] if pos_lancamentos != -1 else bloco

    # Limpeza adicional: remove linhas de totais que podem confundir
    bloco_limpo_linhas = []
    linhas_originais = bloco_de_trabalho.splitlines()
    ignorar_proxima_linha_se_numero = False # Flag para o padrão Label \n Valor

    for i, linha in enumerate(linhas_originais):
        # Remove linhas de resumo que aparecem muito à direita
        match_total_direita = re.search(r'\s{4,}(DÉBITOS DO MÊS ANTERIOR|ENCARGOS POR ATRASO|PAGAMENTO EFETUADO)\s+[\d.,]+$', linha)
        linha_processada = linha[:match_total_direita.start()] if match_total_direita else linha
        linha_processada = linha_processada.strip()

        # Ignora linhas que são cabeçalhos conhecidos ou vazias
        if not linha_processada or any(h.strip().upper() == linha_processada.upper() for h in ["Lançamentos", "Débitos do Mês"]):
            continue

        # Se a flag estiver ativa, ignora esta linha (já foi usada como valor)
        if ignorar_proxima_linha_se_numero:
             ignorar_proxima_linha_se_numero = False
             continue

        # Tenta aplicar o padrão [Label] [Valor] na mesma linha
        match_mesma_linha = PADRAO_PARCELA_MESMA_LINHA.match(linha_processada)
        if match_mesma_linha:
            lbl = limpar_rotulo(match_mesma_linha.group(1))
            val = normalizar_valor(match_mesma_linha.group(2)) # <-- USA A FUNÇÃO CORRIGIDA
            if lbl and lbl not in itens and val is not None:
                itens[lbl] = val
                continue # Pula para a próxima linha

        # Se não casou acima, verifica se é um Label cuja próxima linha é um Valor
        is_potential_label = (
            any(c.isalpha() for c in linha_processada) and # Contém letras
            limpar_rotulo(linha_processada) not in itens # Label ainda não capturado
        )

        if is_potential_label:
            # Verifica a próxima linha NÃO VAZIA
            j = i + 1
            while j < len(linhas_originais) and not linhas_originais[j].strip():
                j += 1
            if j < len(linhas_originais):
                 linha_seguinte_limpa = linhas_originais[j].strip()
                 match_num_puro = PADRAO_NUMERO_PURO.match(linha_seguinte_limpa)
                 # Se a linha seguinte for puramente numérica
                 if match_num_puro:
                      lbl = limpar_rotulo(linha_processada)
                      val = normalizar_valor(match_num_puro.group(1)) # <-- USA A FUNÇÃO CORRIGIDA
                      if lbl and lbl not in itens and val is not None:
                           itens[lbl] = val
                           ignorar_proxima_linha_se_numero = True # Marca a linha j para ser ignorada na próxima iteração
                           continue # Pula para a próxima linha i
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
        VALORES_CORRETOS = fixos_do_emp(emp_atual, modo_separacao) # Passa o modo

        for rot, val in itens.items():
            # 'val' já é um float corrigido pela função normalizar_valor
            linhas_todas.append({"Empreendimento": emp_atual, "Lote": lote, "Cliente": cliente, "Parcela": rot, "Valor": val})

        cov = {"Empreendimento": emp_atual, "Lote": lote, "Cliente": cliente}
        for k in VALORES_CORRETOS.keys(): cov[k] = None # Inicializa colunas
        for rot, val in itens.items():
            if rot in VALORES_CORRETOS: cov[rot] = val # Preenche valores encontrados

        vistos = [k for k in VALORES_CORRETOS if cov[k] is not None]
        cov["QtdParc_Alvo"] = len(vistos)
        cov["Parc_Alvo"] = ", ".join(vistos)
        linhas_cov.append(cov)

        # Validação de valor (apenas se houver valores permitidos definidos)
        if modo_separacao != 'ccb_realiza': # Não valida valores para CCB (lista vazia)
            for rot in vistos:
                val = cov[rot]
                if val is None: continue
                permitidos = VALORES_CORRETOS.get(rot, [])
                if permitidos and all(abs(val - v) > 1e-6 for v in permitidos):
                    linhas_div.append({
                        "Empreendimento": emp_atual, "Lote": lote, "Cliente": cliente,
                        "Parcela": rot, "Valor no Documento": float(val), # val já é float
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

    # Extrai totais (agora com valores corretos de normalizar_valor)
    df_totais_ant = df_todas_ant_raw[df_todas_ant_raw['Parcela'].astype(str).str.strip().str.upper() == 'TOTAL A PAGAR'].copy()
    df_totais_ant = df_totais_ant[['Empreendimento', 'Lote', 'Cliente', 'Valor']].rename(columns={'Valor': 'Total Anterior'})

    df_totais_atu = df_todas_atu_raw[df_todas_atu_raw['Parcela'].astype(str).str.strip().str.upper() == 'TOTAL A PAGAR'].copy()
    df_totais_atu = df_totais_atu[['Empreendimento', 'Lote', 'Cliente', 'Valor']].rename(columns={'Valor': 'Total Atual'})

    # Remove parcelas indesejadas para comparação item a item
    parcelas_para_remover = ['TOTAL A PAGAR', 'DESCONTO', 'DÉBITOS DO MÊS', 'DÉBITOS DO MÊS ANTERIOR', 'ENCARGOS POR ATRASO', 'PAGAMENTO EFETUADO']
    df_todas_ant = df_todas_ant_raw[~df_todas_ant_raw['Parcela'].astype(str).str.strip().str.upper().isin(parcelas_para_remover)].copy()
    df_todas_atu = df_todas_atu_raw[~df_todas_atu_raw['Parcela'].astype(str).str.strip().str.upper().isin(parcelas_para_remover)].copy()

    df_todas_ant = df_todas_ant[~df_todas_ant['Parcela'].astype(str).str.strip().str.upper().str.startswith('TOTAL BANCO')].copy()
    df_todas_atu = df_todas_atu[~df_todas_atu['Parcela'].astype(str).str.strip().str.upper().str.startswith('TOTAL BANCO')].copy()

    df_todas_ant.rename(columns={'Valor': 'Valor Anterior'}, inplace=True)
    df_todas_atu.rename(columns={'Valor': 'Valor Atual'}, inplace=True)

    # Merge para comparação
    df_comp = pd.merge(df_todas_ant, df_todas_atu, on=['Empreendimento', 'Lote', 'Cliente', 'Parcela'], how='outer')

    # Identifica lotes adicionados/removidos
    lotes_ant = df_todas_ant_raw[['Empreendimento', 'Lote', 'Cliente']].drop_duplicates()
    lotes_atu = df_todas_atu_raw[['Empreendimento', 'Lote', 'Cliente']].drop_duplicates()
    lotes_merged = pd.merge(lotes_ant, lotes_atu, on=['Empreendimento', 'Lote', 'Cliente'], how='outer', indicator=True)

    df_adicionados_base = lotes_merged[lotes_merged['_merge'] == 'right_only'][['Empreendimento', 'Lote', 'Cliente']]
    df_removidos_base = lotes_merged[lotes_merged['_merge'] == 'left_only'][['Empreendimento', 'Lote', 'Cliente']]

    # Adiciona valor total aos lotes adicionados/removidos
    df_adicionados = pd.merge(df_adicionados_base, df_totais_atu, on=['Empreendimento', 'Lote', 'Cliente'], how='left')
    df_removidos = pd.merge(df_removidos_base, df_totais_ant, on=['Empreendimento', 'Lote', 'Cliente'], how='left')

    # Identifica divergências de valor, parcelas novas e removidas
    df_divergencias = df_comp[
        (pd.notna(df_comp['Valor Anterior'])) &
        (pd.notna(df_comp['Valor Atual'])) &
        (abs(df_comp['Valor Anterior'] - df_comp['Valor Atual']) > 0.025) # Tolerância de ~2 centavos
    ].copy()
    if not df_divergencias.empty:
         df_divergencias['Diferença'] = df_divergencias['Valor Atual'] - df_divergencias['Valor Anterior']

    df_parcelas_novas = df_comp[df_comp['Valor Anterior'].isna() & pd.notna(df_comp['Valor Atual'])][['Empreendimento', 'Lote', 'Cliente', 'Parcela', 'Valor Atual']].copy()
    df_parcelas_removidas = df_comp[df_comp['Valor Atual'].isna() & pd.notna(df_comp['Valor Anterior'])][['Empreendimento', 'Lote', 'Cliente', 'Parcela', 'Valor Anterior']].copy()

    # Calcula totais para o resumo
    total_adicionados_valor = df_adicionados['Total Atual'].sum() if 'Total Atual' in df_adicionados.columns else 0
    total_removidos_valor = df_removidos['Total Anterior'].sum() if 'Total Anterior' in df_removidos.columns else 0
    total_divergencias_valor = df_divergencias['Diferença'].sum() if 'Diferença' in df_divergencias.columns else 0
    total_mes_anterior_valor = df_totais_ant['Total Anterior'].sum() if 'Total Anterior' in df_totais_ant.columns else 0
    total_mes_atual_valor = df_totais_atu['Total Atual'].sum() if 'Total Atual' in df_totais_atu.columns else 0

    # Cria DataFrame de resumo
    resumo_financeiro_data = {
        ' ': ['Lotes Mês Anterior', 'Lotes Mês Atual', 'Lotes Adicionados', 'Lotes Removidos', 'Parcelas com Valor Alterado'],
        'LOTES': [len(lotes_ant), len(lotes_atu), len(df_adicionados), len(df_removidos), df_divergencias['Lote'].nunique() if not df_divergencias.empty else 0],
        'TOTAIS': [total_mes_anterior_valor, total_mes_atual_valor, total_adicionados_valor, total_removidos_valor, total_divergencias_valor]
    }
    df_resumo_completo = pd.DataFrame(resumo_financeiro_data)

    # Retorna todos os DataFrames gerados
    return df_resumo_completo, df_adicionados, df_removidos, df_divergencias, df_parcelas_novas, df_parcelas_removidas


def formatar_excel(output_stream, dfs: dict):
    """Formata planilhas de Validação e Comparação (não Repasse)."""
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
                                 if isinstance(cell.value, (int, float)) and cell.number_format != 'General' and not is_first_row:
                                     if '##0.00' in cell.number_format:
                                         formatted_value = f"{cell.value:,.2f}"
                                     elif '0' == cell.number_format:
                                         formatted_value = f"{cell.value:d}"
                                     else:
                                         formatted_value = str(cell.value)
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
                    worksheet.column_dimensions[column].width = min(max(adjusted_width, 10), 60)

                worksheet.auto_filter.ref = worksheet.dimensions
                print(f"[LOG] Autofilter aplicado à planilha '{sheet_name}'. Ref: {worksheet.dimensions}")
    return output_stream


def copiar_formatacao(origem, destino):
    """Copia toda a formatação de uma célula para outra."""
    if origem and hasattr(origem, 'has_style') and origem.has_style:
        destino.font = copy(origem.font)
        destino.border = copy(origem.border)
        destino.fill = copy(origem.fill)
        destino.number_format = copy(origem.number_format)
        destino.protection = copy(origem.protection)
        destino.alignment = copy(origem.alignment)

# =======================================================
# === NOVA FUNÇÃO achar_coluna_flex ===
# =======================================================
def achar_coluna_flex(sheet, nomes_possiveis: list):
    """Encontra o número da coluna (1-indexado) para o primeiro nome que corresponder."""
    if sheet.max_row == 0: return None
    nomes_lower = [nome.lower() for nome in nomes_possiveis]
    for cell in sheet[1]:
        if cell.value and str(cell.value).strip().lower() in nomes_lower:
            return cell.column
    return None

# =======================================================
# === FUNÇÃO criar_planilha_saida ATUALIZADA (Repasse) ===
# =======================================================
def criar_planilha_saida(linhas, ws_diario, incluir_status=False):
    """Cria a planilha de saída para o Repasse com formatação customizada."""
    wb_out = Workbook()
    ws_out = wb_out.active

    # Req 1: Sem grades de fundo
    ws_out.sheet_view.showGridLines = False

    # Req 3: Definir estilo do cabeçalho (Verde Forte)
    header_fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid") # Verde forte
    header_font = Font(bold=True, color="FFFFFF") # Fonte Branca
    header_border = Border(bottom=Side(style='thin', color='A0A0A0')) # Borda inferior leve

    # Copia cabeçalho e aplica NOVO estilo
    if ws_diario.max_row > 0:
        num_cols_header = ws_diario.max_column
        for i, cell in enumerate(ws_diario[1], 1):
            if cell:
                novo = ws_out.cell(row=1, column=i, value=cell.value)
                novo.fill = header_fill
                novo.font = header_font
                novo.border = header_border
                
                col_letter = get_column_letter(i)
                if col_letter in ws_diario.column_dimensions:
                     ws_out.column_dimensions[col_letter].width = ws_diario.column_dimensions[col_letter].width
                else: ws_out.column_dimensions[col_letter].width = 15
            else:
                 ws_out.cell(row=1, column=i, value=None)
    else:
        num_cols_header = 0
        print("[AVISO] ws_diario (planilha modelo) estava vazio, nenhum cabeçalho copiado.")

    col_status = 0
    if incluir_status:
        col_status = num_cols_header + 1
        cell_status_header = ws_out.cell(row=1, column=col_status, value="Status")
        cell_status_header.fill = header_fill
        cell_status_header.font = header_font
        cell_status_header.border = header_border
        ws_out.column_dimensions[get_column_letter(col_status)].width = 45

    # Estilos para dados (Req 2: Sem preenchimento, Sem bordas)
    no_fill = PatternFill(fill_type=None)
    no_border = Border()

    # Copia dados SEM formatação (exceto number_format)
    linha_out = 2
    for linha_info in linhas:
        linha, status = linha_info
        
        if linha is not None:
             for i, cell_data in enumerate(linha, 1):
                 try:
                     valor = cell_data.value if hasattr(cell_data, "value") else cell_data
                     novo = ws_out.cell(row=linha_out, column=i, value=valor)
                     
                     if hasattr(cell_data, "number_format") and cell_data.number_format:
                         novo.number_format = cell_data.number_format
                     
                     novo.fill = no_fill
                     novo.border = no_border
                         
                 except Exception as e:
                      print(f"[Aviso] Erro ao processar célula {i} da linha {linha_out}: {e}. Valor: {cell_data}")
                      ws_out.cell(row=linha_out, column=i, value=f"ERRO: {e}")
        
        if incluir_status and col_status > 0:
             cell_status_data = ws_out.cell(row=linha_out, column=col_status, value=status)
             cell_status_data.fill = no_fill
             cell_status_data.border = no_border
        
        linha_out += 1

    if incluir_status and len(linhas) > 0:
         total_cell = ws_out.cell(row=linha_out + 1, column=1)
         total_cell.value = f"Total divergentes/não encontrados: {len(linhas)}"
         total_cell.font = Font(bold=True)
         total_cell.fill = no_fill
         total_cell.border = no_border

    if ws_out.max_row > 0 and ws_out.max_column > 0:
        dimensoes = ws_out.dimensions
        if dimensoes == 'A1:A1' and ws_out.cell(1,1).value is None:
             print("[LOG] Planilha de saída do repasse vazia, autofilter não aplicado.")
        else:
             ws_out.auto_filter.ref = ws_out.calculate_dimension()
             print(f"[LOG] Autofilter aplicado à planilha de saída do repasse. Ref: {ws_out.auto_filter.ref}")
    else:
         print("[LOG] Planilha de saída do repasse vazia, autofilter não aplicado.")

    stream_out = io.BytesIO()
    wb_out.save(stream_out)
    stream_out.seek(0)
    return stream_out
# =======================================================
# === FIM DA FUNÇÃO ATUALIZADA ===
# =======================================================

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

# =======================================================
# === FUNÇÃO PICK MONEY ATUALIZADA (DINÂMICA) ===
# =======================================================
def processar_repasse(diario_stream, sistema_stream, considerar_eql, considerar_parc, considerar_valor):
    """Lógica de conciliação Pick Money (Diário vs Sistema) - Lógica de Contador Dinâmico."""
    print(f"📘 [LOG] Início de processar_repasse (Pick Money) com Lógica Dinâmica: EQL={considerar_eql}, Parc={considerar_parc}, Valor={considerar_valor}")
    start_time = time.time()

    print("📘 [LOG] Carregando workbook 'Diário'...")
    wb_diario = load_workbook(diario_stream, data_only=True)
    ws_diario = wb_diario.worksheets[0]
    print(f"📗 [LOG] 'Diário' carregado ({ws_diario.max_row} linhas).")

    print("📘 [LOG] Carregando workbook 'Sistema'...")
    wb_sistema = load_workbook(sistema_stream, data_only=True)
    ws_sistema = wb_sistema.worksheets[0]
    print(f"📗 [LOG] 'Sistema' carregado ({ws_sistema.max_row} linhas).")

    print("📘 [LOG] Achando colunas (Pick Money)...")
    col_eq_diario = achar_coluna_flex(ws_diario, ["eql"])
    col_parcela_diario = achar_coluna_flex(ws_diario, ["parc", "parcela"])
    col_principal_diario = achar_coluna_flex(ws_diario, ["principal"])
    col_corrmonet_diario = achar_coluna_flex(ws_diario, ["correção monetária", "correção", "corrmonet", "correção monetáriaplano"]) # Adicionado "Correção MonetáriaPlano"

    col_eq_sistema = achar_coluna_flex(ws_sistema, ["eql"])
    col_parcela_sistema = achar_coluna_flex(ws_sistema, ["parc", "parcela"])
    col_valor_sistema = achar_coluna_flex(ws_sistema, ["valor"])

    missing_cols = []
    if not col_eq_diario: missing_cols.append("EQL (Diário)")
    if not col_parcela_diario: missing_cols.append("Parcela/Parc (Diário)")
    if not col_principal_diario: missing_cols.append("Principal (Diário)")
    if not col_corrmonet_diario: missing_cols.append("Correção (Diário)")
    if not col_eq_sistema: missing_cols.append("EQL (Sistema)")
    if not col_parcela_sistema: missing_cols.append("Parcela/Parc (Sistema)")
    if not col_valor_sistema: missing_cols.append("Valor (Sistema)")

    if missing_cols:
         error_msg = f"Colunas não encontradas: {', '.join(missing_cols)}."
         print(f"📕 [ERRO] {error_msg}")
         raise ValueError(error_msg)
         
    if not (considerar_eql or considerar_parc or considerar_valor):
        error_msg = "Pelo menos uma chave de comparação (EQL, Parcela, Valor) deve ser selecionada."
        print(f"📕 [ERRO] {error_msg}")
        raise ValueError(error_msg)

    print("📘 [LOG] Loop 1 (Pick Money): Contando 'Diário' (values_only)...")
    counter_diario = Counter()
    for row in ws_diario.iter_rows(min_row=2, values_only=True):
        eql = str(row[col_eq_diario - 1]).strip() if col_eq_diario <= len(row) and row[col_eq_diario - 1] else ""
        parcela = str(row[col_parcela_diario - 1]).strip() if col_parcela_diario <= len(row) and row[col_parcela_diario - 1] else ""
        principal = normalizar_valor(row[col_principal_diario - 1]) if col_principal_diario <= len(row) else 0.0
        correcao = normalizar_valor(row[col_corrmonet_diario - 1]) if col_corrmonet_diario <= len(row) else 0.0
        total = round(principal + correcao, 2)

        key_parts = []
        if considerar_eql: key_parts.append(eql)
        if considerar_parc: key_parts.append(parcela)
        if considerar_valor: key_parts.append(total)
        
        if len(key_parts) > 0 and any(k for k in key_parts if k): # Só conta se a chave não for vazia
            chave_completa = tuple(key_parts)
            counter_diario.update([chave_completa])

    print(f"📗 [LOG] Fim Loop 1. 'Diário' contado. {len(counter_diario)} chaves únicas. Tempo: {time.time() - start_time:.2f}s")

    print("📘 [LOG] Loop 2 (Pick Money): Contando 'Sistema'...")
    counter_sistema = Counter()
    for row in ws_sistema.iter_rows(min_row=2, values_only=True):
        eql = str(row[col_eq_sistema - 1]).strip() if col_eq_sistema <= len(row) and row[col_eq_sistema - 1] else ""
        parcela = str(row[col_parcela_sistema - 1]).strip() if col_parcela_sistema <= len(row) and row[col_parcela_sistema - 1] else ""
        valor = normalizar_valor(row[col_valor_sistema - 1]) if col_valor_sistema <= len(row) else 0.0

        key_parts = []
        if considerar_eql: key_parts.append(eql)
        if considerar_parc: key_parts.append(parcela)
        if considerar_valor: key_parts.append(valor)
        
        if len(key_parts) > 0 and any(k for k in key_parts if k):
            chave_completa = tuple(key_parts)
            counter_sistema.update([chave_completa])

    print(f"📗 [LOG] Fim Loop 2. 'Sistema' contado. {len(counter_sistema)} chaves únicas. Tempo: {time.time() - start_time:.2f}s")

    chaves_todas = set(counter_diario.keys()) | set(counter_sistema.keys())
    
    chaves_iguais_dict = {k: min(counter_diario[k], counter_sistema[k]) for k in chaves_todas if min(counter_diario[k], counter_sistema[k]) > 0}
    chaves_diario_apenas_dict = {k: counter_diario[k] - counter_sistema.get(k, 0) for k in chaves_todas if counter_diario[k] - counter_sistema.get(k, 0) > 0}
    chaves_sistema_apenas_dict = {k: counter_sistema[k] - counter_diario.get(k, 0) for k in chaves_todas if counter_sistema[k] - counter_diario.get(k, 0) > 0}
    
    iguais = []
    divergentes = []
    nao_encontrados_diario = []
    nao_encontrados_sistema = []

    print("📘 [LOG] Loop 3 (Pick Money): Classificando linhas do 'Diário'...")
    vistos_diario = Counter()
    if ws_diario.max_row >= 2:
        for row_cells in ws_diario.iter_rows(min_row=2):
            celula_eql = row_cells[col_eq_diario - 1] if col_eq_diario <= len(row_cells) else None
            celula_parcela = row_cells[col_parcela_diario - 1] if col_parcela_diario <= len(row_cells) else None
            celula_principal = row_cells[col_principal_diario - 1] if col_principal_diario <= len(row_cells) else None
            celula_correcao = row_cells[col_corrmonet_diario - 1] if col_corrmonet_diario <= len(row_cells) else None

            eql = str(celula_eql.value).strip() if celula_eql and celula_eql.value is not None else ""
            parcela = str(celula_parcela.value).strip() if celula_parcela and celula_parcela.value is not None else ""
            
            principal = normalizar_valor(celula_principal.value if celula_principal else None)
            correcao = normalizar_valor(celula_correcao.value if celula_correcao else None)
            total = round(principal + correcao, 2)
            
            key_parts = []
            if considerar_eql: key_parts.append(eql)
            if considerar_parc: key_parts.append(parcela)
            if considerar_valor: key_parts.append(total)
            if not key_parts or not any(k for k in key_parts if k): continue
            chave_completa = tuple(key_parts)

            vistos_diario.update([chave_completa])

            if chave_completa in chaves_iguais_dict and chaves_iguais_dict[chave_completa] > 0:
                iguais.append((row_cells, ""))
                chaves_iguais_dict[chave_completa] -= 1
            
            elif chave_completa in chaves_diario_apenas_dict and chaves_diario_apenas_dict[chave_completa] > 0:
                nao_encontrados_diario.append((row_cells, f"Não encontrado no 'Sistema' (Chave: {chave_completa})"))
                chaves_diario_apenas_dict[chave_completa] -= 1
            
            elif vistos_diario[chave_completa] > 1 and vistos_diario[chave_completa] > counter_diario.get(chave_completa, 0):
                 divergentes.append((row_cells, f"Duplicado no 'Diário' (Chave: {chave_completa})"))

    print("📘 [LOG] Loop 4 (Pick Money): Classificando linhas do 'Sistema'...")
    if ws_sistema.max_row >= 2:
        for row_cells in ws_sistema.iter_rows(min_row=2):
            celula_eql = row_cells[col_eq_sistema - 1] if col_eq_sistema <= len(row_cells) else None
            celula_parcela = row_cells[col_parcela_sistema - 1] if col_parcela_sistema <= len(row_cells) else None
            celula_valor = row_cells[col_valor_sistema - 1] if col_valor_sistema <= len(row_cells) else None
            
            eql = str(celula_eql.value).strip() if celula_eql and celula_eql.value is not None else ""
            parcela = str(celula_parcela.value).strip() if celula_parcela and celula_parcela.value is not None else ""
            
            valor = normalizar_valor(celula_valor.value if celula_valor else None)
            
            key_parts = []
            if considerar_eql: key_parts.append(eql)
            if considerar_parc: key_parts.append(parcela)
            if considerar_valor: key_parts.append(valor)
            if not key_parts or not any(k for k in key_parts if k): continue
            chave_completa = tuple(key_parts)
            
            if chave_completa in chaves_sistema_apenas_dict and chaves_sistema_apenas_dict[chave_completa] > 0:
                nao_encontrados_sistema.append((row_cells, f"Não encontrado no 'Diário' (Chave: {chave_completa})"))
                chaves_sistema_apenas_dict[chave_completa] -= 1

    print(f"📗 [LOG] Fim Comparação Pick Money. Tempo: {time.time() - start_time:.2f}s")
    
    print("📘 [LOG] Criando planilhas de saída (Pick Money)...")
    iguais_stream = criar_planilha_saida(iguais, ws_diario, incluir_status=False)
    divergentes_stream = criar_planilha_saida(divergentes, ws_diario, incluir_status=True)
    
    nao_encontrados_combinados = nao_encontrados_diario
    # Adiciona os não encontrados do sistema, usando o ws_sistema como modelo
    if nao_encontrados_sistema:
        # Cria uma planilha temporária SÓ para os do sistema, para usar o cabeçalho correto
        nao_encontrados_sistema_stream = criar_planilha_saida(nao_encontrados_sistema, ws_sistema, incluir_status=True)
        # TODO: Idealmente, iriamos mesclar os dados. Por simplicidade, usamos o modelo do Diário.
        # Isso significa que as colunas de "nao_encontrados_sistema" serão mapeadas para as colunas do "Diário"
        # O que pode ser confuso.
        # Solução Rápida: Adicionar como (None, status)
        nao_encontrados_combinados = nao_encontrados_diario
        for row_cells, status in nao_encontrados_sistema:
             nao_encontrados_combinados.append((None, status)) # Perde os dados da linha, mas evita confusão de colunas
             
    nao_encontrados_stream = criar_planilha_saida(nao_encontrados_combinados, ws_diario, incluir_status=True)


    timestamp_str = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
    pasta_saida = os.path.join(app.config['UPLOAD_FOLDER'], f"repasse_pick_money_{timestamp_str}")
    os.makedirs(pasta_saida, exist_ok=True)
    print(f"Pasta de saída criada: {pasta_saida}")

    try:
        salvar_stream_em_arquivo(iguais_stream, os.path.join(pasta_saida, "iguais.xlsx"))
        salvar_stream_em_arquivo(divergentes_stream, os.path.join(pasta_saida, "divergentes.xlsx"))
        salvar_stream_em_arquivo(nao_encontrados_stream, os.path.join(pasta_saida, "nao_encontrados.xlsx"))
        print(f"📗 [LOG] Arquivos Excel (Pick Money) salvos na pasta: {pasta_saida}")
    except Exception as e_save:
         print(f"📕 [ERRO] Falha ao salvar arquivos Excel (Pick Money) na pasta {pasta_saida}: {e_save}")
         raise

    count_nao_encontrados = len(nao_encontrados_combinados)
    print(f"✅ [LOG] Fim de processar_repasse (Pick Money). Totais: Iguais={len(iguais)}, Divergentes={len(divergentes)}, Não Encontrados={count_nao_encontrados}. Tempo total: {time.time() - start_time:.2f}s")
    return pasta_saida, len(iguais), len(divergentes), count_nao_encontrados

# =======================================================
# === FUNÇÃO ABRASMA ATUALIZADA (DINÂMICA) ===
# =======================================================
def processar_repasse_abrasma(anterior_stream, complementar_stream, considerar_eql, considerar_parc, considerar_valor):
    """Lógica de conciliação Abrasma (Anterior vs Complementar) - Lógica de Contador Dinâmico."""
    print(f"📘 [LOG] Início de processar_repasse_abrasma com Lógica Dinâmica: EQL={considerar_eql}, Parc={considerar_parc}, Valor={considerar_valor}")
    start_time = time.time()

    print("📘 [LOG] Carregando workbook 'Planilha Anterior'...")
    wb_ant = load_workbook(anterior_stream, data_only=True)
    ws_ant = wb_ant.worksheets[0]
    print(f"📗 [LOG] 'Anterior' carregada ({ws_ant.max_row} linhas).")

    print("📘 [LOG] Carregando workbook 'Planilha Complementar'...")
    wb_comp = load_workbook(complementar_stream, data_only=True)
    ws_comp = wb_comp.worksheets[0]
    print(f"📗 [LOG] 'Complementar' carregada ({ws_comp.max_row} linhas).")

    print("📘 [LOG] Achando colunas (Abrasma)...")
    col_eql_ant = achar_coluna_flex(ws_ant, ["eql"])
    col_parc_ant = achar_coluna_flex(ws_ant, ["parc", "parcela"])
    col_total_ant = achar_coluna_flex(ws_ant, ["total recebido"])

    col_eql_comp = achar_coluna_flex(ws_comp, ["eql"])
    col_parc_comp = achar_coluna_flex(ws_comp, ["parc", "parcela"])
    col_total_comp = achar_coluna_flex(ws_comp, ["total recebido"])

    print(f"📗 [LOG] Colunas encontradas: Anterior(EQL:{col_eql_ant}, Parc:{col_parc_ant}, Total:{col_total_ant}), Complementar(EQL:{col_eql_comp}, Parc:{col_parc_comp}, Total:{col_total_comp})")

    missing_cols = []
    if considerar_eql and not col_eql_ant: missing_cols.append("EQL (Anterior)")
    if considerar_parc and not col_parc_ant: missing_cols.append("Parc/Parcela (Anterior)")
    if considerar_valor and not col_total_ant: missing_cols.append("Total Recebido (Anterior)")
    if considerar_eql and not col_eql_comp: missing_cols.append("EQL (Complementar)")
    if considerar_parc and not col_parc_comp: missing_cols.append("Parc/Parcela (Complementar)")
    if considerar_valor and not col_total_comp: missing_cols.append("Total Recebido (Complementar)")

    if missing_cols:
         error_msg = f"Colunas selecionadas não encontradas: {', '.join(missing_cols)}. Verifique os nomes nos cabeçalhos."
         print(f"📕 [ERRO] {error_msg}")
         raise ValueError(error_msg)
         
    if not (considerar_eql or considerar_parc or considerar_valor):
        error_msg = "Pelo menos uma chave de comparação (EQL, Parcela, Valor) deve ser selecionada."
        print(f"📕 [ERRO] {error_msg}")
        raise ValueError(error_msg)

    print("📘 [LOG] Loop 1 (Abrasma): Contando 'Anterior' (values_only)...")
    counter_ant = Counter()
    for row in ws_ant.iter_rows(min_row=2, values_only=True):
        key_parts = []
        if considerar_eql:
            eql = str(row[col_eql_ant - 1]).strip() if col_eql_ant <= len(row) and row[col_eql_ant - 1] else ""
            key_parts.append(eql)
        if considerar_parc:
            parc = str(row[col_parc_ant - 1]).strip() if col_parc_ant <= len(row) and row[col_parc_ant - 1] else ""
            key_parts.append(parc)
        if considerar_valor:
            total = normalizar_valor(row[col_total_ant - 1]) if col_total_ant <= len(row) else 0.0
            key_parts.append(total)

        if len(key_parts) > 0 and any(k for k in key_parts if k):
            chave_completa = tuple(key_parts)
            counter_ant.update([chave_completa])

    print(f"📗 [LOG] Fim Loop 1. 'Anterior' contada. {len(counter_ant)} chaves únicas. Tempo: {time.time() - start_time:.2f}s")

    print("📘 [LOG] Loop 2 (Abrasma): Contando 'Complementar'...")
    counter_comp = Counter()
    for row in ws_comp.iter_rows(min_row=2, values_only=True):
        key_parts = []
        if considerar_eql:
            eql = str(row[col_eql_comp - 1]).strip() if col_eql_comp <= len(row) and row[col_eql_comp - 1] else ""
            key_parts.append(eql)
        if considerar_parc:
            parc = str(row[col_parc_comp - 1]).strip() if col_parc_comp <= len(row) and row[col_parc_comp - 1] else ""
            key_parts.append(parc)
        if considerar_valor:
            total = normalizar_valor(row[col_total_comp - 1]) if col_total_comp <= len(row) else 0.0
            key_parts.append(total)
        
        if len(key_parts) > 0 and any(k for k in key_parts if k):
            chave_completa = tuple(key_parts)
            counter_comp.update([chave_completa])

    print(f"📗 [LOG] Fim Loop 2. 'Complementar' contada. {len(counter_comp)} chaves únicas. Tempo: {time.time() - start_time:.2f}s")

    chaves_todas = set(counter_ant.keys()) | set(counter_comp.keys())
    
    chaves_iguais_dict = {k: min(counter_ant[k], counter_comp[k]) for k in chaves_todas if min(counter_ant[k], counter_comp[k]) > 0}
    chaves_ant_apenas_dict = {k: counter_ant[k] - counter_comp.get(k, 0) for k in chaves_todas if counter_ant[k] - counter_comp.get(k, 0) > 0}
    chaves_comp_apenas_dict = {k: counter_comp[k] - counter_ant.get(k, 0) for k in chaves_todas if counter_comp[k] - counter_ant.get(k, 0) > 0}
    
    iguais = []
    divergentes = []
    nao_encontrados_ant = []
    nao_encontrados_comp = []

    print("📘 [LOG] Loop 3 (Abrasma): Classificando linhas da 'Anterior'...")
    vistos_ant = Counter()
    if ws_ant.max_row >= 2:
        for row_cells in ws_ant.iter_rows(min_row=2):
            key_parts = []
            
            if considerar_eql:
                celula_eql = row_cells[col_eql_ant - 1] if col_eql_ant <= len(row_cells) else None
                eql = str(celula_eql.value).strip() if celula_eql and celula_eql.value is not None else ""
                key_parts.append(eql)
            if considerar_parc:
                celula_parc = row_cells[col_parc_ant - 1] if col_parc_ant <= len(row_cells) else None
                parc = str(celula_parc.value).strip() if celula_parc and celula_parc.value is not None else ""
                key_parts.append(parc)
            if considerar_valor:
                celula_total = row_cells[col_total_ant - 1] if col_total_ant <= len(row_cells) else None
                total = normalizar_valor(celula_total.value if celula_total else None)
                key_parts.append(total)

            if not key_parts or not any(k for k in key_parts if k): continue
            chave_completa = tuple(key_parts)
            vistos_ant.update([chave_completa])

            if chave_completa in chaves_iguais_dict and chaves_iguais_dict[chave_completa] > 0:
                iguais.append((row_cells, ""))
                chaves_iguais_dict[chave_completa] -= 1
            
            elif chave_completa in chaves_ant_apenas_dict and chaves_ant_apenas_dict[chave_completa] > 0:
                nao_encontrados_ant.append((row_cells, f"Não encontrado na 'Complementar' (Chave: {chave_completa})"))
                chaves_ant_apenas_dict[chave_completa] -= 1
            
            elif vistos_ant[chave_completa] > 1 and vistos_ant[chave_completa] > counter_ant.get(chave_completa, 0):
                 divergentes.append((row_cells, f"Duplicado na 'Anterior' (Chave: {chave_completa})"))
                      
    print("📘 [LOG] Loop 4 (Abrasma): Classificando linhas da 'Complementar'...")
    if ws_comp.max_row >= 2:
        for row_cells_comp in ws_comp.iter_rows(min_row=2):
            key_parts = []
            
            if considerar_eql:
                celula_eql = row_cells_comp[col_eql_comp - 1] if col_eql_comp <= len(row_cells_comp) else None
                eql = str(celula_eql.value).strip() if celula_eql and celula_eql.value is not None else ""
                key_parts.append(eql)
            if considerar_parc:
                celula_parc = row_cells_comp[col_parc_comp - 1] if col_parc_comp <= len(row_cells_comp) else None
                parc = str(celula_parc.value).strip() if celula_parc and celula_parc.value is not None else ""
                key_parts.append(parc)
            if considerar_valor:
                celula_total = row_cells_comp[col_total_comp - 1] if col_total_comp <= len(row_cells_comp) else None
                total = normalizar_valor(celula_total.value if celula_total else None)
                key_parts.append(total)

            if not key_parts or not any(k for k in key_parts if k): continue
            chave_completa = tuple(key_parts)
            
            if chave_completa in chaves_comp_apenas_dict and chaves_comp_apenas_dict[chave_completa] > 0:
                nao_encontrados_comp.append((row_cells_comp, f"Não encontrado na 'Anterior' (Chave: {chave_completa})"))
                chaves_comp_apenas_dict[chave_completa] -= 1

    print(f"📗 [LOG] Fim Comparação Abrasma. Tempo: {time.time() - start_time:.2f}s")

    print("📘 [LOG] Criando planilhas de saída (Abrasma)...")
    iguais_stream = criar_planilha_saida(iguais, ws_ant, incluir_status=False)
    divergentes_stream = criar_planilha_saida(divergentes, ws_ant, incluir_status=True)
    nao_encontrados_combinados = nao_encontrados_ant + nao_encontrados_comp
    nao_encontrados_stream = criar_planilha_saida(nao_encontrados_combinados, ws_ant, incluir_status=True)

    timestamp_str = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
    pasta_saida = os.path.join(app.config['UPLOAD_FOLDER'], f"repasse_abrasma_{timestamp_str}")
    os.makedirs(pasta_saida, exist_ok=True)
    print(f"Pasta de saída criada: {pasta_saida}")

    try:
        salvar_stream_em_arquivo(iguais_stream, os.path.join(pasta_saida, "iguais.xlsx"))
        salvar_stream_em_arquivo(divergentes_stream, os.path.join(pasta_saida, "divergentes.xlsx"))
        salvar_stream_em_arquivo(nao_encontrados_stream, os.path.join(pasta_saida, "nao_encontrados.xlsx"))
        print(f"📗 [LOG] Arquivos Excel (Abrasma) salvos na pasta: {pasta_saida}")
    except Exception as e_save:
         print(f"📕 [ERRO] Falha ao salvar arquivos Excel (Abrasma) na pasta {pasta_saida}: {e_save}")
         raise

    count_nao_encontrados = len(nao_encontrados_combinados)
    print(f"✅ [LOG] Fim de processar_repasse (Abrasma). Totais: Iguais={len(iguais)}, Divergentes={len(divergentes)}, Não Encontrados={count_nao_encontrados}. Tempo total: {time.time() - start_time:.2f}s")
    return pasta_saida, len(iguais), len(divergentes), count_nao_encontrados


# ==== ROTAS FLASK ====

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
                error_msg = ("Para o modo 'Boleto', o nome do arquivo precisa terminar com um código de empreendimento válido (ex: 'Extrato_RSCI.pdf'). "
                             "Verifique o nome do arquivo ou selecione outro modo de análise.")
                return manual_render_template('error.html', status_code=400,
                    error_title="Empreendimento não identificado (Modo Boleto)", error_message=error_msg)

        elif modo_separacao in ['debito_credito', 'ccb_realiza']:
             if detectar_emp_por_nome_arquivo(file.filename) and modo_separacao == 'debito_credito':
                  error_msg = ("Este arquivo parece ser do tipo 'Boleto' (termina com código de empreendimento), mas o modo 'Débito/Crédito' foi selecionado. "
                               "Por favor, use o modo 'Boleto' ou renomeie o arquivo se ele não for específico de um empreendimento.")
                  return manual_render_template('error.html', status_code=400,
                                                error_title="Modo de Análise Incorreto?", error_message=error_msg)

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

        df_todas_filtrado = df_todas_raw.copy()
        if not df_todas_filtrado.empty:
            parcelas_para_remover = ['TOTAL A PAGAR', 'DESCONTO', 'DÉBITOS DO MÊS ANTERIOR', 'ENCARGOS POR ATRASO', 'PAGAMENTO EFETUADO', 'DÉBITOS DO MÊS']
            df_todas_filtrado = df_todas_filtrado[~df_todas_filtrado['Parcela'].astype(str).str.strip().str.upper().isin(parcelas_para_remover)]
            df_todas_filtrado = df_todas_filtrado[~df_todas_filtrado['Parcela'].astype(str).str.strip().str.upper().str.startswith('TOTAL BANCO')]
        print("Parcelas indesejadas filtradas da aba 'Todas_Parcelas_Extraidas'.")

        output = io.BytesIO()
        dfs_to_excel = {"Divergencias": df_div, "Cobertura_Analise": df_cov, "Todas_Parcelas_Extraidas": df_todas_filtrado}
        print("Gerando arquivo Excel...")
        formatar_excel(output, dfs_to_excel) # Chama a função formatar_excel com autofiltro
        output.seek(0)
        print("Arquivo Excel gerado em memória.")

        base_name = os.path.splitext(file.filename)[0]
        report_filename = f"relatorio_{modo_separacao}_{base_name}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)

        try:
            with open(report_path, 'wb') as f: f.write(output.getvalue())
            print(f"Relatório salvo em: {report_path}")
        except Exception as e_save:
            print(f"Erro ao salvar o arquivo Excel em {report_path}: {e_save}")

        nao_classificados = 0
        if not df_cov.empty and 'Empreendimento' in df_cov.columns:
            nao_classificados = df_cov[df_cov['Empreendimento'] == 'NAO_CLASSIFICADO'].shape[0]
            if nao_classificados > 0: print(f"[AVISO] {nao_classificados} registros não classificados.")

        print("Renderizando página de resultados...")
        return manual_render_template('results.html',
            divergencias_json=df_div.to_json(orient='split', index=False, date_format='iso') if not df_div.empty else 'null',
            total_lotes=len(df_cov),
            total_divergencias=len(df_div),
            nao_classificados=nao_classificados,
            download_url=url_for('download_file', filename=report_filename),
            modo_usado=modo_separacao.replace('_', '/').upper()
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
    modo_separacao = request.form.get('modo_separacao_comp', 'boleto')

    if file_ant.filename == '' or file_atu.filename == '':
        return manual_render_template('error.html', status_code=400,
            error_title="Arquivos faltando",
            error_message="Selecione os dois arquivos PDF para comparar.")

    if not file_ant.filename.lower().endswith('.pdf') or not file_atu.filename.lower().endswith('.pdf'):
         return manual_render_template('error.html', status_code=400,
            error_title="Tipo de Arquivo Inválido",
            error_message="Por favor, envie apenas arquivos no formato PDF para comparação.")


    try:
        emp_fixo_boleto = None
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
            emp_fixo_boleto = emp_ant

        elif modo_separacao in ['debito_credito', 'ccb_realiza']:
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

        report_filename = f"comparativo_{modo_separacao}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
        try:
            with open(report_path, 'wb') as f:
                f.write(output.getvalue())
            print(f"Relatório comparativo salvo em: {report_path}")
        except Exception as e_save:
             print(f"Erro ao salvar o arquivo Excel comparativo em {report_path}: {e_save}")


        resumo_dict_lotes = {}
        resumo_dict_totais = {}
        if not df_resumo_completo.empty:
             resumo_dict_lotes = pd.Series(df_resumo_completo.set_index(' ')['LOTES']).to_dict()
             resumo_dict_totais = pd.Series(df_resumo_completo.set_index(' ')['TOTAIS']).map('{:,.2f}'.format).to_dict()


        print("Renderizando página de resultados da comparação...")
        return manual_render_template('compare_results.html',
             resumo_lotes_mes_anterior=resumo_dict_lotes.get('Lotes Mês Anterior', 0),
             resumo_lotes_mes_atual=resumo_dict_lotes.get('Lotes Mês Atual', 0),
             resumo_lotes_adicionados=resumo_dict_lotes.get('Lotes Adicionados', 0),
             resumo_lotes_removidos=resumo_dict_lotes.get('Lotes Removidos', 0),
             resumo_parcelas_com_valor_alterado=resumo_dict_lotes.get('Parcelas com Valor Alterado', 0),

             total_mes_anterior_str=resumo_dict_totais.get('Lotes Mês Anterior', '0.00'),
             total_mes_atual_str=resumo_dict_totais.get('Lotes Mês Atual', '0.00'),
             total_adicionados_str=resumo_dict_totais.get('Lotes Adicionados', '0.00'),
             total_removidos_str=resumo_dict_totais.get('Lotes Removidos', '0.00'),
             total_diferencas_str=resumo_dict_totais.get('Parcelas com Valor Alterado', '0.00'),

            divergencias_json=df_divergencias.to_json(orient='split', index=False, date_format='iso') if not df_divergencias.empty else 'null',
            adicionados_json=df_adicionados.to_json(orient='split', index=False, date_format='iso') if not df_adicionados.empty else 'null',
            removidos_json=df_removidos.to_json(orient='split', index=False, date_format='iso') if not df_removidos.empty else 'null',

            download_url=url_for('download_file', filename=report_filename),
            modo_usado=modo_separacao.replace('_', '/').upper()
        )


    except Exception as e:
        print(f"📕 [ERRO FATAL] Erro inesperado na rota /compare: {e}")
        traceback.print_exc()
        error_details = f"{type(e).__name__}: {e}"
        return manual_render_template('error.html', status_code=500,
            error_title="Erro inesperado na comparação",
            error_message=f"Ocorreu um erro grave durante a comparação dos arquivos. Detalhes: {error_details}")


@app.route('/repasse', methods=['POST'])
def repasse_file():
    """Rota para a conciliação Pick Money (Diário vs Sistema)"""
    print("\n--- RECEIVED REQUEST /repasse (Pick Money) ---")
    start_time_route = time.time()

    if 'diario_file' not in request.files or 'sistema_file' not in request.files:
        print("📕 [ERRO] Arquivos 'diario_file' ou 'sistema_file' faltando.")
        return manual_render_template('error.html', status_code=400,
            error_title="Arquivos faltando",
            error_message="Você precisa enviar os arquivos 'Diário' e 'Sistema' para a conciliação Pick Money.")

    file_diario = request.files['diario_file']
    file_sistema = request.files['sistema_file']

    if file_diario.filename == '' or file_sistema.filename == '':
        print("📕 [ERRO] Nomes dos arquivos Excel (Pick Money) estão vazios.")
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

    print(f"📘 [LOG] Recebidos (Pick Money): {file_diario.filename}, {file_sistema.filename}")

    try:
        # Pega os valores dos checkboxes
        considerar_eql = request.form.get('considerar_eql_pm') == 'on'
        considerar_parc = request.form.get('considerar_parc_pm') == 'on'
        considerar_valor = request.form.get('considerar_valor_pm') == 'on'
        
        diario_stream = io.BytesIO(file_diario.read())
        sistema_stream = io.BytesIO(file_sistema.read())
        print(f"📘 [LOG] Arquivos Excel (Pick Money) lidos em memória. Tempo: {time.time() - start_time_route:.2f}s")

        # Chama a função de processamento com os novos parâmetros
        pasta_saida, count_iguais, count_divergentes, count_nao_encontrados = processar_repasse(
            diario_stream, sistema_stream, 
            considerar_eql, considerar_parc, considerar_valor
        )

        print(f"📘 [LOG] Processamento (Pick Money) concluído. Criando ZIP da pasta '{pasta_saida}'...")
        zip_stream = io.BytesIO()
        timestamp_str = os.path.basename(pasta_saida).replace('repasse_pick_money_', '')

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
        print(f"📗 [LOG] ZIP (Pick Money) criado em memória.")

        report_filename = f"repasse_pick_money_conciliado_{timestamp_str}.zip"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
        try:
            with open(report_path, 'wb') as f:
                f.write(zip_stream.getvalue())
            print(f"📗 [LOG] Arquivo ZIP (Pick Money) salvo para download em {report_path}.")
        except Exception as e_save:
             print(f"📕 [ERRO] Erro ao salvar o arquivo ZIP (Pick Money) em {report_path}: {e_save}")
             raise e_save

        print("✅ [LOG] Enviando resposta (Pick Money) para 'repasse_results.html'")
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
             error_title="Erro na Conciliação (Pick Money) - Colunas Não Encontradas",
             error_message=f"Verifique os nomes das colunas nas planilhas. Detalhes: {ve}")
    except Exception as e:
        print(f"📕 [ERRO FATAL] Erro inesperado na rota /repasse (Pick Money): {e}")
        traceback.print_exc()
        error_details = f"{type(e).__name__}: {e}"
        return manual_render_template('error.html', status_code=500,
            error_title="Erro inesperado na conciliação (Pick Money)",
            error_message=f"Ocorreu um erro grave durante a análise. Detalhes: {error_details}")


@app.route('/repasse_abrasma', methods=['POST'])
def repasse_abrasma_file():
    """Rota para a conciliação Abrasma (Anterior vs Complementar)"""
    print("\n--- RECEIVED REQUEST /repasse_abrasma ---")
    start_time_route = time.time()

    if 'anterior_file' not in request.files or 'complementar_file' not in request.files:
        print("📕 [ERRO] Arquivos 'anterior_file' ou 'complementar_file' faltando.")
        return manual_render_template('error.html', status_code=400,
            error_title="Arquivos faltando",
            error_message="Você precisa enviar a 'Planilha Anterior' e a 'Planilha Complementar' para a conciliação Abrasma.")

    file_ant = request.files['anterior_file']
    file_comp = request.files['complementar_file']

    if file_ant.filename == '' or file_comp.filename == '':
        print("📕 [ERRO] Nomes dos arquivos Excel (Abrasma) estão vazios.")
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

    print(f"📘 [LOG] Recebidos (Abrasma): {file_ant.filename}, {file_comp.filename}")

    try:
        # Pega os valores dos checkboxes
        considerar_eql = request.form.get('considerar_eql_ab') == 'on'
        considerar_parc = request.form.get('considerar_parc_ab') == 'on'
        considerar_valor = request.form.get('considerar_valor_ab') == 'on'

        anterior_stream = io.BytesIO(file_ant.read())
        complementar_stream = io.BytesIO(file_comp.read())
        print(f"📘 [LOG] Arquivos Excel (Abrasma) lidos em memória. Tempo: {time.time() - start_time_route:.2f}s")

        # Chama a função de processamento Abrasma com os novos parâmetros
        pasta_saida, count_iguais, count_divergentes, count_nao_encontrados = processar_repasse_abrasma(
            anterior_stream, complementar_stream,
            considerar_eql, considerar_parc, considerar_valor
        )

        print(f"📘 [LOG] Processamento (Abrasma) concluído. Criando ZIP da pasta '{pasta_saida}'...")
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
        print(f"📗 [LOG] ZIP (Abrasma) criado em memória.")

        report_filename = f"repasse_abrasma_conciliado_{timestamp_str}.zip"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
        try:
            with open(report_path, 'wb') as f:
                f.write(zip_stream.getvalue())
            print(f"📗 [LOG] Arquivo ZIP (Abrasma) salvo para download em {report_path}.")
        except Exception as e_save:
             print(f"📕 [ERRO] Erro ao salvar o arquivo ZIP (Abrasma) em {report_path}: {e_save}")
             raise e_save

        print("✅ [LOG] Enviando resposta (Abrasma) para 'repasse_results.html'")
        return manual_render_template('repasse_results.html',
            count_iguais=count_iguais,
            count_divergentes=count_divergentes,
            count_nao_encontrados=count_nao_encontrados,
            download_url=url_for('download_file', filename=report_filename)
        )

    except ValueError as ve:
         print(f"📕 [ERRO VALIDAÇÃO Abrasma] {ve}")
         traceback.print_exc()
         return manual_render_template('error.html', status_code=400,
             error_title="Erro na Conciliação (Abrasma) - Colunas Não Encontradas",
             error_message=f"Verifique os nomes das colunas (EQL, Parc, Total Recebido). Detalhes: {ve}")
    except Exception as e:
        print(f"📕 [ERRO FATAL] Erro inesperado na rota /repasse_abrasma: {e}")
        traceback.print_exc()
        error_details = f"{type(e).__name__}: {e}"
        return manual_render_template('error.html', status_code=500,
            error_title="Erro inesperado na conciliação (Abrasma)",
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
    # Verifica variável de ambiente FLASK_DEBUG para modo debug
    debug_mode = os.environ.get('FLASK_DEBUG') == '1'
    # Usa host='0.0.0.0' para ser acessível na rede local ou pelo Render
    print(f"Executando em http://0.0.0.0:{port} (debug={debug_mode})")
    # threaded=True pode ajudar a evitar timeouts em requisições longas localmente
    app.run(debug=debug_mode, host='0.0.0.0', port=port, threaded=True)
