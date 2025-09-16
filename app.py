import os
from flask import Flask, render_template

app = Flask(__name__)

@app.route('/')
def index():
    # --- Início do Código de Diagnóstico ---
    # Vamos escrever informações importantes nos logs do Render
    
    # 1. Qual é o diretório de trabalho atual do app?
    cwd = os.getcwd()
    print(f"DIAGNÓSTICO: Diretório de Trabalho Atual: {cwd}")

    # 2. O que existe dentro deste diretório?
    try:
        dir_contents = os.listdir(cwd)
        print(f"DIAGNÓSTICO: Conteúdo do Diretório: {dir_contents}")
    except Exception as e:
        print(f"DIAGNÓSTICO: Erro ao listar diretório: {e}")

    # 3. A pasta 'templates' existe aqui?
    templates_dir_path = os.path.join(cwd, 'templates')
    templates_dir_exists = os.path.exists(templates_dir_path)
    print(f"DIAGNÓSTICO: A pasta 'templates' existe? {templates_dir_exists}")
    
    # 4. O arquivo 'templates/index.html' existe?
    template_file_path = os.path.join(templates_dir_path, 'index.html')
    template_file_exists = os.path.exists(template_file_path)
    print(f"DIAGNÓSTICO: O arquivo 'index.html' existe? {template_file_exists}")

    # Se o arquivo não for encontrado, exibe uma mensagem de erro clara na tela
    if not template_file_exists:
        return f"""
        <h1>Erro Crítico de Carregamento</h1>
        <p>O arquivo de template não foi encontrado no caminho esperado.</p>
        <p><b>Diretório verificado:</b> {template_file_path}</p>
        <p>Verifique a estrutura de pastas do seu projeto e o nome do arquivo.</p>
        """
    # --- Fim do Código de Diagnóstico ---
    
    return render_template('index.html')

@app.route('/healthz')
def healthz():
    return "OK", 200

if __name__ == '__main__':
    app.run(debug=True)
