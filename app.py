import os
from flask import Flask, make_response

app = Flask(__name__)

@app.route('/')
def index():
    # Caminho para o arquivo de template
    template_path = os.path.join(os.getcwd(), 'templates', 'index.html')
    
    try:
        # Abre o arquivo, lê todo o seu conteúdo como texto
        with open(template_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        # Cria uma resposta HTTP manualmente com o conteúdo lido
        response = make_response(html_content)
        response.headers['Content-Type'] = 'text/html'
        return response

    except Exception as e:
        # Se até mesmo a leitura manual falhar, o problema é extremamente incomum
        print(f"ERRO CRÍTICO FINAL AO LER O ARQUIVO: {e}")
        error_html = f"<h1>Erro Inexplicável na Leitura do Arquivo</h1><p>Não foi possível ler o conteúdo do template diretamente: {e}</p>"
        return make_response(error_html, 500)

@app.route('/healthz')
def healthz():
    """Rota para o Health Check do Render."""
    return "OK", 200

if __name__ == '__main__':
    app.run(debug=True)
