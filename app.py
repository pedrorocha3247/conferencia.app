import os
from flask import Flask, render_template

app = Flask(__name__)

@app.route('/')
def index():
    # --- INÍCIO DO DIAGNÓSTICO AVANÇADO ---
    template_path = os.path.join(os.getcwd(), 'templates', 'index.html')
    
    print("--- INICIANDO DIAGNÓSTICO AVANÇADO ---")
    
    file_exists = os.path.exists(template_path)
    print(f"DIAGNÓSTICO: O arquivo '{template_path}' existe? {file_exists}")
    
    if file_exists:
        try:
            # 1. Verifica o tamanho do arquivo
            file_size = os.path.getsize(template_path)
            print(f"DIAGNÓSTICO: Tamanho do arquivo: {file_size} bytes.")
            
            # 2. Tenta ler os primeiros 100 caracteres do arquivo
            with open(template_path, 'r', encoding='utf-8') as f:
                content_preview = f.read(100)
                print(f"DIAGNÓSTICO: Prévia do conteúdo: '{content_preview}...'")
            
            # Se o arquivo estiver vazio, mostra uma mensagem de erro clara no navegador
            if file_size == 0:
                return "<h1>Erro Confirmado: Arquivo Vazio</h1><p>O arquivo 'index.html' no servidor existe, mas seu conteúdo está em branco (0 bytes). Por favor, verifique o arquivo no seu repositório Git e envie uma nova versão com o conteúdo correto.</p>"

        except Exception as e:
            print(f"DIAGNÓSTICO: ERRO AO LER O ARQUIVO: {e}")
            return f"<h1>Erro ao ler o arquivo</h1><p>{e}</p>"
    else:
        return "<h1>Erro Crítico: Arquivo Não Encontrado</h1><p>O arquivo de template não foi encontrado no caminho '{template_path}'.</p>"

    print("--- DIAGNÓSTICO CONCLUÍDO ---")
    # --- FIM DO DIAGNÓSTICO ---
    
    # Se passar por todas as verificações, tenta renderizar o template
    return render_template('index.html')

@app.route('/healthz')
def healthz():
    return "OK", 200

if __name__ == '__main__':
    app.run(debug=True)
