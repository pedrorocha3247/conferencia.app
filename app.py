# app.py de TESTE para isolar o problema
from flask import Flask, render_template

app = Flask(__name__)

@app.route('/')
def index():
    # A única função deste código é tentar carregar a página inicial.
    # Não há banco de dados nem login.
    try:
        return render_template('index.html')
    except Exception as e:
        # Se ele não conseguir encontrar ou ler o index.html, ele retornará o erro.
        return f"<h1>Ocorreu um erro ao tentar carregar o template:</h1><p>{e}</p>"

@app.route('/healthz')
def healthz():
    return "OK", 200

if __name__ == '__main__':
    app.run(debug=True)
