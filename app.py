# -*- coding: utf-8 -*-

import os
import sys
import re
import unicodedata
import io
import fitz  # PyMuPDF
import pandas as pd
from collections import OrderedDict
from flask import Flask, render_template, request, send_file, url_for, redirect, flash
import traceback
import json
import zipfile
import tempfile
import shutil
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user

# --- Configuração Inicial do App ---
app = Flask(__name__)
app.config['SECRET_KEY'] = 'uma-chave-secreta-muito-forte'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///app.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# --- Configuração do Banco de Dados e Login ---
db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

# --- Modelos do Banco de Dados ---
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(150), unique=True, nullable=False)
    password = db.Column(db.String(150), nullable=False)

class Config(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    key = db.Column(db.String(50), unique=True)
    value = db.Column(db.Text) # Storing JSON as text

# --- Funções de Carregamento de Config e User ---
@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

def get_config_from_db():
    emp_map_db = Config.query.filter_by(key='EMP_MAP').first()
    base_fixos_db = Config.query.filter_by(key='BASE_FIXOS').first()
    
    EMP_MAP = json.loads(emp_map_db.value) if emp_map_db else {}
    BASE_FIXOS = json.loads(base_fixos_db.value) if base_fixos_db else {}
    
    # Converte os valores de lista de strings para lista de floats em BASE_FIXOS
    for key, val in BASE_FIXOS.items():
        if isinstance(val, list):
            BASE_FIXOS[key] = [float(i) for i in val]
            
    return EMP_MAP, BASE_FIXOS

# --- Constantes e Funções de Lógica (semelhantes às versões anteriores) ---
# PADRAO_LOTE, CODIGO_EMP_MAP, etc.
# extrair_texto_pdf, limpar_rotulo, fatiar_blocos, extrair_parcelas...
# processar_pdf_validacao, processar_comparativo, formatar_excel...
# (O corpo dessas funções permanece o mesmo, mas agora `processar_pdf_validacao` 
#  receberá EMP_MAP e BASE_FIXOS como argumentos em vez de usar globais)

def processar_pdf_validacao(texto_pdf: str, modo_separacao: str, emp_fixo_boleto: str = None):
    EMP_MAP, BASE_FIXOS = get_config_from_db() # Carrega do DB
    # ... resto da função ...
    # a função fixos_do_emp também precisaria receber BASE_FIXOS
    return df_todas, df_cov, df_div

# --- Rotas da Aplicação ---
@app.route('/')
def index():
    return render_template('index.html')

# (Rotas /upload e /compare agora usam redirect com erro)
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'pdf_file' not in request.files or request.files['pdf_file'].filename == '':
        return redirect(url_for('index', error="Nenhum arquivo enviado."))
    # ... lógica de processamento ...
    # Em caso de erro, usar: return redirect(url_for('index', error="Sua mensagem de erro aqui"))
    return render_template('results.html', empreendimentos=list(df_cov['Empreendimento'].unique()), ...)

@app.route('/compare', methods=['POST'])
def compare_files():
    # ... lógica similar de validação e erro com redirect ...
    return render_template('compare_results.html', empreendimentos=list(df_adicionados['Empreendimento'].unique()), ...)

# --- NOVO: Rota de Processamento em Lote ---
@app.route('/batch_upload', methods=['POST'])
def batch_upload():
    if 'zip_file' not in request.files or request.files['zip_file'].filename == '':
        return redirect(url_for('index', error="Nenhum arquivo .zip enviado."))
    
    zip_file = request.files['zip_file']
    if not zip_file.filename.lower().endswith('.zip'):
        return redirect(url_for('index', error="Formato de arquivo inválido. Por favor, envie um .zip."))

    temp_dir = tempfile.mkdtemp()
    try:
        zip_path = os.path.join(temp_dir, zip_file.filename)
        zip_file.save(zip_path)

        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(temp_dir)

        all_divergencias = []
        for item in os.listdir(temp_dir):
            if item.lower().endswith('.pdf'):
                file_path = os.path.join(temp_dir, item)
                with open(file_path, 'rb') as f:
                    pdf_stream = f.read()
                    texto_pdf = extrair_texto_pdf(pdf_stream)
                    _, _, df_div = processar_pdf_validacao(texto_pdf, 'debito_credito') # ou modo boleto
                    if not df_div.empty:
                        df_div['Arquivo Origem'] = item
                        all_divergencias.append(df_div)
        
        if not all_divergencias:
            return render_template('batch_results.html', total_files=len(all_divergencias), divergencias_json='null')

        df_final = pd.concat(all_divergencias, ignore_index=True)
        return render_template('batch_results.html', total_files=len(all_divergencias), divergencias_json=df_final.to_json(orient='split'))

    finally:
        shutil.rmtree(temp_dir)

# --- NOVO: Rotas de Admin e Login ---
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        user = User.query.filter_by(username=username).first()
        if user and user.password == password: # Em produção, usar senhas com hash!
            login_user(user)
            return redirect(url_for('admin'))
        flash('Login inválido')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('index'))

@app.route('/admin', methods=['GET', 'POST'])
@login_required
def admin():
    emp_map_obj = Config.query.filter_by(key='EMP_MAP').first()
    base_fixos_obj = Config.query.filter_by(key='BASE_FIXOS').first()

    if request.method == 'POST':
        # ... Lógica para pegar os dados do form e salvar no banco de dados ...
        # (Ex: new_emp_map = {}, for key in request.form: ... new_emp_map[key] = ...)
        # emp_map_obj.value = json.dumps(new_emp_map)
        # db.session.commit()
        flash('Configurações salvas com sucesso!')
        return redirect(url_for('admin'))

    emp_map = json.loads(emp_map_obj.value) if emp_map_obj else {}
    base_fixos = json.loads(base_fixos_obj.value) if base_fixos_obj else {}
    return render_template('admin.html', emp_map=emp_map, base_fixos=base_fixos)

# --- Setup Inicial do Banco de Dados ---
def setup_database(app):
    with app.app_context():
        db.create_all()
        # Cria usuário admin padrão e carrega config inicial se não existir
        if not User.query.filter_by(username='admin').first():
            db.session.add(User(username='admin', password='password')) # Mude essa senha!
            db.session.commit()
        if not Config.query.filter_by(key='EMP_MAP').first():
            with open('config.json', 'r', encoding='utf-8') as f:
                config_data = json.load()
            emp_map_json = json.dumps(config_data['EMP_MAP'])
            base_fixos_json = json.dumps(config_data['BASE_FIXOS'])
            db.session.add(Config(key='EMP_MAP', value=emp_map_json))
            db.session.add(Config(key='BASE_FIXOS', value=base_fixos_json))
            db.session.commit()

if __name__ == '__main__':
    setup_database(app)
    app.run(debug=True)

