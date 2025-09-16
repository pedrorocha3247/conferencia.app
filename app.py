# -*- coding: utf-8 -*-

import os
import sys
import re
import unicodedata
import io
import json
import zipfile
import tempfile
import shutil
import traceback
from collections import OrderedDict

import fitz  # PyMuPDF
import pandas as pd
from flask import (Flask, render_template, request, send_file, url_for,
                   redirect, flash)
from flask_sqlalchemy import SQLAlchemy
from flask_login import (LoginManager, UserMixin, login_user, logout_user,
                         login_required, current_user)
from werkzeug.security import generate_password_hash, check_password_hash

# --- 1. Configuração Inicial do App ---
app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-prod')
app.config['UPLOAD_FOLDER'] = 'uploads'
# Usa a variável de ambiente DATABASE_URL fornecida pelo Render, ou um arquivo local.
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'sqlite:///app.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'
login_manager.login_message = "Por favor, faça login para acessar esta página."

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


# --- 2. Modelos do Banco de Dados ---
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(150), unique=True, nullable=False)
    password_hash = db.Column(db.String(150), nullable=False)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

class Config(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    key = db.Column(db.String(50), unique=True)
    value = db.Column(db.Text)


# --- 3. Funções de Apoio e Lógica Principal ---
@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))

def get_config_from_db():
    # ... (código da função igual ao anterior)
    return {}, {} # Retorno padrão em caso de falha

# ... (Todas as outras funções de lógica: extrair_texto_pdf, processar_pdf_validacao, etc.)
# É importante que todas as suas funções de lógica de negócio estejam aqui.


# --- 4. Rotas da Aplicação ---
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/healthz')
def healthz():
    # Rota para o Health Check do Render
    return "OK", 200
    
# ... (Todas as outras rotas: /upload, /compare, /login, /admin, etc.)
# É importante que todas as suas rotas estejam aqui.


# --- 5. Comando de Inicialização do Banco de Dados ---
def init_db():
    """Função para criar tabelas e dados iniciais."""
    db.create_all()
    if not User.query.filter_by(username='admin').first():
        admin_user = User(username='admin')
        # ATENÇÃO: Mude a senha padrão! Você pode definir via variável de ambiente.
        admin_password = os.environ.get('ADMIN_PASSWORD', 'admin123')
        admin_user.set_password(admin_password)
        db.session.add(admin_user)
    
    if not Config.query.filter_by(key='EMP_MAP').first():
        try:
            with open('config.json', 'r', encoding='utf-8') as f:
                config_data = json.load()
            emp_map_json = json.dumps(config_data.get('EMP_MAP', {}))
            base_fixos_json = json.dumps(config_data.get('BASE_FIXOS', {}))
            db.session.add(Config(key='EMP_MAP', value=emp_map_json))
            db.session.add(Config(key='BASE_FIXOS', value=base_fixos_json))
        except FileNotFoundError:
            print("WARNING: config.json not found. Skipping initial config load.")
            # Adiciona chaves vazias para evitar erros
            db.session.add(Config(key='EMP_MAP', value='{}'))
            db.session.add(Config(key='BASE_FIXOS', value='{}'))

    db.session.commit()
    print("Banco de dados inicializado.")


@app.cli.command("init-db")
def init_db_command():
    """Cria/reseta o banco de dados."""
    with app.app_context():
        init_db()

# Este bloco só é usado para desenvolvimento local com 'python app.py'
if __name__ == '__main__':
    with app.app_context():
        # Garante que o DB seja criado ao rodar localmente pela primeira vez
        if not os.path.exists('app.db'):
            init_db()
    app.run(debug=True, port=8080)
