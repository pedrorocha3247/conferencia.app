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
from werkzeug.security import generate_password_hash, check_password_hash

# --- Configuração Inicial do App ---
app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'uma-chave-secreta-muito-forte-padrao')
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'sqlite:///app.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# --- Configuração do Banco de Dados e Login ---
db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message = "Por favor, faça login para acessar esta página."

# --- Modelos do Banco de Dados ---
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

# --- Funções de Carregamento e Lógica ---
@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))

def get_config_from_db():
    # ... (função igual à anterior)
    return EMP_MAP, BASE_FIXOS

# ... (Todas as outras funções de lógica: extrair_texto_pdf, processar_pdf_validacao, etc.)

# --- Rotas da Aplicação ---
@app.route('/')
def index():
    # ... (rota igual à anterior)
    return render_template('index.html')

# ... (Todas as outras rotas: /upload, /compare, /batch_upload, /logout, /admin, etc.)

# --- Setup Inicial e Comandos CLI ---
def setup_database():
    """Cria tabelas e dados iniciais se não existirem."""
    db.create_all()
    if not User.query.filter_by(username='admin').first():
        admin_user = User(username='admin')
        admin_user.set_password('admin123') # Mude essa senha em produção!
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

    db.session.commit()

# NOVO: Comando para inicializar o banco de dados
@app.cli.command("init-db")
def init_db_command():
    """Cria as tabelas do banco de dados e os dados iniciais."""
    setup_database()
    print("Banco de dados inicializado com sucesso.")

if __name__ == '__main__':
    with app.app_context():
        setup_database() # Mantido para facilidade de desenvolvimento local
    app.run(debug=True)
