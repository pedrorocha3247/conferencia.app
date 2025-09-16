import pytest
from app import extrair_parcelas, limpar_rotulo

def test_limpar_rotulo():
    assert limpar_rotulo("Taxa de Conservação - 01/12") == "Taxa de Conservação"
    assert limpar_rotulo(" TAMA - Contribuição Social ") == "Contribuição Social"

def test_extrair_parcelas_simples():
    bloco_texto = """
    Lançamentos
    Taxa de Conservação     450,99
    Fundo de Transporte      15,00
    """
    parcelas = extrair_parcelas(bloco_texto)
    assert "Taxa de Conservação" in parcelas
    assert parcelas["Taxa de Conservação"] == 450.99
    assert parcelas["Fundo de Transporte"] == 15.00

def test_extrair_parcelas_com_lixo():
    bloco_texto = """
    Débitos do Mês
    Lançamentos         Vencimento
    Contrib. Social SLIM - 01/02      103,00     Débitos do Mês
    TOTAL A PAGAR                     103,00
    """
    parcelas = extrair_parcelas(bloco_texto)
    assert "Contrib. Social SLIM" in parcelas
    assert parcelas["Contrib. Social SLIM"] == 103.00
    assert "TOTAL A PAGAR" not in parcelas