"""Tests para validaciones básicas de RUC y período."""
from src.validator import validar_periodo, validar_ruc


def test_ruc_valido_persona_natural():
    assert validar_ruc("1793205060001") is True


def test_ruc_invalido_corto():
    assert validar_ruc("123") is False


def test_ruc_invalido_no_termina_001():
    assert validar_ruc("1234567890123") is False


def test_ruc_invalido_con_letras():
    assert validar_ruc("179320506000A") is False


def test_ruc_vacio():
    assert validar_ruc("") is False


def test_periodo_valido():
    assert validar_periodo("2026-04") is True
    assert validar_periodo("2025-12") is True
    assert validar_periodo("2026-01") is True


def test_periodo_mes_invalido():
    assert validar_periodo("2026-13") is False
    assert validar_periodo("2026-00") is False


def test_periodo_formato_invalido():
    assert validar_periodo("2026/04") is False
    assert validar_periodo("abr-2026") is False
    assert validar_periodo("") is False
