"""Tests para autenticación PBKDF2."""
from pathlib import Path
from unittest.mock import patch

import pytest

from src import auth


@pytest.fixture
def users_file(tmp_path, monkeypatch):
    """Aisla el archivo de usuarios en tmp_path para no tocar el real."""
    fake_path = tmp_path / "users.json"
    monkeypatch.setattr(auth, "USERS_FILE", fake_path)
    yield fake_path


def test_hash_password_diferente_cada_vez():
    s1, h1 = auth.hash_password("clave_super_segura")
    s2, h2 = auth.hash_password("clave_super_segura")
    assert s1 != s2  # sal aleatoria
    assert h1 != h2  # por ende hash distinto


def test_verify_password_correcta(users_file):
    salt, h = auth.hash_password("mi_clave_2026")
    assert auth.verify_password("mi_clave_2026", salt, h) is True


def test_verify_password_incorrecta(users_file):
    salt, h = auth.hash_password("mi_clave_2026")
    assert auth.verify_password("otra_clave", salt, h) is False


def test_create_y_authenticate(users_file):
    auth.create_user("0930452024001", "tempPwd99", display_name="Admin", is_admin=True)
    user = auth.authenticate("0930452024001", "tempPwd99")
    assert user is not None
    assert user.username == "0930452024001"
    assert user.is_admin is True
    assert user.display_name == "Admin"


def test_authenticate_clave_mala(users_file):
    auth.create_user("u1", "real_pwd_123")
    assert auth.authenticate("u1", "wrong") is None


def test_authenticate_usuario_no_existe(users_file):
    assert auth.authenticate("inexistente", "x") is None


def test_must_change_password_flag(users_file):
    auth.create_user("u1", "temp", must_change_password=True)
    user = auth.authenticate("u1", "temp")
    assert user is not None
    assert user.must_change_password is True


def test_change_password_limpia_flag(users_file):
    auth.create_user("u1", "viejaTemp", must_change_password=True)
    ok = auth.change_password("u1", "nuevaPwdSegura1!")
    assert ok is True
    user = auth.authenticate("u1", "nuevaPwdSegura1!")
    assert user is not None
    assert user.must_change_password is False
    # la vieja ya no funciona
    assert auth.authenticate("u1", "viejaTemp") is None


def test_change_password_usuario_inexistente(users_file):
    assert auth.change_password("noexiste", "xxx") is False


def test_delete_user(users_file):
    auth.create_user("u1", "x")
    assert auth.delete_user("u1") is True
    assert auth.authenticate("u1", "x") is None
    assert auth.delete_user("u1") is False  # ya no existe


def test_list_users(users_file):
    auth.create_user("u1", "x", display_name="Uno")
    auth.create_user("u2", "y", display_name="Dos")
    users = list(auth.list_users())
    assert {u.username for u in users} == {"u1", "u2"}


def test_user_data_root_aislamiento(tmp_path):
    base = tmp_path / "data" / "clientes"
    p1 = auth.user_data_root("0930452024001", base)
    p2 = auth.user_data_root("0916635154", base)
    assert p1 != p2
    assert "0930452024001" in str(p1)
    assert "0916635154" in str(p2)


def test_archivo_users_no_contiene_password_en_plano(users_file):
    auth.create_user("0930452024001", "mi_clave_secreta_2026")
    contenido = users_file.read_text(encoding="utf-8")
    assert "mi_clave_secreta_2026" not in contenido  # nunca en plano
    assert "salt" in contenido
    assert "hash" in contenido
