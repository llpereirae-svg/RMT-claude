"""Autenticación local con PBKDF2-SHA256 (stdlib, sin dependencias nuevas).

Las contraseñas NUNCA se guardan en plano: el archivo de usuarios contiene solo
sal y hash. Cada usuario tiene su propia carpeta de datos para aislamiento.
"""
from __future__ import annotations

import hashlib
import hmac
import json
import os
import secrets
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator

ROOT = Path(__file__).resolve().parent.parent
USERS_FILE = ROOT / ".streamlit" / "users.json"
PBKDF2_ITERATIONS = 200_000


@dataclass
class Usuario:
    username: str                   # RUC o cédula
    display_name: str               # Nombre amistoso para mostrar
    is_admin: bool                  # Reservado para roles futuros (sin uso actual)
    must_change_password: bool      # True hasta que cambie la contraseña provisional


# ─── Hashing ─────────────────────────────────────────────────────────────────
def hash_password(password: str, salt_hex: str | None = None) -> tuple[str, str]:
    """Devuelve (salt_hex, hash_hex). Si salt_hex es None, se genera uno nuevo."""
    salt = bytes.fromhex(salt_hex) if salt_hex else secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS)
    return salt.hex(), digest.hex()


def verify_password(password: str, salt_hex: str, hash_hex: str) -> bool:
    """Compara en tiempo constante para evitar ataques de timing."""
    try:
        salt = bytes.fromhex(salt_hex)
        expected = bytes.fromhex(hash_hex)
    except ValueError:
        return False
    actual = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS)
    return hmac.compare_digest(expected, actual)


# ─── Persistencia del archivo de usuarios ────────────────────────────────────
def _load_users_raw() -> dict[str, dict]:
    if not USERS_FILE.exists():
        return {}
    try:
        return json.loads(USERS_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def _save_users_raw(users: dict[str, dict]) -> None:
    USERS_FILE.parent.mkdir(parents=True, exist_ok=True)
    USERS_FILE.write_text(json.dumps(users, indent=2, ensure_ascii=False), encoding="utf-8")


def _to_usuario(username: str, info: dict) -> Usuario:
    return Usuario(
        username=username,
        display_name=info.get("display_name", username),
        is_admin=bool(info.get("is_admin", False)),
        must_change_password=bool(info.get("must_change_password", False)),
    )


def list_users() -> Iterator[Usuario]:
    for username, info in _load_users_raw().items():
        yield _to_usuario(username, info)


def get_user(username: str) -> Usuario | None:
    info = _load_users_raw().get(username)
    return _to_usuario(username, info) if info else None


def create_user(
    username: str,
    password: str,
    display_name: str = "",
    is_admin: bool = False,
    must_change_password: bool = False,
) -> None:
    """Crea o reemplaza un usuario. Idempotente."""
    salt_hex, hash_hex = hash_password(password)
    users = _load_users_raw()
    users[username] = {
        "salt": salt_hex,
        "hash": hash_hex,
        "display_name": display_name or username,
        "is_admin": is_admin,
        "must_change_password": must_change_password,
    }
    _save_users_raw(users)


def change_password(username: str, new_password: str) -> bool:
    """Actualiza la contraseña y limpia el flag must_change_password."""
    users = _load_users_raw()
    if username not in users:
        return False
    salt_hex, hash_hex = hash_password(new_password)
    users[username]["salt"] = salt_hex
    users[username]["hash"] = hash_hex
    users[username]["must_change_password"] = False
    _save_users_raw(users)
    return True


def delete_user(username: str) -> bool:
    users = _load_users_raw()
    if username not in users:
        return False
    del users[username]
    _save_users_raw(users)
    return True


def authenticate(username: str, password: str) -> Usuario | None:
    """Devuelve el Usuario si la contraseña coincide, None en caso contrario."""
    info = _load_users_raw().get(username.strip())
    if not info:
        return None
    if not verify_password(password, info.get("salt", ""), info.get("hash", "")):
        return None
    return _to_usuario(username.strip(), info)


# ─── Aislamiento de datos por usuario ────────────────────────────────────────
def user_data_root(username: str, base: Path) -> Path:
    """Cada usuario tiene su propia carpeta data/clientes_{username}/.
    El base original (data/clientes/) queda como referencia legacy."""
    safe = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in username)
    return base.parent / f"clientes_{safe}"
