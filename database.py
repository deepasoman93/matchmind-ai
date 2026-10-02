from __future__ import annotations

import hashlib
import hmac
import json
import os
import secrets
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd


DATABASE_PATH = Path(os.getenv("RESUME_APP_DB", "resume_shortlisting.db"))
PBKDF2_ITERATIONS = 600_000


def connect() -> sqlite3.Connection:
    connection = sqlite3.connect(DATABASE_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def initialize_database() -> None:
    with connect() as connection:
        connection.executescript("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE COLLATE NOCASE,
                password_hash TEXT NOT NULL,
                salt TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS analyses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                job_title TEXT NOT NULL,
                results_json TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
            );
        """)
        existing_columns = {
            row["name"] for row in connection.execute("PRAGMA table_info(users)").fetchall()
        }
        optional_columns = {
            "email": "TEXT",
            "display_name": "TEXT",
            "auth_provider": "TEXT NOT NULL DEFAULT 'local'",
            "oauth_subject": "TEXT",
        }
        for column, definition in optional_columns.items():
            if column not in existing_columns:
                connection.execute(f"ALTER TABLE users ADD COLUMN {column} {definition}")
        connection.execute(
            "CREATE UNIQUE INDEX IF NOT EXISTS idx_users_oauth_identity "
            "ON users(auth_provider, oauth_subject) WHERE oauth_subject IS NOT NULL"
        )
        analysis_columns = {
            row["name"] for row in connection.execute("PRAGMA table_info(analyses)").fetchall()
        }
        if "context_json" not in analysis_columns:
            connection.execute("ALTER TABLE analyses ADD COLUMN context_json TEXT")


def hash_password(password: str, salt: bytes) -> str:
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS).hex()


def create_user(username: str, password: str) -> tuple[bool, str]:
    username = username.strip()
    if len(username) < 3:
        return False, "Username must contain at least 3 characters."
    if len(password) < 8:
        return False, "Password must contain at least 8 characters."
    salt = secrets.token_bytes(32)
    try:
        with connect() as connection:
            connection.execute(
                "INSERT INTO users(username, password_hash, salt, created_at) VALUES (?, ?, ?, ?)",
                (username, hash_password(password, salt), salt.hex(), datetime.now(timezone.utc).isoformat()),
            )
        return True, "Account created. You can now log in."
    except sqlite3.IntegrityError:
        return False, "That username is already registered."


def authenticate(username: str, password: str) -> dict | None:
    with connect() as connection:
        row = connection.execute(
            "SELECT id, username, password_hash, salt, display_name, auth_provider "
            "FROM users WHERE username = ? AND auth_provider = 'local'", (username.strip(),)
        ).fetchone()
    if row is None:
        return None
    attempted = hash_password(password, bytes.fromhex(row["salt"]))
    if not hmac.compare_digest(attempted, row["password_hash"]):
        return None
    return {
        "id": row["id"],
        "username": row["username"],
        "display_name": row["display_name"] or row["username"],
        "auth_provider": row["auth_provider"],
    }


def get_or_create_oauth_user(
    provider: str,
    subject: str,
    email: str,
    display_name: str,
) -> dict:
    """Return one internal user for a verified external OIDC identity."""
    provider = provider.strip().lower()
    subject = subject.strip()
    if not provider or not subject:
        raise ValueError("The identity provider did not return a valid user identifier.")

    with connect() as connection:
        row = connection.execute(
            "SELECT id, username, email, display_name, auth_provider "
            "FROM users WHERE auth_provider = ? AND oauth_subject = ?",
            (provider, subject),
        ).fetchone()
        if row is None:
            internal_username = f"{provider}_{hashlib.sha256(subject.encode('utf-8')).hexdigest()[:24]}"
            cursor = connection.execute(
                "INSERT INTO users(username, password_hash, salt, created_at, email, display_name, auth_provider, oauth_subject) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    internal_username,
                    "EXTERNAL_OIDC_ACCOUNT",
                    "",
                    datetime.now(timezone.utc).isoformat(),
                    email.strip().lower() or None,
                    display_name.strip() or email.strip() or "Google user",
                    provider,
                    subject,
                ),
            )
            user_id = int(cursor.lastrowid)
            row = connection.execute(
                "SELECT id, username, email, display_name, auth_provider FROM users WHERE id = ?",
                (user_id,),
            ).fetchone()
        else:
            connection.execute(
                "UPDATE users SET email = ?, display_name = ? WHERE id = ?",
                (
                    email.strip().lower() or row["email"],
                    display_name.strip() or row["display_name"],
                    row["id"],
                ),
            )
            row = connection.execute(
                "SELECT id, username, email, display_name, auth_provider FROM users WHERE id = ?",
                (row["id"],),
            ).fetchone()
    return dict(row)


def save_analysis(
    user_id: int,
    job_title: str,
    results: pd.DataFrame,
    context: dict | None = None,
) -> int:
    with connect() as connection:
        cursor = connection.execute(
            "INSERT INTO analyses(user_id, job_title, results_json, context_json, created_at) "
            "VALUES (?, ?, ?, ?, ?)",
            (
                user_id,
                job_title.strip() or "Untitled role",
                results.to_json(orient="records"),
                json.dumps(context) if context else None,
                datetime.now(timezone.utc).isoformat(),
            ),
        )
        return int(cursor.lastrowid)


def list_analyses(user_id: int) -> list[dict]:
    with connect() as connection:
        rows = connection.execute(
            "SELECT id, job_title, created_at FROM analyses WHERE user_id = ? ORDER BY id DESC", (user_id,)
        ).fetchall()
    return [dict(row) for row in rows]


def load_analysis(user_id: int, analysis_id: int) -> pd.DataFrame | None:
    with connect() as connection:
        row = connection.execute(
            "SELECT results_json FROM analyses WHERE id = ? AND user_id = ?", (analysis_id, user_id)
        ).fetchone()
    return pd.DataFrame(json.loads(row["results_json"])) if row else None


def load_analysis_bundle(user_id: int, analysis_id: int) -> tuple[pd.DataFrame | None, dict | None]:
    """Load scores and the source text required for later tailoring."""
    with connect() as connection:
        row = connection.execute(
            "SELECT results_json, context_json FROM analyses WHERE id = ? AND user_id = ?",
            (analysis_id, user_id),
        ).fetchone()
    if row is None:
        return None, None
    results = pd.DataFrame(json.loads(row["results_json"]))
    context = json.loads(row["context_json"]) if row["context_json"] else None
    return results, context
