from __future__ import annotations

import hashlib
import hmac
import os
import sqlite3
import time
import uuid
from contextlib import closing
from pathlib import Path

from flask import Flask, jsonify, request


DATABASE_PATH = Path(os.environ.get("MULTIDOCSYNC_DB", "/var/lib/multidoc-sync/telemetry.sqlite3"))
HASH_SALT = os.environ.get("MULTIDOCSYNC_TELEMETRY_SALT", "")
ADMIN_TOKEN = os.environ.get("MULTIDOCSYNC_ADMIN_TOKEN", "")
ALLOWED_EVENTS = {"session_start", "heartbeat", "session_end"}

if not HASH_SALT:
    raise RuntimeError("MULTIDOCSYNC_TELEMETRY_SALT is required")
if not ADMIN_TOKEN:
    raise RuntimeError("MULTIDOCSYNC_ADMIN_TOKEN is required")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 4096


def connect() -> sqlite3.Connection:
    DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DATABASE_PATH, timeout=10)
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            received_at INTEGER NOT NULL,
            install_hash TEXT NOT NULL,
            session_id TEXT NOT NULL,
            event TEXT NOT NULL,
            version TEXT NOT NULL,
            platform TEXT NOT NULL,
            architecture TEXT NOT NULL,
            elapsed_seconds INTEGER NOT NULL
        )
        """
    )
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_events_received ON events(received_at)"
    )
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_events_install ON events(install_hash)"
    )
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_events_session ON events(session_id)"
    )
    return connection


def short_text(value, maximum: int) -> str:
    if not isinstance(value, str) or not value or len(value) > maximum:
        raise ValueError("invalid text field")
    return value


def valid_uuid(value) -> str:
    return str(uuid.UUID(short_text(value, 64)))


def install_hash(install_id: str) -> str:
    return hmac.new(HASH_SALT.encode(), install_id.encode(), hashlib.sha256).hexdigest()


@app.get("/healthz")
def healthz():
    return jsonify(status="ok")


@app.post("/v1/events")
def ingest_event():
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify(error="invalid_json"), 400
    try:
        event = short_text(payload.get("event"), 32)
        if event not in ALLOWED_EVENTS:
            raise ValueError("invalid event")
        raw_install_id = valid_uuid(payload.get("install_id"))
        session_id = valid_uuid(payload.get("session_id"))
        version = short_text(payload.get("version"), 32)
        platform_name = short_text(payload.get("platform"), 32)
        architecture = short_text(payload.get("architecture"), 32)
        elapsed = int(payload.get("elapsed_seconds", 0))
        if elapsed < 0 or elapsed > 7 * 24 * 3600:
            raise ValueError("invalid elapsed time")
    except (TypeError, ValueError, AttributeError):
        return jsonify(error="invalid_payload"), 400

    with closing(connect()) as connection:
        connection.execute(
            """
            INSERT INTO events (
                received_at, install_hash, session_id, event,
                version, platform, architecture, elapsed_seconds
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                int(time.time()),
                install_hash(raw_install_id),
                session_id,
                event,
                version,
                platform_name,
                architecture,
                elapsed,
            ),
        )
        connection.commit()
    return jsonify(status="accepted"), 202


def authorized() -> bool:
    supplied = request.headers.get("Authorization", "")
    expected = f"Bearer {ADMIN_TOKEN}"
    return hmac.compare_digest(supplied, expected)


@app.get("/v1/admin/summary")
def summary():
    if not authorized():
        return jsonify(error="unauthorized"), 401
    now = int(time.time())
    with closing(connect()) as connection:
        users = connection.execute("SELECT COUNT(DISTINCT install_hash) FROM events").fetchone()[0]
        sessions = connection.execute("SELECT COUNT(DISTINCT session_id) FROM events").fetchone()[0]
        usage_seconds = connection.execute(
            "SELECT COALESCE(SUM(max_elapsed), 0) FROM "
            "(SELECT MAX(elapsed_seconds) AS max_elapsed FROM events GROUP BY session_id)"
        ).fetchone()[0]
        active_7d = connection.execute(
            "SELECT COUNT(DISTINCT install_hash) FROM events WHERE received_at >= ?",
            (now - 7 * 86400,),
        ).fetchone()[0]
        active_30d = connection.execute(
            "SELECT COUNT(DISTINCT install_hash) FROM events WHERE received_at >= ?",
            (now - 30 * 86400,),
        ).fetchone()[0]
        versions = connection.execute(
            "SELECT version, COUNT(DISTINCT install_hash) FROM events "
            "GROUP BY version ORDER BY 2 DESC"
        ).fetchall()
    return jsonify(
        users=users,
        sessions=sessions,
        usage_seconds=usage_seconds,
        active_7d=active_7d,
        active_30d=active_30d,
        versions=[{"version": row[0], "users": row[1]} for row in versions],
    )
