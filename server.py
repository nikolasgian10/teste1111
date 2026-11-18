from __future__ import annotations

import hashlib
import hmac
import html
import ipaddress
import json
import logging
import os
import re
import secrets
import sqlite3
import socket
import sys
import threading
import time
import traceback
import unicodedata
import uuid
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from http import HTTPStatus
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse


ROOT = Path(__file__).resolve().parent
DEFAULT_DATA_DIR = ROOT / "data"
SESSION_SECONDS = 12 * 60 * 60
PASSWORD_ITERATIONS = 310_000
PRIVACY_POLICY_VERSION = "2026-11-06"
DEFAULT_TICKET_TYPES = {
    "instalacao": ("Solicitar instalação", "Instalar o Q-DAS em uma máquina"),
    "mascara_errada": ("Máscara errada", "Máscara de entrada/relatório incorreta"),
    "usuario_sem_acesso": ("Usuário específico sem acesso", "Um usuário não consegue acessar"),
    "criacao_usuario": ("Criação de usuário", "Novo usuário no Q-DAS"),
    "qdas_inativo": ("Q-DAS inativo", "Q-DAS está inativo ou não inicia"),
    "outros": ("Outros", "Qualquer outra solicitação"),
}
STATUSES = {"aberto", "em_analise", "em_andamento", "resolvido", "cancelado"}
NOTE_STATUSES = {"em_andamento", "resolvido", "cancelado"}
DETAIL_FIELDS = {
    "instalacao": {"computer_name_maquina", "operacao", "bp"},
    "mascara_errada": {"operacao"},
    "usuario_sem_acesso": {"usuario", "operacao"},
    "criacao_usuario": {"nome", "np", "operacao_departamento"},
    "qdas_inativo": set(),
    "outros": set(),
}
DETAIL_FIELD_LABELS = {
    "computer_name_maquina": "Computer name da máquina",
    "operacao": "Operação",
    "bp": "BP da máquina",
    "usuario": "Usuário (nome ou NP)",
    "nome": "Nome",
    "np": "NP",
    "operacao_departamento": "Operação / Departamento",
}
INITIAL_FACTORIES = ("MFAN", "MFSC", "MFCC", "MFBL", "MFBR", "MFBA", "MFPA")


class ApiError(Exception):
    def __init__(self, status: int, message: str):
        super().__init__(message)
        self.status = status


def data_dir() -> Path:
    return Path(os.environ.get("QDAS_DATA_DIR", DEFAULT_DATA_DIR)).resolve()


def database_path() -> Path:
    return data_dir() / "qdas.sqlite3"


def token_file_path() -> Path:
    return data_dir() / "first_admin_setup_token.txt"


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")


NOTIFICATION_RECIPIENTS = ("nikolas.souza@fts.org.br",)
DEFAULT_SYSTEM_URL = "http://localhost:8000"
_EMAIL_LOG_LOCK = threading.Lock()


def _log_email_notification(
    recipients: tuple[str, ...], success: bool, error: Exception | None = None
) -> None:
    timestamp = datetime.now().astimezone().isoformat(timespec="seconds")
    result = "sucesso" if success else "erro"
    message = (
        f"{timestamp} Envio de notificação por Outlook: {result}; "
        f"destinatários={', '.join(recipients)}"
    )
    if error is not None:
        message += f"; erro={error}"
    try:
        log_dir = data_dir() / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / "email_notifications.log"
        logger = logging.getLogger("qdas.email_notifications")
        logger.setLevel(logging.INFO)
        with _EMAIL_LOG_LOCK:
            if not any(
                getattr(handler, "baseFilename", None) == str(log_path.resolve())
                for handler in logger.handlers
            ):
                handler = logging.FileHandler(log_path, encoding="utf-8")
                handler.setFormatter(logging.Formatter("%(message)s"))
                logger.addHandler(handler)
        logger.error(message) if not success else logger.info(message)
    except Exception as logging_error:
        print(
            f"{message}; falha ao gravar arquivo de log: {logging_error}",
            file=sys.stderr,
        )


def enviar_notificacao_chamado(chamado: dict) -> None:
    """Envia a notificação usando o perfil do Outlook configurado no Windows."""
    recipients = NOTIFICATION_RECIPIENTS
    try:
        if not isinstance(chamado, dict):
            raise ValueError("Os dados do chamado para notificação são inválidos.")

        from win32com.client import Dispatch
        import pythoncom

        base_url = os.environ.get("QDAS_PUBLIC_URL", DEFAULT_SYSTEM_URL).strip().rstrip("/")
        parsed_url = urlparse(base_url)
        if (
            parsed_url.scheme not in {"http", "https"}
            or not parsed_url.netloc
            or parsed_url.username
            or parsed_url.password
        ):
            raise ValueError("Configure QDAS_PUBLIC_URL com o endereço HTTP(S) do sistema.")

        number = chamado.get("number", chamado.get("id", ""))
        title = " ".join(
            str(chamado.get("title") or chamado.get("ticket_type") or "Chamado").split()
        )
        requester = str(chamado.get("requester_name") or "Não informado")
        opened_at = str(chamado.get("created_at") or "Não informado")
        try:
            opened_at = datetime.fromisoformat(opened_at.replace("Z", "+00:00")).astimezone().strftime(
                "%d/%m/%Y %H:%M:%S"
            )
        except ValueError:
            pass
        description = str(chamado.get("description") or "Sem descrição.")
        system_link = f"{base_url}/"
        subject = f"[NOVO CHAMADO] #{number} - {title}"
        fields = (
            ("Número", f"#{number}"),
            ("Título", title),
            ("Solicitante", requester),
            ("Data de abertura", opened_at),
        )
        table_rows = "\n".join(
            "<tr>"
            '<td style="padding: 4px 12px 4px 0;"><strong>'
            f"{html.escape(label)}</strong></td>"
            f"<td>{html.escape(value)}</td></tr>"
            for label, value in fields
        )
        body = f"""\
<html>
  <body style="font-family: Arial, sans-serif; color: #222;">
    <h2>Novo chamado registrado</h2>
    <table style="border-collapse: collapse;">
      {table_rows}
    </table>
    <p><strong>Descrição</strong></p>
    <p style="white-space: pre-wrap;">{html.escape(description)}</p>
    <p><a href="{html.escape(system_link, quote=True)}">Abrir o sistema de chamados</a></p>
  </body>
</html>"""

        pythoncom.CoInitialize()
        try:
            message = Dispatch("Outlook.Application").CreateItem(0)
            message.To = ";".join(recipients)
            message.Subject = subject
            message.HTMLBody = body
            message.Send()
        finally:
            pythoncom.CoUninitialize()
        _log_email_notification(recipients, success=True)
    except Exception as error:
        _log_email_notification(recipients, success=False, error=error)


def connect() -> sqlite3.Connection:
    data_dir().mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(database_path(), timeout=20)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    connection.execute("PRAGMA busy_timeout = 20000")
    return connection


@contextmanager
def db_connection():
    connection = connect()
    try:
        with connection:
            yield connection
    finally:
        connection.close()


def hash_password(password: str, salt: bytes | None = None) -> tuple[str, str]:
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PASSWORD_ITERATIONS)
    return salt.hex(), digest.hex()


def verify_password(password: str, salt_hex: str, digest_hex: str) -> bool:
    _, candidate = hash_password(password, bytes.fromhex(salt_hex))
    return hmac.compare_digest(candidate, digest_hex)


def normalize_np(value: object) -> str:
    normalized = value.strip().lower() if isinstance(value, str) else ""
    return normalized if re.fullmatch(r"[a-z0-9._-]{2,40}", normalized) else ""


def resolve_client_computer_name(client_ip: str) -> str | None:
    try:
        if ipaddress.ip_address(client_ip).is_loopback:
            return None
        hostname = socket.gethostbyaddr(client_ip)[0].rstrip(".").split(".", 1)[0]
    except (OSError, ValueError):
        return None
    if not re.fullmatch(r"[a-zA-Z0-9][a-zA-Z0-9_-]{0,62}", hostname):
        return None
    if hostname.lower() == "localhost":
        return None
    return hostname.upper()


def initialize_database() -> None:
    data_dir().mkdir(parents=True, exist_ok=True)
    with db_connection() as db:
        db.execute("PRAGMA journal_mode = WAL")
        db.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                id TEXT PRIMARY KEY,
                np TEXT NOT NULL UNIQUE COLLATE NOCASE,
                name TEXT NOT NULL,
                password_salt TEXT NOT NULL,
                password_hash TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS user_roles (
                user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                role TEXT NOT NULL CHECK (role IN ('admin', 'support')),
                PRIMARY KEY (user_id, role)
            );
            CREATE TABLE IF NOT EXISTS sessions (
                token_hash TEXT PRIMARY KEY,
                user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                expires_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS tickets (
                id TEXT NOT NULL UNIQUE,
                number INTEGER PRIMARY KEY AUTOINCREMENT,
                computer_name TEXT NOT NULL,
                ticket_type TEXT NOT NULL,
                requester_name TEXT,
                description TEXT,
                details TEXT NOT NULL DEFAULT '{}',
                status TEXT NOT NULL DEFAULT 'aberto',
                assignee TEXT,
                support_notes TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                privacy_policy_version TEXT NOT NULL,
                privacy_accepted_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS ticket_types (
                id TEXT PRIMARY KEY,
                label TEXT NOT NULL,
                hint TEXT NOT NULL,
                fields TEXT NOT NULL DEFAULT '[]',
                description_required INTEGER NOT NULL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS ticket_status_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_id TEXT NOT NULL REFERENCES tickets(id) ON DELETE CASCADE,
                status TEXT NOT NULL,
                assignee TEXT,
                note TEXT,
                is_public INTEGER NOT NULL DEFAULT 0 CHECK (is_public IN (0, 1)),
                changed_at TEXT NOT NULL,
                changed_by TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS mini_fabricas (
                name TEXT PRIMARY KEY,
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS app_settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS login_attempts (
                ip TEXT PRIMARY KEY,
                window_started REAL NOT NULL,
                attempts INTEGER NOT NULL
            );
            CREATE INDEX IF NOT EXISTS tickets_computer_created
                ON tickets(computer_name, created_at DESC);
            CREATE INDEX IF NOT EXISTS tickets_created
                ON tickets(created_at DESC);
            """
        )
        ticket_columns = {
            column["name"] for column in db.execute("PRAGMA table_info(tickets)")
        }
        if "privacy_policy_version" not in ticket_columns:
            db.execute(
                "ALTER TABLE tickets ADD COLUMN privacy_policy_version TEXT NOT NULL DEFAULT ''"
            )
        if "privacy_accepted_at" not in ticket_columns:
            db.execute(
                "ALTER TABLE tickets ADD COLUMN privacy_accepted_at TEXT NOT NULL DEFAULT ''"
            )
        history_columns = {
            column["name"]
            for column in db.execute("PRAGMA table_info(ticket_status_history)")
        }
        if "is_public" not in history_columns:
            db.execute(
                "ALTER TABLE ticket_status_history "
                "ADD COLUMN is_public INTEGER NOT NULL DEFAULT 0"
            )
        for ticket_type, (label, hint) in DEFAULT_TICKET_TYPES.items():
            fields = [
                {"key": key, "label": DETAIL_FIELD_LABELS[key], "required": True}
                for key in sorted(DETAIL_FIELDS[ticket_type])
            ]
            db.execute(
                """INSERT OR IGNORE INTO ticket_types
                   (id, label, hint, fields, description_required) VALUES (?, ?, ?, ?, ?)""",
                (
                    ticket_type,
                    label,
                    hint,
                    json.dumps(fields, ensure_ascii=False),
                    int(ticket_type in {"mascara_errada", "qdas_inativo", "outros"}),
                ),
            )
        db.execute(
            """INSERT INTO ticket_status_history
               (ticket_id, status, assignee, note, changed_at, changed_by)
               SELECT t.id, t.status, t.assignee, t.support_notes, t.updated_at, 'Sistema'
               FROM tickets t
               WHERE t.support_notes IS NOT NULL AND trim(t.support_notes) != ''
                 AND NOT EXISTS (
                     SELECT 1 FROM ticket_status_history h WHERE h.ticket_id = t.id
                 )"""
        )
        db.executemany(
            "INSERT OR IGNORE INTO mini_fabricas(name, created_at) VALUES (?, ?)",
            [(name, now_iso()) for name in INITIAL_FACTORIES],
        )
        db.execute("DELETE FROM sessions WHERE expires_at <= ?", (now_iso(),))
        setup_pending = db.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0
        token_hash = db.execute(
            "SELECT value FROM app_settings WHERE key = 'first_admin_token_hash'"
        ).fetchone()

        if setup_pending and token_hash is None:
            configured_token = os.environ.get("FIRST_ADMIN_SETUP_TOKEN", "")
            if configured_token and not 32 <= len(configured_token) <= 256:
                raise RuntimeError("FIRST_ADMIN_SETUP_TOKEN precisa ter entre 32 e 256 caracteres.")
            token = configured_token or secrets.token_urlsafe(32)
            db.execute(
                "INSERT INTO app_settings(key, value) VALUES ('first_admin_token_hash', ?)",
                (hashlib.sha256(token.encode()).hexdigest(),),
            )
            if not configured_token:
                token_file_path().write_text(token, encoding="utf-8")
                print(
                    f"\nCódigo de configuração do primeiro administrador: {token}\n"
                    f"Ele também está salvo em: {token_file_path()}\n"
                    "Guarde-o em local seguro; o código será removido após o primeiro cadastro.\n",
                    flush=True,
                )
        elif setup_pending and token_hash is not None and not os.environ.get("FIRST_ADMIN_SETUP_TOKEN"):
            if token_file_path().exists():
                print(
                    f"Cadastro inicial pendente. Código: "
                    f"{token_file_path().read_text(encoding='utf-8').strip()}",
                    flush=True,
                )
            else:
                raise RuntimeError(
                    "Cadastro inicial pendente, mas o código não está disponível. "
                    "Defina FIRST_ADMIN_SETUP_TOKEN igual ao código original."
                )


def public_user(db: sqlite3.Connection, user_id: str) -> dict:
    row = db.execute("SELECT id, np, name FROM users WHERE id = ?", (user_id,)).fetchone()
    if row is None:
        raise ApiError(401, "Sessão inválida.")
    roles = {
        role["role"]
        for role in db.execute("SELECT role FROM user_roles WHERE user_id = ?", (user_id,))
    }
    return {
        "id": row["id"],
        "np": row["np"],
        "name": row["name"],
        "admin": "admin" in roles,
        "support": "support" in roles,
    }


def slugify(value: str) -> str:
    ascii_value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "_", ascii_value.lower()).strip("_")[:60]


def validate_ticket_type(payload: object) -> dict:
    if not isinstance(payload, dict):
        raise ApiError(400, "Configuração do tipo inválida.")
    label = payload.get("label")
    hint = payload.get("hint", "")
    fields = payload.get("fields")
    description_required = payload.get("description_required", False)
    if not isinstance(label, str) or not label.strip() or len(label.strip()) > 80:
        raise ApiError(400, "Informe um nome de tipo com até 80 caracteres.")
    if not isinstance(hint, str) or len(hint) > 160:
        raise ApiError(400, "Descrição do tipo inválida.")
    if not isinstance(description_required, bool):
        raise ApiError(400, "A obrigatoriedade da descrição é inválida.")
    if not isinstance(fields, list) or len(fields) > 9:
        raise ApiError(400, "Configure no máximo 9 campos por tipo.")
    clean_fields = []
    seen = set()
    for field in fields:
        if not isinstance(field, dict):
            raise ApiError(400, "Configuração de campo inválida.")
        key, field_label, required = (
            field.get("key"),
            field.get("label"),
            field.get("required"),
        )
        if (
            not isinstance(key, str)
            or not re.fullmatch(r"[a-z][a-z0-9_]{0,59}", key)
            or key in {"mini_fabrica", "description"}
            or key in seen
            or not isinstance(field_label, str)
            or not field_label.strip()
            or len(field_label.strip()) > 80
            or not isinstance(required, bool)
        ):
            raise ApiError(400, "Confira os nomes e a obrigatoriedade dos campos.")
        seen.add(key)
        clean_fields.append(
            {"key": key, "label": field_label.strip(), "required": required}
        )
    return {
        "label": label.strip(),
        "hint": hint.strip(),
        "fields": clean_fields,
        "description_required": description_required,
    }


def ticket_type_dict(row: sqlite3.Row) -> dict:
    item = dict(row)
    item["fields"] = json.loads(item["fields"])
    item["description_required"] = bool(item["description_required"])
    return item


def validate_ticket(payload: object, db: sqlite3.Connection) -> dict:
    if not isinstance(payload, dict):
        raise ApiError(400, "Dados do chamado inválidos.")
    computer = payload.get("computer_name")
    ticket_type = payload.get("ticket_type")
    requester = payload.get("requester_name", "")
    description = payload.get("description", "")
    details = payload.get("details", {})
    privacy_consent = payload.get("privacy_consent")
    if not isinstance(computer, str) or not computer.strip() or len(computer.strip()) > 64:
        raise ApiError(400, "Computer name inválido.")
    if not isinstance(ticket_type, str):
        raise ApiError(400, "Tipo de chamado inválido.")
    type_config = db.execute(
        "SELECT fields, description_required FROM ticket_types WHERE id = ?", (ticket_type,)
    ).fetchone()
    if type_config is None:
        raise ApiError(400, "Tipo de chamado inválido.")
    if not isinstance(requester, str) or len(requester) > 120:
        raise ApiError(400, "Nome do solicitante muito longo.")
    if not isinstance(description, str) or len(description) > 2000:
        raise ApiError(400, "Descrição muito longa.")
    if (
        not isinstance(privacy_consent, dict)
        or privacy_consent.get("accepted") is not True
        or privacy_consent.get("version") != PRIVACY_POLICY_VERSION
    ):
        raise ApiError(400, "Leia e aceite a Política de Privacidade antes de abrir o chamado.")
    if not isinstance(details, dict) or len(details) > 10:
        raise ApiError(400, "Detalhes do chamado inválidos.")
    if len(json.dumps(details, ensure_ascii=False).encode("utf-8")) > 8192:
        raise ApiError(400, "Detalhes do chamado inválidos.")
    if any(not isinstance(key, str) or not isinstance(value, str) or len(value) > 120
           for key, value in details.items()):
        raise ApiError(400, "Campo de detalhe inválido.")
    fields = json.loads(type_config["fields"])
    allowed_keys = {field["key"] for field in fields} | {"mini_fabrica"}
    if set(details) - allowed_keys:
        raise ApiError(400, "O chamado contém campos não configurados para este tipo.")
    if any(field["required"] and not details.get(field["key"], "").strip() for field in fields):
        raise ApiError(400, "Preencha os campos obrigatórios do chamado.")
    if type_config["description_required"] and not description.strip():
        raise ApiError(400, "Informe a descrição do chamado.")
    factory = details.get("mini_fabrica", "")
    if (
        not isinstance(factory, str)
        or not 1 <= len(factory) <= 40
        or factory != factory.strip().upper()
        or db.execute("SELECT 1 FROM mini_fabricas WHERE name = ?", (factory,)).fetchone() is None
    ):
        raise ApiError(400, "Mini fábrica inválida.")
    cleaned = {key: value.strip() for key, value in details.items()}
    return {
        "computer_name": computer.strip().upper(),
        "ticket_type": ticket_type,
        "requester_name": requester.strip() or None,
        "description": description.strip() or None,
        "details": cleaned,
        "privacy_policy_version": PRIVACY_POLICY_VERSION,
    }


def ticket_dict(row: sqlite3.Row) -> dict:
    item = dict(row)
    item["details"] = json.loads(item["details"])
    return item


class QdasHandler(BaseHTTPRequestHandler):
    server_version = "QdasPython/1.0"
    protocol_version = "HTTP/1.1"

    def log_message(self, format: str, *args: object) -> None:
        print(f"{self.address_string()} - {format % args}")

    def send_json(self, status: int, payload: object, headers: dict | None = None) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Content-Security-Policy", "frame-ancestors 'none'")
        if headers:
            for name, value in headers.items():
                self.send_header(name, value)
        self.end_headers()
        self.wfile.write(body)

    def read_json(self) -> object:
        if self.headers.get_content_type() != "application/json":
            raise ApiError(415, "Envie os dados no formato JSON.")
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError as exc:
            raise ApiError(400, "Tamanho da requisição inválido.") from exc
        if length < 1 or length > 32_768:
            raise ApiError(413, "Requisição vazia ou muito grande.")
        try:
            return json.loads(self.rfile.read(length))
        except (json.JSONDecodeError, UnicodeDecodeError) as exc:
            raise ApiError(400, "JSON inválido.") from exc

    def check_origin(self) -> None:
        origin = self.headers.get("Origin")
        if origin:
            scheme = os.environ.get("QDAS_PUBLIC_SCHEME", self.server.scheme)
            expected = f"{scheme}://{self.headers.get('Host', '')}"
            if origin.rstrip("/") != expected.rstrip("/"):
                raise ApiError(403, "Origem da requisição não permitida.")

    def session_token(self) -> str:
        cookie = SimpleCookie()
        try:
            cookie.load(self.headers.get("Cookie", ""))
        except Exception:
            return ""
        morsel = cookie.get("qdas_session")
        return morsel.value if morsel else ""

    def current_user(self, db: sqlite3.Connection, required_role: str | None = None) -> dict:
        token = self.session_token()
        if not token:
            raise ApiError(401, "Faça login para continuar.")
        digest = hashlib.sha256(token.encode()).hexdigest()
        row = db.execute(
            "SELECT user_id FROM sessions WHERE token_hash = ? AND expires_at > ?",
            (digest, now_iso()),
        ).fetchone()
        if row is None:
            raise ApiError(401, "Sessão expirada. Entre novamente.")
        user = public_user(db, row["user_id"])
        if required_role and not user[required_role]:
            raise ApiError(403, "Você não tem permissão para esta ação.")
        return user

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if path.startswith("/api/"):
            self.handle_api("GET", path)
        else:
            self.serve_static(path)

    def do_POST(self) -> None:
        self.handle_api("POST", urlparse(self.path).path)

    def do_PATCH(self) -> None:
        self.handle_api("PATCH", urlparse(self.path).path)

    def do_PUT(self) -> None:
        self.handle_api("PUT", urlparse(self.path).path)

    def do_DELETE(self) -> None:
        self.handle_api("DELETE", urlparse(self.path).path)

    def handle_api(self, method: str, path: str) -> None:
        try:
            if not path.startswith("/api/"):
                raise ApiError(404, "Rota não encontrada.")
            if method in {"POST", "PUT", "PATCH", "DELETE"}:
                self.check_origin()
            with db_connection() as db:
                response, headers = self.dispatch(db, method, path)
            if (
                method == "POST"
                and path == "/api/tickets"
                and isinstance(response, dict)
                and isinstance(response.get("number"), int)
            ):
                self.notify_ticket_created(response["number"])
            self.send_json(HTTPStatus.OK, response, headers)
        except ApiError as error:
            self.send_json(error.status, {"error": str(error)})
        except (sqlite3.Error, OSError):
            print("Erro interno ao acessar SQLite ou responder à requisição.", file=sys.stderr)
            self.send_json(500, {"error": "Erro interno do servidor. Consulte o log do servidor."})
        except Exception:
            traceback.print_exc()
            self.send_json(500, {"error": "Erro interno do servidor. Consulte o log do servidor."})

    def dispatch(
        self, db: sqlite3.Connection, method: str, path: str
    ) -> tuple[object, dict | None]:
        if method == "GET" and path == "/api/health":
            return {"ok": True}, None
        if method == "GET" and path == "/api/setup":
            return {"has_access": db.execute("SELECT 1 FROM users LIMIT 1").fetchone() is not None}, None
        if method == "GET" and path == "/api/client-computer":
            return {"computer_name": resolve_client_computer_name(self.client_address[0])}, None
        if method == "POST" and path == "/api/setup":
            self.bootstrap_admin(db, self.read_json())
            return {"ok": True}, None
        if method == "POST" and path == "/api/login":
            return self.login(db, self.read_json())
        if method == "GET" and path == "/api/me":
            try:
                return {"user": self.current_user(db)}, None
            except ApiError as error:
                if error.status == 401:
                    return {"user": None}, None
                raise
        if method == "POST" and path == "/api/logout":
            token = self.session_token()
            if token:
                db.execute(
                    "DELETE FROM sessions WHERE token_hash = ?",
                    (hashlib.sha256(token.encode()).hexdigest(),),
                )
            return {"ok": True}, {"Set-Cookie": self.cookie_header("", 0)}
        if method == "GET" and path == "/api/mini-fabricas":
            return [row["name"] for row in db.execute("SELECT name FROM mini_fabricas ORDER BY name")], None
        if method == "GET" and path == "/api/ticket-types":
            return [
                ticket_type_dict(row)
                for row in db.execute("SELECT * FROM ticket_types ORDER BY label COLLATE NOCASE")
            ], None
        if method == "POST" and path == "/api/ticket-types":
            self.current_user(db, "admin")
            return self.create_ticket_type(db, self.read_json()), None
        if method == "PUT" and path.startswith("/api/ticket-types/"):
            self.current_user(db, "admin")
            ticket_type = unquote(path.rsplit("/", 1)[-1])
            return self.update_ticket_type(db, ticket_type, self.read_json()), None
        if method == "DELETE" and path.startswith("/api/ticket-types/"):
            self.current_user(db, "admin")
            ticket_type = unquote(path.rsplit("/", 1)[-1])
            self.delete_ticket_type(db, ticket_type)
            return {"ok": True}, None
        if method == "POST" and path == "/api/mini-fabricas":
            self.current_user(db, "admin")
            body = self.read_json()
            name = body.get("name") if isinstance(body, dict) else None
            normalized = name.strip().upper() if isinstance(name, str) else ""
            if not normalized or len(normalized) > 40:
                raise ApiError(400, "Nome da mini fábrica inválido.")
            try:
                db.execute(
                    "INSERT INTO mini_fabricas(name, created_at) VALUES (?, ?)",
                    (normalized, now_iso()),
                )
            except sqlite3.IntegrityError as exc:
                raise ApiError(409, "Esta mini fábrica já está cadastrada.") from exc
            return {"ok": True}, None
        if method == "POST" and path == "/api/tickets":
            self.create_ticket(db, self.read_json())
            row = db.execute("SELECT number FROM tickets ORDER BY number DESC LIMIT 1").fetchone()
            return {"number": row["number"]}, None
        if method == "GET" and path.startswith("/api/computers/") and path.endswith("/tickets"):
            computer = unquote(path[len("/api/computers/") : -len("/tickets")]).strip().upper()
            if not computer or len(computer) > 64:
                raise ApiError(400, "Computer name inválido.")
            rows = db.execute(
                """SELECT id, number, ticket_type, status, created_at FROM tickets
                   WHERE computer_name = ?
                   ORDER BY (status NOT IN ('resolvido', 'cancelado')) DESC, created_at DESC
                   LIMIT 30""",
                (computer,),
            ).fetchall()
            tickets = [dict(row, public_notes=[]) for row in rows]
            if tickets:
                ticket_ids = [ticket["id"] for ticket in tickets]
                placeholders = ",".join("?" for _ in ticket_ids)
                notes = db.execute(
                    f"""SELECT ticket_id, note, changed_at FROM ticket_status_history
                        WHERE ticket_id IN ({placeholders})
                          AND is_public = 1 AND note IS NOT NULL AND trim(note) != ''
                        ORDER BY changed_at, id""",
                    ticket_ids,
                )
                notes_by_ticket: dict[str, list[dict]] = {}
                for entry in notes:
                    notes_by_ticket.setdefault(entry["ticket_id"], []).append(
                        {"note": entry["note"], "changed_at": entry["changed_at"]}
                    )
                for ticket in tickets:
                    ticket["public_notes"] = notes_by_ticket.get(ticket["id"], [])
                    del ticket["id"]
            return tickets, None
        if method == "GET" and path == "/api/tickets":
            self.current_user(db, "support")
            tickets = [ticket_dict(row) for row in db.execute(
                "SELECT * FROM tickets ORDER BY created_at DESC, number DESC"
            )]
            history = db.execute(
                """SELECT * FROM ticket_status_history
                   WHERE ticket_id IN (SELECT id FROM tickets)
                   ORDER BY changed_at, id"""
            )
            histories: dict[str, list[dict]] = {}
            for entry in history:
                item = dict(entry)
                item["is_public"] = bool(item["is_public"])
                histories.setdefault(entry["ticket_id"], []).append(item)
            for ticket in tickets:
                ticket["status_history"] = histories.get(ticket["id"], [])
            return tickets, None
        if method == "PATCH" and path.startswith("/api/tickets/"):
            self.update_ticket(db, path.rsplit("/", 1)[-1], self.read_json())
            return {"ok": True}, None
        if method == "GET" and path == "/api/users":
            self.current_user(db, "admin")
            rows = db.execute(
                """SELECT u.id, u.np, u.name,
                   EXISTS(SELECT 1 FROM user_roles r WHERE r.user_id=u.id AND r.role='admin') AS admin
                   FROM users u ORDER BY u.name COLLATE NOCASE"""
            )
            return [dict(row) for row in rows], None
        if method == "POST" and path == "/api/users":
            self.create_user(db, self.read_json())
            return {"ok": True}, None
        if method == "DELETE" and path.startswith("/api/users/"):
            user = self.current_user(db, "admin")
            target = path.rsplit("/", 1)[-1]
            if target == user["id"]:
                raise ApiError(400, "Você não pode remover seu próprio acesso.")
            cursor = db.execute("DELETE FROM users WHERE id = ?", (target,))
            if not cursor.rowcount:
                raise ApiError(404, "Usuário não encontrado.")
            return {"ok": True}, None
        raise ApiError(404, "Rota não encontrada.")

    def bootstrap_admin(self, db: sqlite3.Connection, body: object) -> None:
        if not isinstance(body, dict):
            raise ApiError(400, "Dados do usuário inválidos.")
        name, np, password, token = (
            body.get("name"), normalize_np(body.get("np")), body.get("password"), body.get("setupToken")
        )
        if not np:
            raise ApiError(400, "NP inválido (use letras e números).")
        if not isinstance(name, str) or not name.strip() or len(name.strip()) > 100:
            raise ApiError(400, "Informe o nome.")
        if not isinstance(password, str) or not 4 <= len(password) <= 128:
            raise ApiError(400, "A senha deve ter entre 4 e 128 caracteres.")
        if not isinstance(token, str) or len(token) > 256:
            raise ApiError(400, "Código de configuração inválido.")

        db.execute("BEGIN IMMEDIATE")
        if db.execute("SELECT 1 FROM users LIMIT 1").fetchone() is not None:
            raise ApiError(409, "O administrador já foi configurado.")
        configured = db.execute(
            "SELECT value FROM app_settings WHERE key = 'first_admin_token_hash'"
        ).fetchone()
        actual_hash = hashlib.sha256(token.encode()).hexdigest()
        if configured is None or not hmac.compare_digest(actual_hash, configured["value"]):
            raise ApiError(403, "Código de configuração incorreto.")
        self.insert_user(db, np, name.strip(), password, admin=True)
        db.execute("DELETE FROM app_settings WHERE key = 'first_admin_token_hash'")
        db.commit()
        try:
            token_file_path().unlink(missing_ok=True)
        except OSError:
            print("Aviso: não foi possível remover o arquivo do código inicial.", file=sys.stderr)

    def insert_user(
        self, db: sqlite3.Connection, np: str, name: str, password: str, admin: bool
    ) -> str:
        salt, digest = hash_password(password)
        user_id = str(uuid.uuid4())
        try:
            db.execute(
                """INSERT INTO users(id, np, name, password_salt, password_hash, created_at)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (user_id, np, name, salt, digest, now_iso()),
            )
        except sqlite3.IntegrityError as exc:
            raise ApiError(409, "Este NP já tem acesso.") from exc
        roles = ("admin", "support") if admin else ("support",)
        db.executemany(
            "INSERT INTO user_roles(user_id, role) VALUES (?, ?)",
            [(user_id, role) for role in roles],
        )
        return user_id

    def create_user(self, db: sqlite3.Connection, body: object) -> None:
        self.current_user(db, "admin")
        if not isinstance(body, dict):
            raise ApiError(400, "Dados do usuário inválidos.")
        np = normalize_np(body.get("np"))
        name = body.get("name")
        password = body.get("password")
        if not np:
            raise ApiError(400, "NP inválido (use letras e números).")
        if not isinstance(name, str) or not name.strip() or len(name.strip()) > 100:
            raise ApiError(400, "Informe o nome.")
        if not isinstance(password, str) or not 4 <= len(password) <= 128:
            raise ApiError(400, "A senha deve ter entre 4 e 128 caracteres.")
        admin = body.get("admin") is True
        self.insert_user(db, np, name.strip(), password, admin)

    def login(self, db: sqlite3.Connection, body: object) -> tuple[object, dict | None]:
        self.check_origin()
        ip = self.client_address[0]
        current_time = time.time()
        db.execute("DELETE FROM login_attempts WHERE window_started <= ?", (current_time - 600,))
        attempt = db.execute("SELECT * FROM login_attempts WHERE ip = ?", (ip,)).fetchone()
        if attempt and current_time - attempt["window_started"] < 600 and attempt["attempts"] >= 10:
            raise ApiError(429, "Muitas tentativas. Aguarde 10 minutos e tente novamente.")
        if not isinstance(body, dict):
            raise ApiError(400, "NP ou senha incorretos.")
        np, password = normalize_np(body.get("np")), body.get("password")
        row = db.execute(
            "SELECT id, password_salt, password_hash FROM users WHERE np = ? COLLATE NOCASE",
            (np,),
        ).fetchone() if np else None
        valid = (
            row is not None
            and isinstance(password, str)
            and verify_password(password, row["password_salt"], row["password_hash"])
        )
        if not valid:
            if attempt and current_time - attempt["window_started"] < 600:
                db.execute(
                    "UPDATE login_attempts SET attempts = attempts + 1 WHERE ip = ?", (ip,)
                )
            else:
                db.execute(
                    """INSERT INTO login_attempts(ip, window_started, attempts) VALUES (?, ?, 1)
                       ON CONFLICT(ip) DO UPDATE SET window_started=excluded.window_started, attempts=1""",
                    (ip, current_time),
                )
            db.commit()
            raise ApiError(401, "NP ou senha incorretos.")
        db.execute("DELETE FROM login_attempts WHERE ip = ?", (ip,))
        token = secrets.token_urlsafe(32)
        db.execute(
            "INSERT INTO sessions(token_hash, user_id, expires_at) VALUES (?, ?, ?)",
            (
                hashlib.sha256(token.encode()).hexdigest(),
                row["id"],
                (datetime.now(timezone.utc) + timedelta(seconds=SESSION_SECONDS))
                .isoformat(timespec="seconds")
                .replace("+00:00", "Z"),
            ),
        )
        return {"user": public_user(db, row["id"])}, {
            "Set-Cookie": self.cookie_header(token, SESSION_SECONDS)
        }

    def create_ticket(self, db: sqlite3.Connection, body: object) -> None:
        ticket = validate_ticket(body, db)
        db.execute(
            """INSERT INTO tickets(id, computer_name, ticket_type, requester_name, description,
               details, status, created_at, updated_at, privacy_policy_version, privacy_accepted_at)
               VALUES (?, ?, ?, ?, ?, ?, 'aberto', ?, ?, ?, ?)""",
            self.ticket_insert_values(ticket),
        )

    @staticmethod
    def notify_ticket_created(ticket_number: int) -> None:
        try:
            with db_connection() as db:
                row = db.execute(
                    """SELECT t.number, t.ticket_type, tt.label AS title,
                              t.requester_name, t.description, t.created_at
                       FROM tickets t
                       LEFT JOIN ticket_types tt ON tt.id = t.ticket_type
                       WHERE t.number = ?""",
                    (ticket_number,),
                ).fetchone()
            if row is None:
                raise RuntimeError(
                    f"O chamado #{ticket_number} não foi encontrado após a gravação."
                )
            chamado = dict(row)
            chamado["title"] = chamado["title"] or chamado["ticket_type"]
            enviar_notificacao_chamado(chamado)
        except Exception as error:
            _log_email_notification(
                NOTIFICATION_RECIPIENTS, success=False, error=error
            )

    def create_ticket_type(self, db: sqlite3.Connection, body: object) -> dict:
        config = validate_ticket_type(body)
        ticket_type = slugify(config["label"])
        if not ticket_type:
            raise ApiError(400, "Nome do tipo inválido.")
        if db.execute("SELECT 1 FROM ticket_types WHERE id = ?", (ticket_type,)).fetchone():
            raise ApiError(409, "Já existe um tipo de chamado com este nome.")
        db.execute(
            """INSERT INTO ticket_types(id, label, hint, fields, description_required)
               VALUES (?, ?, ?, ?, ?)""",
            (
                ticket_type,
                config["label"],
                config["hint"],
                json.dumps(config["fields"], ensure_ascii=False),
                int(config["description_required"]),
            ),
        )
        return {"id": ticket_type, **config}

    def update_ticket_type(
        self, db: sqlite3.Connection, ticket_type: str, body: object
    ) -> dict:
        config = validate_ticket_type(body)
        cursor = db.execute(
            """UPDATE ticket_types
               SET label=?, hint=?, fields=?, description_required=? WHERE id=?""",
            (
                config["label"],
                config["hint"],
                json.dumps(config["fields"], ensure_ascii=False),
                int(config["description_required"]),
                ticket_type,
            ),
        )
        if not cursor.rowcount:
            raise ApiError(404, "Tipo de chamado não encontrado.")
        return {"id": ticket_type, **config}

    @staticmethod
    def delete_ticket_type(db: sqlite3.Connection, ticket_type: str) -> None:
        if db.execute("SELECT 1 FROM ticket_types WHERE id = ?", (ticket_type,)).fetchone() is None:
            raise ApiError(404, "Tipo de chamado não encontrado.")
        if db.execute(
            "SELECT 1 FROM tickets WHERE ticket_type = ? LIMIT 1", (ticket_type,)
        ).fetchone():
            raise ApiError(
                409,
                "Este tipo já está associado a chamados e não pode ser excluído.",
            )
        db.execute("DELETE FROM ticket_types WHERE id = ?", (ticket_type,))

    @staticmethod
    def ticket_insert_values(ticket: dict) -> tuple:
        timestamp = now_iso()
        return (
            str(uuid.uuid4()),
            ticket["computer_name"],
            ticket["ticket_type"],
            ticket["requester_name"],
            ticket["description"],
            json.dumps(ticket["details"], ensure_ascii=False),
            timestamp,
            timestamp,
            ticket["privacy_policy_version"],
            timestamp,
        )

    def update_ticket(self, db: sqlite3.Connection, ticket_id: str, body: object) -> None:
        user = self.current_user(db, "support")
        try:
            uuid.UUID(ticket_id)
        except ValueError as exc:
            raise ApiError(400, "Identificador de chamado inválido.") from exc
        if not isinstance(body, dict):
            raise ApiError(400, "Dados do chamado inválidos.")
        status = body.get("status")
        assignee = body.get("assignee", "")
        notes = body.get("support_notes", "")
        note_is_public = body.get("note_is_public", False)
        if not isinstance(status, str) or status not in STATUSES:
            raise ApiError(400, "Status inválido.")
        if not isinstance(assignee, str) or len(assignee) > 100:
            raise ApiError(400, "Responsável inválido.")
        if not isinstance(notes, str) or len(notes) > 2000:
            raise ApiError(400, "Parecer inválido.")
        if not isinstance(note_is_public, bool):
            raise ApiError(400, "Visibilidade do parecer inválida.")
        current = db.execute(
            "SELECT status FROM tickets WHERE id = ?", (ticket_id,)
        ).fetchone()
        if current is None:
            raise ApiError(404, "Chamado não encontrado.")
        if status != current["status"] and status in NOTE_STATUSES and not notes.strip():
            raise ApiError(400, "Escreva um parecer para o usuário.")
        timestamp = now_iso()
        cleaned_note = notes.strip() or None
        db.execute(
            """UPDATE tickets SET status=?, assignee=?, support_notes=COALESCE(?, support_notes),
               updated_at=? WHERE id=?""",
            (status, assignee.strip() or None, cleaned_note, timestamp, ticket_id),
        )
        if status != current["status"] or cleaned_note:
            db.execute(
                """INSERT INTO ticket_status_history
                   (ticket_id, status, assignee, note, is_public, changed_at, changed_by)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (
                    ticket_id,
                    status,
                    assignee.strip() or None,
                    cleaned_note,
                    int(note_is_public and cleaned_note is not None),
                    timestamp,
                    user["name"],
                ),
            )

    def cookie_header(self, token: str, max_age: int) -> str:
        secure = "; Secure" if os.environ.get("QDAS_COOKIE_SECURE") == "1" else ""
        return (
            f"qdas_session={token}; Path=/; Max-Age={max_age}; HttpOnly; "
            f"SameSite=Strict{secure}"
        )

    def serve_static(self, path: str) -> None:
        static_root = ROOT / "dist"
        requested = (static_root / unquote(path.lstrip("/"))).resolve()
        if static_root.resolve() not in requested.parents and requested != static_root.resolve():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        target = requested if requested.is_file() else static_root / "index.html"
        if not target.is_file():
            self.send_error(
                HTTPStatus.SERVICE_UNAVAILABLE,
                "Interface não compilada. Execute npm run build antes de iniciar o servidor.",
            )
            return
        content = target.read_bytes()
        content_type = {
            ".html": "text/html; charset=utf-8",
            ".js": "text/javascript; charset=utf-8",
            ".css": "text/css; charset=utf-8",
            ".svg": "image/svg+xml",
            ".png": "image/png",
            ".ico": "image/x-icon",
            ".woff2": "font/woff2",
        }.get(target.suffix.lower(), "application/octet-stream")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(content)


class QdasServer(ThreadingHTTPServer):
    daemon_threads = True
    scheme = "http"


def main() -> None:
    initialize_database()
    host = os.environ.get("QDAS_HOST", "0.0.0.0")
    port = int(os.environ.get("QDAS_PORT", "8000"))
    server = QdasServer((host, port), QdasHandler)
    if host in {"0.0.0.0", "::"}:
        print("Q-DAS Chamados iniciado.", flush=True)
        print(f"Neste computador, abra: http://localhost:{port}", flush=True)
        print(
            f"Em outros computadores, use http://{socket.gethostname()}:{port} "
            "ou o IP de rede deste servidor.",
            flush=True,
        )
        print("Não digite 0.0.0.0 no navegador.", flush=True)
    else:
        print(f"Q-DAS Chamados disponível em http://{host}:{port}", flush=True)
    print(f"Banco SQLite: {database_path()}", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor encerrado.", flush=True)
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
