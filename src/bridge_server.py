import json
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import quote_plus

from sqlalchemy import create_engine, text

try:
    import clickhouse_connect
except Exception:  # pragma: no cover - optional dependency at runtime
    clickhouse_connect = None


HOST = "127.0.0.1"
PORT = 8765
DEFAULT_DIALECT = "mysql"
DEFAULT_DRIVER = "pymysql"
DEFAULT_DB_TYPE = "mysql"


def json_response(handler, status_code, payload):
    body = json.dumps(payload, ensure_ascii=False, default=str).encode("utf-8")
    handler.send_response(status_code)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.send_header("Access-Control-Allow-Origin", "*")
    handler.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
    handler.send_header("Access-Control-Allow-Headers", "Content-Type")
    handler.end_headers()
    handler.wfile.write(body)


def build_conn_url(profile):
    dialect = DEFAULT_DIALECT
    driver = DEFAULT_DRIVER
    username = quote_plus(str(profile.get("username") or profile.get("USERNAME") or ""))
    password = quote_plus(str(profile.get("password") or profile.get("PASSWORD") or ""))
    host = profile.get("host") or profile.get("HOST") or "127.0.0.1"
    port = profile.get("port") or profile.get("PORT") or ""
    database = profile.get("database") or profile.get("DATABASE") or ""
    if port:
        return f"{dialect}+{driver}://{username}:{password}@{host}:{port}/{database}"
    return f"{dialect}+{driver}://{username}:{password}@{host}/{database}"


def normalize_db_type(profile):
    value = (
        profile.get("dbType")
        or profile.get("db_type")
        or profile.get("type")
        or profile.get("dialect")
        or profile.get("DIALECT")
        or DEFAULT_DB_TYPE
    )
    text_value = str(value).strip().lower()
    if text_value in {"clickhouse", "ch"} or "clickhouse" in text_value:
        return "clickhouse"
    return "mysql"


def create_clickhouse_client(profile):
    if clickhouse_connect is None:
        raise RuntimeError("未安装 clickhouse-connect，请先执行 pip install clickhouse-connect")

    host = str(profile.get("host") or profile.get("HOST") or "127.0.0.1")
    port_raw = profile.get("port") or profile.get("PORT") or "8123"
    username = str(profile.get("username") or profile.get("USERNAME") or "default")
    password = str(profile.get("password") or profile.get("PASSWORD") or "")
    database = str(profile.get("database") or profile.get("DATABASE") or "default")
    secure_raw = str(profile.get("secure") or profile.get("SECURE") or "").strip().lower()
    secure = secure_raw in {"1", "true", "yes", "on"}

    try:
        port = int(str(port_raw).strip()) if str(port_raw).strip() else 8123
    except ValueError as exc:
        raise ValueError(f"ClickHouse 端口非法: {port_raw}") from exc

    return clickhouse_connect.get_client(
        host=host,
        port=port,
        username=username,
        password=password,
        database=database,
        secure=secure,
    )


def escape_sql_value(value):
    return str(value).replace("'", "''")


def format_sql(sql_template, values):
    in_values = ", ".join(f"'{escape_sql_value(v)}'" for v in values)
    if "{}" in sql_template:
        return sql_template.replace("{}", in_values)
    if "{{values}}" in sql_template:
        return sql_template.replace("{{values}}", in_values)
    return sql_template


def test_connection(profile):
    db_type = normalize_db_type(profile)
    if db_type == "clickhouse":
        client = create_clickhouse_client(profile)
        try:
            result = client.query("SELECT 1 AS ok")
            rows = result.result_rows or []
            return bool(rows and int(rows[0][0]) == 1)
        finally:
            client.close()

    engine = create_engine(build_conn_url(profile), pool_pre_ping=True)
    try:
        with engine.connect() as conn:
            result = conn.execute(text("SELECT 1 AS ok"))
            row = result.fetchone()
            return bool(row and int(row[0]) == 1)
    finally:
        engine.dispose()


def execute_batch(profile, sql_template, values):
    query = format_sql(sql_template, values)
    db_type = normalize_db_type(profile)

    if db_type == "clickhouse":
        client = create_clickhouse_client(profile)
        try:
            result = client.query(query)
            keys = list(result.column_names or [])
            rows = result.result_rows or []
            return [dict(zip(keys, row)) for row in rows], keys
        finally:
            client.close()

    engine = create_engine(build_conn_url(profile), pool_pre_ping=True)
    try:
        with engine.connect() as conn:
            result = conn.execute(text(query))
            rows = result.fetchall()
            keys = list(result.keys())
            return [dict(zip(keys, row)) for row in rows], keys
    finally:
        engine.dispose()


class Handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        json_response(self, 200, {"ok": True})

    def do_POST(self):
        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(content_length) if content_length > 0 else b"{}"
            payload = json.loads(raw.decode("utf-8"))
        except Exception as exc:
            json_response(self, 400, {"ok": False, "message": f"JSON 解析失败: {exc}"})
            return

        try:
            if self.path == "/api/test-connection":
                profile = payload.get("profile") or {}
                ok = test_connection(profile)
                json_response(self, 200, {"ok": ok})
                return

            if self.path == "/api/query-batch":
                profile = payload.get("profile") or {}
                sql_template = str(payload.get("sql_template") or "")
                values = payload.get("values") or []
                if not isinstance(values, list):
                    raise ValueError("values 必须是数组")
                start = time.time()
                rows, columns = execute_batch(profile, sql_template, values)
                elapsed_ms = int((time.time() - start) * 1000)
                json_response(
                    self,
                    200,
                    {
                        "ok": True,
                        "rows": rows,
                        "columns": columns,
                        "row_count": len(rows),
                        "elapsed_ms": elapsed_ms,
                    },
                )
                return

            json_response(self, 404, {"ok": False, "message": f"未找到接口: {self.path}"})
        except Exception as exc:
            json_response(self, 500, {"ok": False, "message": str(exc)})

    def log_message(self, format, *args):
        return


def main():
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"Bridge API running: http://{HOST}:{PORT}")
    print("POST /api/test-connection")
    print("POST /api/query-batch")
    server.serve_forever()


if __name__ == "__main__":
    main()
