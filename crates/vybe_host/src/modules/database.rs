//! vybe:database — SQL database access (SQLite, PostgreSQL, MySQL).
//! Uses typed sqlx pools per driver — no AnyPool, so MySQL TINYINT/BIT work natively.

use std::collections::HashMap;
use std::sync::{Arc, Mutex, atomic::{AtomicU64, Ordering}};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};
use sqlx::{Column, Row};

static NEXT_CONN: AtomicU64 = AtomicU64::new(1);

/// Typed pool — one variant per supported driver.
enum DbPool {
    Sqlite(sqlx::SqlitePool),
    Mysql(sqlx::MySqlPool),
    Postgres(sqlx::PgPool),
}

struct DbConn {
    pool: DbPool,
    conn_str: String,
}

struct DbState {
    connections: HashMap<u64, DbConn>,
}

fn get_state() -> Arc<Mutex<DbState>> {
    use std::sync::OnceLock;
    static STATE: OnceLock<Arc<Mutex<DbState>>> = OnceLock::new();
    STATE.get_or_init(|| Arc::new(Mutex::new(DbState { connections: HashMap::new() }))).clone()
}

fn block_on<F: std::future::Future>(f: F) -> F::Output {
    use std::sync::OnceLock;
    static RT: OnceLock<tokio::runtime::Runtime> = OnceLock::new();
    let rt = RT.get_or_init(|| {
        tokio::runtime::Builder::new_current_thread()
            .enable_all()
            .build()
            .unwrap()
    });
    rt.block_on(f)
}

/// Connect to a URL and return a typed DbPool.
fn connect_pool(url: &str) -> Result<DbPool, String> {
    if url.starts_with("sqlite:") {
        let pool = block_on(
            sqlx::sqlite::SqlitePoolOptions::new()
                .max_connections(1)
                .connect(url)
        ).map_err(|e| e.to_string())?;
        Ok(DbPool::Sqlite(pool))
    } else if url.starts_with("mysql:") {
        let pool = block_on(
            sqlx::mysql::MySqlPoolOptions::new()
                .max_connections(5)
                .connect(url)
        ).map_err(|e| e.to_string())?;
        Ok(DbPool::Mysql(pool))
    } else if url.starts_with("postgres:") || url.starts_with("postgresql:") {
        let pool = block_on(
            sqlx::postgres::PgPoolOptions::new()
                .max_connections(5)
                .connect(url)
        ).map_err(|e| e.to_string())?;
        Ok(DbPool::Postgres(pool))
    } else {
        Err(format!("Unsupported connection URL scheme: {}", url))
    }
}

/// Close a typed pool.
fn close_pool(pool: DbPool) {
    match pool {
        DbPool::Sqlite(p) => block_on(p.close()),
        DbPool::Mysql(p) => block_on(p.close()),
        DbPool::Postgres(p) => block_on(p.close()),
    }
}

// ── Row → Value helpers ──────────────────────────────────────────────────────

fn val_from_str(raw: &str) -> Value {
    if let Ok(n) = raw.parse::<f64>() {
        Value::F64(n)
    } else if raw == "true" {
        Value::Bool(true)
    } else if raw == "false" {
        Value::Bool(false)
    } else if raw == "null" || raw.is_empty() {
        Value::Null
    } else {
        Value::String(Arc::from(raw))
    }
}

fn sqlite_row_to_value(row: &sqlx::sqlite::SqliteRow) -> Value {
    let mut obj = Object::new();
    for col in row.columns() {
        let name = col.name().to_string();
        let raw: String = row.try_get::<String, _>(col.ordinal())
            .or_else(|_| row.try_get::<i64, _>(col.ordinal()).map(|n| n.to_string()))
            .or_else(|_| row.try_get::<f64, _>(col.ordinal()).map(|n| n.to_string()))
            .or_else(|_| row.try_get::<bool, _>(col.ordinal()).map(|b| b.to_string()))
            .unwrap_or_else(|_| "null".to_string());
        obj.properties.insert(name, val_from_str(&raw));
    }
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn mysql_row_to_value(row: &sqlx::mysql::MySqlRow) -> Value {
    let mut obj = Object::new();
    for col in row.columns() {
        let name = col.name().to_string();
        let raw: String = row.try_get::<String, _>(col.ordinal())
            .or_else(|_| row.try_get::<i64, _>(col.ordinal()).map(|n| n.to_string()))
            .or_else(|_| row.try_get::<u64, _>(col.ordinal()).map(|n| n.to_string()))
            .or_else(|_| row.try_get::<f64, _>(col.ordinal()).map(|n| n.to_string()))
            .or_else(|_| row.try_get::<bool, _>(col.ordinal()).map(|b| b.to_string()))
            .or_else(|_| row.try_get::<i8, _>(col.ordinal()).map(|n| n.to_string()))
            .unwrap_or_else(|_| "null".to_string());
        obj.properties.insert(name, val_from_str(&raw));
    }
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn pg_row_to_value(row: &sqlx::postgres::PgRow) -> Value {
    let mut obj = Object::new();
    for col in row.columns() {
        let name = col.name().to_string();
        let raw: String = row.try_get::<String, _>(col.ordinal())
            .or_else(|_| row.try_get::<i64, _>(col.ordinal()).map(|n| n.to_string()))
            .or_else(|_| row.try_get::<f64, _>(col.ordinal()).map(|n| n.to_string()))
            .or_else(|_| row.try_get::<bool, _>(col.ordinal()).map(|b| b.to_string()))
            .unwrap_or_else(|_| "null".to_string());
        obj.properties.insert(name, val_from_str(&raw));
    }
    Value::Object(Arc::new(Mutex::new(obj)))
}

// ── Typed fetch helpers ──────────────────────────────────────────────────────

fn fetch_all_rows(pool: &DbPool, sql: &str) -> Result<Vec<Value>, sqlx::Error> {
    match pool {
        DbPool::Sqlite(p) => {
            let rows = block_on(sqlx::query(sql).fetch_all(p))?;
            Ok(rows.iter().map(sqlite_row_to_value).collect())
        }
        DbPool::Mysql(p) => {
            let rows = block_on(sqlx::query(sql).fetch_all(p))?;
            Ok(rows.iter().map(mysql_row_to_value).collect())
        }
        DbPool::Postgres(p) => {
            let rows = block_on(sqlx::query(sql).fetch_all(p))?;
            Ok(rows.iter().map(pg_row_to_value).collect())
        }
    }
}

fn fetch_optional_row(pool: &DbPool, sql: &str) -> Result<Option<Value>, sqlx::Error> {
    match pool {
        DbPool::Sqlite(p) => {
            Ok(block_on(sqlx::query(sql).fetch_optional(p))?.map(|r| sqlite_row_to_value(&r)))
        }
        DbPool::Mysql(p) => {
            Ok(block_on(sqlx::query(sql).fetch_optional(p))?.map(|r| mysql_row_to_value(&r)))
        }
        DbPool::Postgres(p) => {
            Ok(block_on(sqlx::query(sql).fetch_optional(p))?.map(|r| pg_row_to_value(&r)))
        }
    }
}

fn execute_sql(pool: &DbPool, sql: &str) -> Result<u64, sqlx::Error> {
    match pool {
        DbPool::Sqlite(p) => Ok(block_on(sqlx::query(sql).execute(p))?.rows_affected()),
        DbPool::Mysql(p) => Ok(block_on(sqlx::query(sql).execute(p))?.rows_affected()),
        DbPool::Postgres(p) => Ok(block_on(sqlx::query(sql).execute(p))?.rows_affected()),
    }
}

/// Fetch column names for a query by running it with LIMIT 0.
fn fetch_column_names(pool: &DbPool, sql: &str) -> Vec<String> {
    let limited = format!("{} LIMIT 0", sql.trim().trim_end_matches(';'));
    match pool {
        DbPool::Sqlite(p) => {
            if let Ok(rows) = block_on(sqlx::query(&limited).fetch_all(p)) {
                if let Some(r) = rows.first() {
                    return r.columns().iter().map(|c| c.name().to_string()).collect();
                }
            }
            vec![]
        }
        DbPool::Mysql(p) => {
            if let Ok(rows) = block_on(sqlx::query(&limited).fetch_all(p)) {
                if let Some(r) = rows.first() {
                    return r.columns().iter().map(|c| c.name().to_string()).collect();
                }
            }
            vec![]
        }
        DbPool::Postgres(p) => {
            if let Ok(rows) = block_on(sqlx::query(&limited).fetch_all(p)) {
                if let Some(r) = rows.first() {
                    return r.columns().iter().map(|c| c.name().to_string()).collect();
                }
            }
            vec![]
        }
    }
}

// ── Scalar extraction ────────────────────────────────────────────────────────

fn scalar_from_value(v: Value) -> Value {
    // v is already a row object — grab the first property
    if let Value::Object(obj) = v {
        let o = obj.lock().unwrap();
        if let Some(first) = o.properties.values().next() {
            return first.clone();
        }
    }
    Value::Null
}

// ── register ─────────────────────────────────────────────────────────────────

pub fn register(vm: &mut VM) {
    vm.register_host_fn("vybe:database", "connect", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let raw = s(args, 0);
        if raw.is_empty() || raw == "null" {
            let mut obj = Object::new();
            obj.properties.insert("__type".into(), Value::String(Arc::from("SqlConnection")));
            obj.properties.insert("__conn_id".into(), Value::F64(0.0));
            obj.properties.insert("state".into(), Value::String(Arc::from("Closed")));
            return Value::Object(Arc::new(Mutex::new(obj)));
        }
        let conn_str = normalize_conn_str_full(&raw);
        match connect_pool(&conn_str) {
            Ok(pool) => {
                let id = NEXT_CONN.fetch_add(1, Ordering::Relaxed);
                get_state().lock().unwrap().connections.insert(id, DbConn { pool, conn_str: conn_str.clone() });
                let mut obj = Object::new();
                obj.properties.insert("__type".into(), Value::String(Arc::from("SqlConnection")));
                obj.properties.insert("__conn_id".into(), Value::F64(id as f64));
                obj.properties.insert("connectionstring".into(), Value::String(Arc::from(conn_str.as_str())));
                obj.properties.insert("state".into(), Value::String(Arc::from("Open")));
                Value::Object(Arc::new(Mutex::new(obj)))
            }
            Err(e) => { eprintln!("db.connect error: {}", e); Value::Null }
        }
    }));

    vm.register_host_fn("vybe:database", "query", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let conn_id = get_conn_id(args);
        let sql = get_sql(args);
        let params = extract_params(args, 2);
        let sql = substitute_params(&sql, &params);

        let state = get_state();
        let guard = state.lock().unwrap();
        let conn = match guard.connections.get(&conn_id) {
            Some(c) => c,
            None => return Value::Object(Arc::new(Mutex::new(Object::new_array(vec![])))),
        };
        match fetch_all_rows(&conn.pool, &sql) {
            Ok(rows) => Value::Object(Arc::new(Mutex::new(Object::new_array(rows)))),
            Err(e) => {
                eprintln!("db.query error: {}", e);
                Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
            }
        }
    }));

    vm.register_host_fn("vybe:database", "execute", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let conn_id = get_conn_id(args);
        let sql = get_sql(args);
        let params = extract_params(args, 2);
        let sql = substitute_params(&sql, &params);

        let state = get_state();
        let guard = state.lock().unwrap();
        let conn = match guard.connections.get(&conn_id) {
            Some(c) => c,
            None => return Value::F64(-1.0),
        };
        match execute_sql(&conn.pool, &sql) {
            Ok(n) => Value::F64(n as f64),
            Err(e) => { eprintln!("db.execute error: {}", e); Value::F64(-1.0) }
        }
    }));

    vm.register_host_fn("vybe:database", "scalar", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let conn_id = get_conn_id(args);
        let sql = get_sql(args);
        let params = extract_params(args, 2);
        let sql = substitute_params(&sql, &params);

        let state = get_state();
        let guard = state.lock().unwrap();
        let conn = match guard.connections.get(&conn_id) {
            Some(c) => c,
            None => return Value::Null,
        };
        match fetch_optional_row(&conn.pool, &sql) {
            Ok(Some(row)) => scalar_from_value(row),
            Ok(None) => Value::Null,
            Err(e) => { eprintln!("db.scalar error: {}", e); Value::Null }
        }
    }));

    vm.register_host_fn("vybe:database", "open", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let raw = {
                let o = obj.lock().unwrap();
                o.properties.get("connectionstring").map(|v| format!("{}", v)).unwrap_or_default()
            };
            if raw.is_empty() { return Value::Null; }
            let conn_str = normalize_conn_str_full(&raw);
            match connect_pool(&conn_str) {
                Ok(pool) => {
                    let id = NEXT_CONN.fetch_add(1, Ordering::Relaxed);
                    get_state().lock().unwrap().connections.insert(id, DbConn { pool, conn_str });
                    let mut o = obj.lock().unwrap();
                    o.properties.insert("__conn_id".into(), Value::F64(id as f64));
                    o.properties.insert("state".into(), Value::String(Arc::from("Open")));
                }
                Err(e) => eprintln!("db.open error: {}", e),
            }
        }
        Value::Null
    }));

    vm.register_host_fn("vybe:database", "createCommand", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let conn_id = get_conn_id(args);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("SqlCommand")));
        obj.properties.insert("__conn_id".into(), Value::F64(conn_id as f64));
        obj.properties.insert("commandtext".into(), Value::String(Arc::from("")));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    vm.register_host_fn("vybe:database", "close", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let conn_id = get_conn_id(args);
        let state = get_state();
        let mut guard = state.lock().unwrap();
        if let Some(conn) = guard.connections.remove(&conn_id) {
            close_pool(conn.pool);
        }
        Value::Null
    }));

    vm.register_host_fn("vybe:database", "tables", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let conn_id = get_conn_id(args);
        let state = get_state();
        let guard = state.lock().unwrap();
        let conn = match guard.connections.get(&conn_id) {
            Some(c) => c,
            None => return Value::Object(Arc::new(Mutex::new(Object::new_array(vec![])))),
        };
        let sql = tables_sql(&conn.conn_str);
        match fetch_all_rows(&conn.pool, sql) {
            Ok(rows) => {
                let names: Vec<Value> = rows.into_iter().filter_map(|v| {
                    if let Value::Object(o) = v {
                        o.lock().unwrap().properties.values().next().cloned()
                    } else { None }
                }).collect();
                Value::Object(Arc::new(Mutex::new(Object::new_array(names))))
            }
            Err(_) => Value::Object(Arc::new(Mutex::new(Object::new_array(vec![])))),
        }
    }));

    vm.register_host_fn("vybe:database", "columns", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let conn_id = get_conn_id(args);
        let table = s(args, 1);
        let state = get_state();
        let guard = state.lock().unwrap();
        let conn = match guard.connections.get(&conn_id) {
            Some(c) => c,
            None => return Value::Object(Arc::new(Mutex::new(Object::new_array(vec![])))),
        };
        let sql = columns_sql(&conn.conn_str, &table);
        match fetch_all_rows(&conn.pool, &sql) {
            Ok(rows) => {
                let col_idx = if conn.conn_str.starts_with("sqlite:") { 1 } else { 0 };
                let names: Vec<Value> = rows.into_iter().filter_map(|v| {
                    if let Value::Object(o) = v {
                        o.lock().unwrap().properties.values().nth(col_idx).cloned()
                    } else { None }
                }).collect();
                Value::Object(Arc::new(Mutex::new(Object::new_array(names))))
            }
            Err(_) => Value::Object(Arc::new(Mutex::new(Object::new_array(vec![])))),
        }
    }));

    vm.register_host_fn("vybe:database", "beginTransaction", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let conn_id = get_conn_id(args);
        let state = get_state();
        let guard = state.lock().unwrap();
        let conn = match guard.connections.get(&conn_id) { Some(c) => c, None => return Value::Bool(false) };
        match execute_sql(&conn.pool, "BEGIN") {
            Ok(_) => Value::Bool(true),
            Err(_) => Value::Bool(false),
        }
    }));

    vm.register_host_fn("vybe:database", "commit", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let conn_id = get_conn_id(args);
        let state = get_state();
        let guard = state.lock().unwrap();
        let conn = match guard.connections.get(&conn_id) { Some(c) => c, None => return Value::Bool(false) };
        match execute_sql(&conn.pool, "COMMIT") {
            Ok(_) => Value::Bool(true),
            Err(_) => Value::Bool(false),
        }
    }));

    vm.register_host_fn("vybe:database", "rollback", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let conn_id = get_conn_id(args);
        let state = get_state();
        let guard = state.lock().unwrap();
        let conn = match guard.connections.get(&conn_id) { Some(c) => c, None => return Value::Bool(false) };
        match execute_sql(&conn.pool, "ROLLBACK") {
            Ok(_) => Value::Bool(true),
            Err(_) => Value::Bool(false),
        }
    }));
}

// ── Small helpers ─────────────────────────────────────────────────────────────

fn s(args: &[Value], idx: usize) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_default()
}

/// Get SQL from either a SqlCommand object's commandtext or a plain string arg.
fn get_sql(args: &[Value]) -> String {
    if let Some(Value::Object(obj)) = args.first() {
        let o = obj.lock().unwrap();
        let ct = o.properties.get("commandtext").map(|v| format!("{}", v)).unwrap_or_default();
        if !ct.is_empty() { return ct; }
    }
    s(args, 1)
}

fn get_conn_id(args: &[Value]) -> u64 {
    match args.first() {
        Some(Value::F64(n)) => *n as u64,
        Some(Value::Object(obj)) => {
            let o = obj.lock().unwrap();
            o.properties.get("__conn_id").map(|v| v.as_f64() as u64).unwrap_or(0)
        }
        _ => 0,
    }
}

fn extract_params(args: &[Value], idx: usize) -> Vec<String> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(ref elems) = o.kind {
            return elems.iter().map(|v| format!("{}", v)).collect();
        }
    }
    vec![]
}

fn substitute_params(sql: &str, params: &[String]) -> String {
    let mut result = String::new();
    let mut idx = 0;
    for ch in sql.chars() {
        if ch == '?' && idx < params.len() {
            let val = &params[idx];
            if val.parse::<f64>().is_ok() || val == "null" || val == "true" || val == "false" {
                result.push_str(val);
            } else {
                result.push('\'');
                result.push_str(&val.replace('\'', "''"));
                result.push('\'');
            }
            idx += 1;
        } else {
            result.push(ch);
        }
    }
    result
}

fn tables_sql(conn_str: &str) -> &'static str {
    if conn_str.starts_with("sqlite:") {
        "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
    } else if conn_str.starts_with("mysql:") {
        "SELECT table_name as name FROM information_schema.tables WHERE table_schema = DATABASE() ORDER BY table_name"
    } else {
        "SELECT table_name as name FROM information_schema.tables WHERE table_schema='public' ORDER BY table_name"
    }
}

fn columns_sql(conn_str: &str, table: &str) -> String {
    if conn_str.starts_with("sqlite:") {
        format!("PRAGMA table_info({})", table)
    } else {
        format!("SELECT column_name as name FROM information_schema.columns WHERE table_name='{}' ORDER BY ordinal_position", table)
    }
}

// ── Design-time utilities ─────────────────────────────────────────────────────

/// Test a connection string and return a list of table names.
pub fn test_connection_and_list_tables(conn_str: &str) -> Result<Vec<String>, String> {
    let url = normalize_conn_str_full(conn_str);
    let pool = connect_pool(&url)?;
    let sql = tables_sql(&url);
    let rows = fetch_all_rows(&pool, sql).map_err(|e| format!("Query failed: {}", e))?;
    let tables: Vec<String> = rows.into_iter().filter_map(|v| {
        if let Value::Object(o) = v {
            if let Some(Value::String(s)) = o.lock().unwrap().properties.values().next() {
                let name = s.to_string();
                if !name.starts_with("sqlite_") && !name.starts_with("_sqlx") {
                    return Some(name);
                }
            }
        }
        None
    }).collect();
    close_pool(pool);
    Ok(tables)
}

/// Fetch column names for a SELECT query (design-time schema inspection).
pub fn fetch_columns_for_query(conn_str: &str, select_cmd: &str) -> Result<Vec<String>, String> {
    if conn_str.is_empty() || select_cmd.is_empty() {
        return Ok(Vec::new());
    }
    let url = normalize_conn_str_full(conn_str);
    let pool = connect_pool(&url)?;
    let cols = fetch_column_names(&pool, select_cmd);
    close_pool(pool);
    Ok(cols)
}

/// Query rows from a database — used by the data binding system.
pub fn query_rows(conn_str: &str, sql: &str) -> Result<(Vec<String>, Vec<HashMap<String, String>>), String> {
    let url = normalize_conn_str_full(conn_str);
    let pool = connect_pool(&url).map_err(|e| format!("Connection failed: {}", e))?;

    let value_rows = fetch_all_rows(&pool, sql)
        .map_err(|e| format!("Query failed: {}", e))?;

    // Collect column names from first row
    let columns: Vec<String> = if let Some(Value::Object(first)) = value_rows.first() {
        first.lock().unwrap().properties.keys().cloned().collect()
    } else {
        fetch_column_names(&pool, sql)
    };

    let result: Vec<HashMap<String, String>> = value_rows.into_iter().filter_map(|v| {
        if let Value::Object(obj) = v {
            let map = obj.lock().unwrap().properties.iter().map(|(k, v)| {
                (k.clone(), format!("{}", v))
            }).collect();
            Some(map)
        } else { None }
    }).collect();

    close_pool(pool);
    Ok((columns, result))
}

/// Normalize any ADO.NET / VB-style connection string to a sqlx URL.
pub fn normalize_conn_str_full(raw: &str) -> String {
    let lower = raw.to_lowercase();

    if raw.starts_with("sqlite:") || raw.starts_with("postgres:") || raw.starts_with("postgresql:")
        || raw.starts_with("mysql:") || raw.contains("://") {
        return raw.to_string();
    }

    if lower == ":memory:" || lower.contains("data source=:memory:") {
        return "sqlite::memory:".to_string();
    }

    if lower.starts_with("data source=") || lower.starts_with("datasource=") {
        let path = raw.split('=').nth(1).unwrap_or("").split(';').next().unwrap_or("").trim().trim_matches('"');
        return format!("sqlite:{}?mode=rwc", path);
    }

    if lower.contains("server=") || lower.contains("host=") {
        let pairs: HashMap<String, String> = raw.split(';')
            .filter_map(|p| {
                let mut parts = p.splitn(2, '=');
                Some((parts.next()?.trim().to_lowercase(), parts.next()?.trim().to_string()))
            })
            .collect();
        let host = pairs.get("server").or(pairs.get("host")).map(|s| s.as_str()).unwrap_or("localhost");
        let port = pairs.get("port").map(|s| s.as_str()).unwrap_or("3306");
        let db   = pairs.get("database").or(pairs.get("db")).or(pairs.get("initial catalog")).map(|s| s.as_str()).unwrap_or("");
        let user = pairs.get("uid").or(pairs.get("user")).or(pairs.get("user id")).map(|s| s.as_str()).unwrap_or("root");
        let pass = pairs.get("pwd").or(pairs.get("password")).map(|s| s.as_str()).unwrap_or("");
        let driver = if port == "5432" { "postgres" } else { "mysql" };
        return format!("{}://{}:{}@{}:{}/{}", driver, user, pass, host, port, db);
    }

    format!("sqlite:{}?mode=rwc", raw)
}
