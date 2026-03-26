//! vybe:database — SQL database access (SQLite, PostgreSQL, MySQL).
//! Wraps sqlx with synchronous blocking API for the VM.

use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;
use std::sync::{Arc, Mutex, atomic::{AtomicU64, Ordering}};
use vybe_bytecode::{VM, Value};
use vybe_bytecode::value::{Object, ObjectKind};
use sqlx::{Column, Row};

static NEXT_CONN: AtomicU64 = AtomicU64::new(1);

/// Shared database state — connections are stored globally.
struct DbState {
    connections: HashMap<u64, DbConn>,
}

struct DbConn {
    pool: sqlx::AnyPool,
    conn_str: String,
}

fn get_state() -> Arc<Mutex<DbState>> {
    use std::sync::OnceLock;
    static STATE: OnceLock<Arc<Mutex<DbState>>> = OnceLock::new();
    STATE.get_or_init(|| Arc::new(Mutex::new(DbState { connections: HashMap::new() }))).clone()
}

fn block_on<F: std::future::Future>(f: F) -> F::Output {
    tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap()
        .block_on(f)
}

pub fn register(vm: &mut VM) {
    // Install sqlx any drivers (SQLite, Postgres, MySQL)
    sqlx::any::install_default_drivers();
    // db.connect(connectionString) → connection id (number)
    vm.register_host_fn("vybe:database", "connect", Box::new(|args: &[Value]| {
        let conn_str = s(args, 0);
        match block_on(sqlx::any::AnyPoolOptions::new().max_connections(5).connect(&conn_str)) {
            Ok(pool) => {
                let id = NEXT_CONN.fetch_add(1, Ordering::Relaxed);
                let state = get_state();
                state.lock().unwrap().connections.insert(id, DbConn { pool, conn_str });
                Value::F64(id as f64)
            }
            Err(e) => {
                eprintln!("db.connect error: {}", e);
                Value::Null
            }
        }
    }));

    // db.query(conn, sql, params?) → array of row objects
    vm.register_host_fn("vybe:database", "query", Box::new(|args: &[Value]| {
        let conn_id = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
        let sql = s(args, 1);
        let params = extract_params(args, 2);

        let state = get_state();
        let guard = state.lock().unwrap();
        let conn = match guard.connections.get(&conn_id) {
            Some(c) => c,
            None => return Value::Object(Rc::new(RefCell::new(Object::new_array(vec![])))),
        };

        let sql_with_params = substitute_params(&sql, &params);
        match block_on(sqlx::query(&sql_with_params).fetch_all(&conn.pool)) {
            Ok(rows) => {
                let result: Vec<Value> = rows.iter().map(|row| {
                    let mut obj = Object::new();
                    for col in row.columns() {
                        let name = col.name().to_string();
                        let val: String = row.try_get::<String, _>(col.ordinal())
                            .or_else(|_| row.try_get::<i64, _>(col.ordinal()).map(|n| n.to_string()))
                            .or_else(|_| row.try_get::<f64, _>(col.ordinal()).map(|n| n.to_string()))
                            .or_else(|_| row.try_get::<bool, _>(col.ordinal()).map(|b| b.to_string()))
                            .unwrap_or_else(|_| "null".to_string());
                        let value = if let Ok(n) = val.parse::<f64>() {
                            Value::F64(n)
                        } else if val == "true" || val == "false" {
                            Value::Bool(val == "true")
                        } else if val == "null" {
                            Value::Null
                        } else {
                            Value::String(Rc::from(val.as_str()))
                        };
                        obj.properties.insert(name, value);
                    }
                    Value::Object(Rc::new(RefCell::new(obj)))
                }).collect();
                Value::Object(Rc::new(RefCell::new(Object::new_array(result))))
            }
            Err(e) => {
                eprintln!("db.query error: {}", e);
                Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
            }
        }
    }));

    // db.execute(conn, sql, params?) → rows affected (number)
    vm.register_host_fn("vybe:database", "execute", Box::new(|args: &[Value]| {
        let conn_id = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
        let sql = s(args, 1);
        let params = extract_params(args, 2);

        let state = get_state();
        let guard = state.lock().unwrap();
        let conn = match guard.connections.get(&conn_id) {
            Some(c) => c,
            None => return Value::F64(-1.0),
        };

        let sql_with_params = substitute_params(&sql, &params);
        match block_on(sqlx::query(&sql_with_params).execute(&conn.pool)) {
            Ok(result) => Value::F64(result.rows_affected() as f64),
            Err(e) => {
                eprintln!("db.execute error: {}", e);
                Value::F64(-1.0)
            }
        }
    }));

    // db.scalar(conn, sql, params?) → single value
    vm.register_host_fn("vybe:database", "scalar", Box::new(|args: &[Value]| {
        let conn_id = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
        let sql = s(args, 1);
        let params = extract_params(args, 2);

        let state = get_state();
        let guard = state.lock().unwrap();
        let conn = match guard.connections.get(&conn_id) {
            Some(c) => c,
            None => return Value::Null,
        };

        let sql_with_params = substitute_params(&sql, &params);
        match block_on(sqlx::query(&sql_with_params).fetch_optional(&conn.pool)) {
            Ok(Some(row)) => {
                if row.columns().is_empty() { return Value::Null; }
                let val: String = row.try_get::<String, _>(0)
                    .or_else(|_| row.try_get::<i64, _>(0).map(|n| n.to_string()))
                    .or_else(|_| row.try_get::<f64, _>(0).map(|n| n.to_string()))
                    .unwrap_or_else(|_| "null".to_string());
                if let Ok(n) = val.parse::<f64>() { Value::F64(n) }
                else if val == "null" { Value::Null }
                else { Value::String(Rc::from(val.as_str())) }
            }
            Ok(None) => Value::Null,
            Err(e) => {
                eprintln!("db.scalar error: {}", e);
                Value::Null
            }
        }
    }));

    // db.close(conn)
    vm.register_host_fn("vybe:database", "close", Box::new(|args: &[Value]| {
        let conn_id = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
        let state = get_state();
        let mut guard = state.lock().unwrap();
        if let Some(conn) = guard.connections.remove(&conn_id) {
            block_on(conn.pool.close());
        }
        Value::Null
    }));

    // db.tables(conn) → array of table names
    vm.register_host_fn("vybe:database", "tables", Box::new(|args: &[Value]| {
        let conn_id = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
        let state = get_state();
        let guard = state.lock().unwrap();
        let conn = match guard.connections.get(&conn_id) {
            Some(c) => c,
            None => return Value::Object(Rc::new(RefCell::new(Object::new_array(vec![])))),
        };

        // SQLite: query sqlite_master. For others: information_schema.
        let sql = if conn.conn_str.starts_with("sqlite:") {
            "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
        } else {
            "SELECT table_name as name FROM information_schema.tables WHERE table_schema='public' ORDER BY table_name"
        };

        match block_on(sqlx::query(sql).fetch_all(&conn.pool)) {
            Ok(rows) => {
                let names: Vec<Value> = rows.iter()
                    .filter_map(|r| r.try_get::<String, _>(0).ok())
                    .map(|n| Value::String(Rc::from(n.as_str())))
                    .collect();
                Value::Object(Rc::new(RefCell::new(Object::new_array(names))))
            }
            Err(_) => Value::Object(Rc::new(RefCell::new(Object::new_array(vec![])))),
        }
    }));

    // db.columns(conn, tableName) → array of column name strings
    vm.register_host_fn("vybe:database", "columns", Box::new(|args: &[Value]| {
        let conn_id = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
        let table = s(args, 1);
        let state = get_state();
        let guard = state.lock().unwrap();
        let conn = match guard.connections.get(&conn_id) {
            Some(c) => c,
            None => return Value::Object(Rc::new(RefCell::new(Object::new_array(vec![])))),
        };

        let sql = if conn.conn_str.starts_with("sqlite:") {
            format!("PRAGMA table_info({})", table)
        } else {
            format!("SELECT column_name as name FROM information_schema.columns WHERE table_name='{}' ORDER BY ordinal_position", table)
        };

        match block_on(sqlx::query(&sql).fetch_all(&conn.pool)) {
            Ok(rows) => {
                let col_idx = if conn.conn_str.starts_with("sqlite:") { 1 } else { 0 };
                let names: Vec<Value> = rows.iter()
                    .filter_map(|r| r.try_get::<String, _>(col_idx).ok())
                    .map(|n| Value::String(Rc::from(n.as_str())))
                    .collect();
                Value::Object(Rc::new(RefCell::new(Object::new_array(names))))
            }
            Err(_) => Value::Object(Rc::new(RefCell::new(Object::new_array(vec![])))),
        }
    }));

    // db.transaction(conn) → transaction id
    vm.register_host_fn("vybe:database", "beginTransaction", Box::new(|args: &[Value]| {
        let conn_id = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
        let state = get_state();
        let guard = state.lock().unwrap();
        let conn = match guard.connections.get(&conn_id) {
            Some(c) => c,
            None => return Value::Null,
        };
        match block_on(sqlx::query("BEGIN").execute(&conn.pool)) {
            Ok(_) => Value::Bool(true),
            Err(_) => Value::Bool(false),
        }
    }));

    // db.commit(conn)
    vm.register_host_fn("vybe:database", "commit", Box::new(|args: &[Value]| {
        let conn_id = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
        let state = get_state();
        let guard = state.lock().unwrap();
        let conn = match guard.connections.get(&conn_id) {
            Some(c) => c,
            None => return Value::Bool(false),
        };
        match block_on(sqlx::query("COMMIT").execute(&conn.pool)) {
            Ok(_) => Value::Bool(true),
            Err(_) => Value::Bool(false),
        }
    }));

    // db.rollback(conn)
    vm.register_host_fn("vybe:database", "rollback", Box::new(|args: &[Value]| {
        let conn_id = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
        let state = get_state();
        let guard = state.lock().unwrap();
        let conn = match guard.connections.get(&conn_id) {
            Some(c) => c,
            None => return Value::Bool(false),
        };
        match block_on(sqlx::query("ROLLBACK").execute(&conn.pool)) {
            Ok(_) => Value::Bool(true),
            Err(_) => Value::Bool(false),
        }
    }));
}

fn s(args: &[Value], idx: usize) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_default()
}

/// Extract params array from args[idx] if it's an array.
fn extract_params(args: &[Value], idx: usize) -> Vec<String> {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.borrow();
        if let ObjectKind::Array(ref elems) = o.kind {
            return elems.iter().map(|v| format!("{}", v)).collect();
        }
    }
    vec![]
}

/// Replace ? placeholders with actual values (escaped).
fn substitute_params(sql: &str, params: &[String]) -> String {
    let mut result = String::new();
    let mut param_idx = 0;
    for ch in sql.chars() {
        if ch == '?' && param_idx < params.len() {
            // Simple string escaping — replace ' with ''
            let val = &params[param_idx];
            if val.parse::<f64>().is_ok() || val == "null" || val == "true" || val == "false" {
                result.push_str(val);
            } else {
                result.push('\'');
                result.push_str(&val.replace('\'', "''"));
                result.push('\'');
            }
            param_idx += 1;
        } else {
            result.push(ch);
        }
    }
    result
}
