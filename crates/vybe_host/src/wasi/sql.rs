//! `wasi:sql@0.2.0-draft` — SQL host implementation for Vybe.
//!
//! Implements the WIT surface from `proposals/wasi-sql/wit/`:
//!   - `wasi:sql/types`     → [static]connection.open, [static]statement.prepare,
//!                            [method]error.trace, [method]connection.close
//!   - `wasi:sql/readwrite` → query, exec
//!
//! Also registers a flat `wasi:sql` helper surface for older language adapters
//! that have not moved to the typed `wasi:sql/types` resource model yet.
//!
//! Architecture — one file per concern:
//!   state.rs    — global resource registry (connections, statements)
//!   driver.rs   — SqlDriver trait + open() URL dispatcher
//!   sqlite.rs   — rusqlite backend  (bundled libsqlite3, no system dep)
//!   postgres.rs — postgres  backend (sync, pure Rust; ?→$N rewrite)
//!   mysql.rs    — mysql     backend (sync, pure Rust; mysql2:// alias)

mod driver;
mod mysql;
mod postgres;
mod sqlite;
mod state;

use driver::{SqlDriver, open};
use state::{ConnEntry, StmtEntry, next_id, state};
use std::collections::HashMap;
use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{HostContext, VM, Value};

// ── Shared scalar parser (used by postgres.rs and mysql.rs) ──────────────────

pub(super) fn parse_scalar(s: &str) -> Value {
    if s.is_empty() || s.eq_ignore_ascii_case("null") {
        return Value::Null;
    }
    if s == "true" || s == "t" || s == "TRUE" {
        return Value::Bool(true);
    }
    if s == "false" || s == "f" || s == "FALSE" {
        return Value::Bool(false);
    }
    if let Ok(n) = s.parse::<f64>() {
        return Value::F64(n);
    }
    Value::String(Arc::from(s))
}

// ── Core dispatch — releases SqlState lock before the actual DB call ──────────

fn get_driver(conn_id: u64) -> Result<Arc<dyn SqlDriver>, String> {
    let s = state();
    let g = s.lock().unwrap();
    g.conns
        .get(&conn_id)
        .map(|e| Arc::clone(&e.driver))
        .ok_or_else(|| "Connection not found".to_string())
}

fn do_query(conn_id: u64, sql: &str, params: &[String]) -> Result<Vec<Value>, String> {
    get_driver(conn_id)?.query(sql, params)
}

fn do_exec(conn_id: u64, sql: &str, params: &[String]) -> Result<u64, String> {
    get_driver(conn_id)?.exec(sql, params)
}

fn do_query_columns(conn_id: u64, sql: &str, params: &[String]) -> Result<Vec<String>, String> {
    get_driver(conn_id)?.query_columns(sql, params)
}

fn driver_provider_name(url: &str) -> &'static str {
    if url.starts_with("sqlite:") {
        "sqlite"
    } else if url.starts_with("postgres:") || url.starts_with("postgresql:") {
        "postgres"
    } else if url.starts_with("mysql:") || url.starts_with("mysql2:") {
        "mysql"
    } else {
        "sql"
    }
}

fn driver_server_version(url: &str) -> &'static str {
    if url.starts_with("sqlite:") {
        "sqlite"
    } else if url.starts_with("postgres:") || url.starts_with("postgresql:") {
        "postgres"
    } else if url.starts_with("mysql:") || url.starts_with("mysql2:") {
        "mysql"
    } else {
        "sql"
    }
}

// ── Object constructors ───────────────────────────────────────────────────────

fn make_conn_obj(id: u64, url: &str) -> Value {
    let mut obj = Object::new();
    obj.properties.insert(
        "__wasi_kind".into(),
        Value::String(Arc::from("sql-connection")),
    );
    obj.properties
        .insert("__wasi_id".into(), Value::F64(id as f64));
    obj.properties
        .insert("__conn_id".into(), Value::F64(id as f64));
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("SqlConnection")));
    obj.properties
        .insert("connectionstring".into(), Value::String(Arc::from(url)));
    obj.properties
        .insert("provider".into(), string_value(driver_provider_name(url)));
    obj.properties.insert(
        "serverversion".into(),
        string_value(driver_server_version(url)),
    );
    obj.properties
        .insert("connectiontimeout".into(), Value::F64(30.0));
    obj.properties
        .insert("state".into(), Value::String(Arc::from("Open")));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_stmt_obj(id: u64) -> Value {
    let mut obj = Object::new();
    obj.properties.insert(
        "__wasi_kind".into(),
        Value::String(Arc::from("sql-statement")),
    );
    obj.properties
        .insert("__wasi_id".into(), Value::F64(id as f64));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_error_obj(msg: &str) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__wasi_kind".into(), Value::String(Arc::from("sql-error")));
    obj.properties
        .insert("message".into(), Value::String(Arc::from(msg)));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn rows_array(rows: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(rows))))
}

fn string_value(value: &str) -> Value {
    Value::String(Arc::from(value))
}

fn bool_prop(obj: &Arc<Mutex<Object>>, key: &str) -> bool {
    match obj.lock().unwrap().properties.get(key) {
        Some(Value::Bool(value)) => *value,
        Some(Value::F64(value)) => *value != 0.0,
        Some(Value::I32(value)) => *value != 0,
        Some(Value::String(value)) => !value.is_empty() && value.as_ref() != "false",
        Some(Value::Null) | None => false,
        Some(_) => true,
    }
}

fn f64_prop(obj: &Arc<Mutex<Object>>, key: &str) -> f64 {
    obj.lock()
        .unwrap()
        .properties
        .get(key)
        .map(Value::as_f64)
        .unwrap_or(0.0)
}

fn string_prop(obj: &Arc<Mutex<Object>>, key: &str) -> String {
    obj.lock()
        .unwrap()
        .properties
        .get(key)
        .map(|v| format!("{}", v))
        .unwrap_or_default()
}

fn array_prop(obj: &Arc<Mutex<Object>>, key: &str) -> Vec<Value> {
    let guard = obj.lock().unwrap();
    let Some(Value::Object(values)) = guard.properties.get(key) else {
        return vec![];
    };
    let values_guard = values.lock().unwrap();
    let ObjectKind::Array(ref elems) = values_guard.kind else {
        return vec![];
    };
    elems.clone()
}

fn set_prop(obj: &Arc<Mutex<Object>>, key: &str, value: Value) {
    obj.lock().unwrap().properties.insert(key.into(), value);
}

fn col_names_from_rows(rows: &[Value]) -> Vec<String> {
    let Some(Value::Object(first)) = rows.first() else {
        return vec![];
    };
    let guard = first.lock().unwrap();
    if let Some(Value::Object(names)) = guard.properties.get("__col_names") {
        let names_guard = names.lock().unwrap();
        if let ObjectKind::Array(ref elems) = names_guard.kind {
            return elems.iter().map(|value| format!("{}", value)).collect();
        }
    }
    guard
        .properties
        .keys()
        .filter(|key| !key.starts_with("__"))
        .cloned()
        .collect()
}

fn row_value_by_index(row: &Value, index: usize) -> Value {
    let Value::Object(obj) = row else {
        return Value::Null;
    };
    let col_names = {
        let guard = obj.lock().unwrap();
        if let Some(Value::Object(names)) = guard.properties.get("__col_names") {
            let names_guard = names.lock().unwrap();
            if let ObjectKind::Array(ref elems) = names_guard.kind {
                elems
                    .iter()
                    .map(|value| format!("{}", value))
                    .collect::<Vec<_>>()
            } else {
                vec![]
            }
        } else {
            guard
                .properties
                .keys()
                .filter(|key| key.as_str() != "__col_names")
                .cloned()
                .collect::<Vec<_>>()
        }
    };
    col_names
        .get(index)
        .map(|name| row_value_by_name(row, name))
        .unwrap_or(Value::Null)
}

fn row_value_by_name(row: &Value, name: &str) -> Value {
    let Value::Object(obj) = row else {
        return Value::Null;
    };
    let guard = obj.lock().unwrap();
    if let Some(value) = guard.properties.get(name) {
        return value.clone();
    }
    guard
        .properties
        .iter()
        .find(|(key, _)| key.eq_ignore_ascii_case(name))
        .map(|(_, value)| value.clone())
        .unwrap_or(Value::Null)
}

fn data_table_from_rows(name: &str, rows: &[Value], col_names: &[String]) -> Value {
    let mut table = Object::new();
    table
        .properties
        .insert("__type".into(), string_value("DataTable"));
    table
        .properties
        .insert("tablename".into(), string_value(name));
    table.properties.insert(
        "columns".into(),
        rows_array(col_names.iter().map(|col| string_value(col)).collect()),
    );
    table
        .properties
        .insert("rows".into(), rows_array(rows.to_vec()));
    Value::Object(Arc::new(Mutex::new(table)))
}

fn make_params_obj() -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), string_value("SqlParameterCollection"));
    obj.properties.insert("__items".into(), rows_array(vec![]));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_command_obj(type_name: &str, conn_id: u64, conn_string: &str) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), string_value(type_name));
    obj.properties
        .insert("__conn_id".into(), Value::F64(conn_id as f64));
    obj.properties
        .insert("commandtext".into(), string_value(""));
    obj.properties
        .insert("commandtimeout".into(), Value::F64(30.0));
    obj.properties.insert("commandtype".into(), Value::F64(1.0));
    obj.properties
        .insert("connectionstring".into(), string_value(conn_string));
    obj.properties
        .insert("parameters".into(), make_params_obj());
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_transaction_obj(conn_id: u64) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), string_value("SqlTransaction"));
    obj.properties
        .insert("__conn_id".into(), Value::F64(conn_id as f64));
    obj.properties.insert("isclosed".into(), Value::Bool(false));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_reader_obj_with_cols(rows: Vec<Value>, col_names: Vec<String>) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), string_value("SqlDataReader"));
    obj.properties
        .insert("__rows".into(), rows_array(rows.clone()));
    obj.properties.insert(
        "__col_names".into(),
        rows_array(col_names.iter().map(|col| string_value(col)).collect()),
    );
    obj.properties.insert("__pos".into(), Value::F64(-1.0));
    obj.properties
        .insert("hasrows".into(), Value::Bool(!rows.is_empty()));
    obj.properties
        .insert("fieldcount".into(), Value::F64(col_names.len() as f64));
    obj.properties.insert("isclosed".into(), Value::Bool(false));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_reader_obj(rows: Vec<Value>) -> Value {
    let col_names = col_names_from_rows(&rows);
    make_reader_obj_with_cols(rows, col_names)
}

fn make_data_adapter_obj(sql: &str, conn_id: u64, conn_string: &str) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), string_value("SqlDataAdapter"));
    obj.properties
        .insert("selectcommand".into(), string_value(sql));
    obj.properties
        .insert("__conn_id".into(), Value::F64(conn_id as f64));
    obj.properties
        .insert("connectionstring".into(), string_value(conn_string));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn current_reader_row(reader: &Arc<Mutex<Object>>) -> Option<Value> {
    let pos = f64_prop(reader, "__pos") as isize;
    let rows = array_prop(reader, "__rows");
    if pos < 0 || pos as usize >= rows.len() {
        None
    } else {
        Some(rows[pos as usize].clone())
    }
}

fn command_sql_and_params(command: &Arc<Mutex<Object>>) -> (String, Vec<String>) {
    let sql = string_prop(command, "commandtext");
    let Some(Value::Object(params)) = command
        .lock()
        .unwrap()
        .properties
        .get("parameters")
        .cloned()
    else {
        return (sql, vec![]);
    };
    let items = array_prop(&params, "__items");
    if items.is_empty() {
        return (sql, vec![]);
    }
    let mut values_by_name: HashMap<String, String> = HashMap::new();
    let mut ordered_values = Vec::new();
    for item in &items {
        let Value::Object(param) = item else {
            continue;
        };
        let name = string_prop(param, "name");
        let value = string_prop(param, "value");
        if !name.is_empty() {
            values_by_name.insert(name, value.clone());
        }
        ordered_values.push(value);
    }
    let mut rewritten = String::with_capacity(sql.len());
    let mut out_params = Vec::new();
    let chars: Vec<char> = sql.chars().collect();
    let mut index = 0usize;
    let mut in_str = false;
    while index < chars.len() {
        let ch = chars[index];
        if ch == '\'' {
            in_str = !in_str;
            rewritten.push(ch);
            index += 1;
            continue;
        }
        if !in_str && ch == '@' {
            let mut end = index + 1;
            while end < chars.len() && (chars[end].is_ascii_alphanumeric() || chars[end] == '_') {
                end += 1;
            }
            let name: String = chars[index..end].iter().collect();
            if let Some(value) = values_by_name.get(&name) {
                rewritten.push('?');
                out_params.push(value.clone());
                index = end;
                continue;
            }
        }
        rewritten.push(ch);
        index += 1;
    }
    if out_params.is_empty() {
        (sql, ordered_values)
    } else {
        (rewritten, out_params)
    }
}

fn create_command_from_connection(args: &[Value], type_name: &str) -> Value {
    let conn_id = get_conn_id(args);
    let conn_string = match args.first() {
        Some(Value::Object(obj)) => string_prop(obj, "connectionstring"),
        _ => String::new(),
    };
    make_command_obj(type_name, conn_id, &conn_string)
}

// ── Arg helpers ───────────────────────────────────────────────────────────────

fn arg_str(args: &[Value], i: usize) -> String {
    args.get(i).map(|v| format!("{}", v)).unwrap_or_default()
}

fn wasi_id(v: &Value) -> u64 {
    match v {
        Value::F64(n) => *n as u64,
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            o.properties
                .get("__wasi_id")
                .or_else(|| o.properties.get("__conn_id"))
                .map(|v| v.as_f64() as u64)
                .unwrap_or(0)
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

fn get_conn_id(args: &[Value]) -> u64 {
    args.first().map(|v| wasi_id(v)).unwrap_or(0)
}

fn get_sql(args: &[Value]) -> String {
    if let Some(Value::Object(obj)) = args.first() {
        let o = obj.lock().unwrap();
        let ct = o
            .properties
            .get("commandtext")
            .map(|v| format!("{}", v))
            .unwrap_or_default();
        if !ct.is_empty() {
            return ct;
        }
    }
    arg_str(args, 1)
}

// ── register ──────────────────────────────────────────────────────────────────

pub fn register(vm: &mut VM) {
    // wasi:sql/types ──────────────────────────────────────────────────────────

    vm.register_host_fn(
        "wasi:sql/types",
        "connection.new",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let raw = arg_str(args, 0);
            let normalized = if raw.is_empty() {
                String::new()
            } else {
                normalize_conn_str_full(&raw)
            };
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), string_value("SqlConnection"));
            obj.properties.insert("__conn_id".into(), Value::F64(0.0));
            obj.properties
                .insert("connectionstring".into(), string_value(&raw));
            obj.properties.insert(
                "provider".into(),
                string_value(driver_provider_name(&normalized)),
            );
            obj.properties
                .insert("serverversion".into(), string_value(""));
            obj.properties
                .insert("connectiontimeout".into(), Value::F64(30.0));
            obj.properties
                .insert("state".into(), string_value("Closed"));
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[static]connection.open",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let raw = arg_str(args, 0);
            if raw.is_empty() || raw == "null" {
                return make_error_obj("connection string is empty");
            }
            let url = normalize_conn_str_full(&raw);
            match open(&url) {
                Ok(driver) => {
                    let id = next_id();
                    state().lock().unwrap().conns.insert(
                        id,
                        ConnEntry {
                            driver,
                            url: url.clone(),
                        },
                    );
                    make_conn_obj(id, &url)
                }
                Err(e) => make_error_obj(&e),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]connection.open",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(obj)) = args.first() else {
                return Value::Null;
            };
            if get_conn_id(args) != 0 {
                return Value::Null;
            }
            let raw = string_prop(obj, "connectionstring");
            if raw.is_empty() || raw == "null" {
                return Value::Null;
            }
            let url = normalize_conn_str_full(&raw);
            match open(&url) {
                Ok(driver) => {
                    let id = next_id();
                    state()
                        .lock()
                        .unwrap()
                        .conns
                        .insert(id, ConnEntry { driver, url });
                    set_prop(obj, "__conn_id", Value::F64(id as f64));
                    set_prop(obj, "__wasi_id", Value::F64(id as f64));
                    set_prop(obj, "provider", string_value(driver_provider_name(&raw)));
                    set_prop(
                        obj,
                        "serverversion",
                        string_value(driver_server_version(&raw)),
                    );
                    set_prop(obj, "state", string_value("Open"));
                }
                Err(e) => eprintln!("wasi:sql/types connection.open: {}", e),
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[static]statement.prepare",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let query = arg_str(args, 0);
            let params = extract_params(args, 1);
            if query.is_empty() {
                return make_error_obj("empty query");
            }
            let id = next_id();
            state()
                .lock()
                .unwrap()
                .stmts
                .insert(id, StmtEntry { query, params });
            make_stmt_obj(id)
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]error.trace",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                if let Some(v) = obj.lock().unwrap().properties.get("message") {
                    return Value::String(Arc::from(format!("{}", v).as_str()));
                }
            }
            Value::String(Arc::from("unknown error"))
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]connection.close",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let conn_id = get_conn_id(args);
            if conn_id != 0 {
                state().lock().unwrap().conns.remove(&conn_id);
            }
            if let Some(Value::Object(obj)) = args.first() {
                set_prop(obj, "__conn_id", Value::F64(0.0));
                set_prop(obj, "__wasi_id", Value::F64(0.0));
                set_prop(obj, "state", string_value("Closed"));
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]connection.create-command",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            create_command_from_connection(args, "SqlCommand")
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]connection.begin-transaction",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let conn_id = get_conn_id(args);
            if conn_id == 0 {
                return Value::Null;
            }
            let _ = do_exec(conn_id, "BEGIN", &[]);
            make_transaction_obj(conn_id)
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]connection.get-schema",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let conn_id = get_conn_id(args);
            let schema = arg_str(args, 1).to_lowercase();
            let driver = match get_driver(conn_id) {
                Ok(driver) => driver,
                Err(_) => return data_table_from_rows("Schema", &[], &[]),
            };
            let (table_name, sql, fallback_columns) = match schema.as_str() {
                "" | "tables" => (
                    "Tables",
                    driver.tables_sql().to_string(),
                    vec!["name".to_string()],
                ),
                "columns" => {
                    let table = arg_str(args, 2);
                    (
                        "Columns",
                        driver.columns_sql(&table),
                        vec!["column_name".to_string()],
                    )
                }
                _ => return data_table_from_rows("Schema", &[], &[]),
            };
            match driver.query(&sql, &[]) {
                Ok(rows) => {
                    let col_names = if rows.is_empty() {
                        fallback_columns
                    } else {
                        col_names_from_rows(&rows)
                    };
                    data_table_from_rows(table_name, &rows, &col_names)
                }
                Err(_) => data_table_from_rows(table_name, &[], &fallback_columns),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "command.new",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let command_text = arg_str(args, 0);
            let conn_id = args.get(1).map(wasi_id).unwrap_or(0);
            let conn_string = match args.get(1) {
                Some(Value::Object(obj)) => string_prop(obj, "connectionstring"),
                _ => String::new(),
            };
            let command = make_command_obj("SqlCommand", conn_id, &conn_string);
            if let Value::Object(obj) = &command {
                set_prop(obj, "commandtext", string_value(&command_text));
            }
            command
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "data-adapter.new",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let sql = arg_str(args, 0);
            let conn_id = args.get(1).map(wasi_id).unwrap_or(0);
            let conn_string = match args.get(1) {
                Some(Value::Object(obj)) => string_prop(obj, "connectionstring"),
                _ => String::new(),
            };
            make_data_adapter_obj(&sql, conn_id, &conn_string)
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]command.execute-non-query",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(command)) = args.first() else {
                return Value::F64(-1.0);
            };
            let conn_id = get_conn_id(args);
            let (sql, params) = command_sql_and_params(command);
            match do_exec(conn_id, &sql, &params) {
                Ok(count) => Value::F64(count as f64),
                Err(e) => {
                    eprintln!("wasi:sql/types command.execute-non-query: {}", e);
                    Value::F64(-1.0)
                }
            }
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]command.execute-scalar",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(command)) = args.first() else {
                return Value::Null;
            };
            let conn_id = get_conn_id(args);
            let (sql, params) = command_sql_and_params(command);
            match do_query(conn_id, &sql, &params) {
                Ok(mut rows) if !rows.is_empty() => {
                    let row = rows.swap_remove(0);
                    let col_names = col_names_from_rows(&[row.clone()]);
                    col_names
                        .first()
                        .map(|name| row_value_by_name(&row, name))
                        .unwrap_or(Value::Null)
                }
                Ok(_) => Value::Null,
                Err(e) => {
                    eprintln!("wasi:sql/types command.execute-scalar: {}", e);
                    Value::Null
                }
            }
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]command.execute-reader",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(command)) = args.first() else {
                return make_reader_obj(vec![]);
            };
            let conn_id = get_conn_id(args);
            let (sql, params) = command_sql_and_params(command);
            match do_query(conn_id, &sql, &params) {
                Ok(rows) => {
                    let col_names = if rows.is_empty() {
                        do_query_columns(conn_id, &sql, &params).unwrap_or_default()
                    } else {
                        col_names_from_rows(&rows)
                    };
                    make_reader_obj_with_cols(rows, col_names)
                }
                Err(e) => {
                    eprintln!("wasi:sql/types command.execute-reader: {}", e);
                    make_reader_obj(vec![])
                }
            }
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]command.create-parameter",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let name = arg_str(args, 1);
            let value = args.get(4).cloned().unwrap_or(Value::Null);
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), string_value("SqlParameter"));
            obj.properties.insert("name".into(), string_value(&name));
            obj.properties.insert("value".into(), value);
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]params.add-with-value",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(params)) = args.first() else {
                return Value::Null;
            };
            let mut param = Object::new();
            param
                .properties
                .insert("__type".into(), string_value("SqlParameter"));
            param
                .properties
                .insert("name".into(), string_value(&arg_str(args, 1)));
            param
                .properties
                .insert("value".into(), args.get(2).cloned().unwrap_or(Value::Null));
            let Some(Value::Object(items)) =
                params.lock().unwrap().properties.get("__items").cloned()
            else {
                return Value::Null;
            };
            let mut items_guard = items.lock().unwrap();
            if let ObjectKind::Array(ref mut elems) = items_guard.kind {
                elems.push(Value::Object(Arc::new(Mutex::new(param))));
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]params.clear",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(params)) = args.first() else {
                return Value::Null;
            };
            let Some(Value::Object(items)) =
                params.lock().unwrap().properties.get("__items").cloned()
            else {
                return Value::Null;
            };
            let mut items_guard = items.lock().unwrap();
            if let ObjectKind::Array(ref mut elems) = items_guard.kind {
                elems.clear();
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]params.count",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(params)) = args.first() else {
                return Value::F64(0.0);
            };
            Value::F64(array_prop(params, "__items").len() as f64)
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]transaction.commit",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let result = Value::Bool(do_exec(get_conn_id(args), "COMMIT", &[]).is_ok());
            if let Some(Value::Object(obj)) = args.first() {
                set_prop(obj, "isclosed", Value::Bool(true));
            }
            result
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]transaction.rollback",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let result = Value::Bool(do_exec(get_conn_id(args), "ROLLBACK", &[]).is_ok());
            if let Some(Value::Object(obj)) = args.first() {
                set_prop(obj, "isclosed", Value::Bool(true));
            }
            result
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]reader.read",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(reader)) = args.first() else {
                return Value::Bool(false);
            };
            let next_pos = f64_prop(reader, "__pos") as isize + 1;
            let rows = array_prop(reader, "__rows");
            set_prop(reader, "__pos", Value::F64(next_pos as f64));
            Value::Bool(next_pos >= 0 && (next_pos as usize) < rows.len())
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]reader.get-value",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(reader)) = args.first() else {
                return Value::Null;
            };
            let Some(row) = current_reader_row(reader) else {
                return Value::Null;
            };
            row_value_by_index(&row, args.get(1).map(Value::as_f64).unwrap_or(0.0) as usize)
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]reader.get-string",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(reader)) = args.first() else {
                return string_value("");
            };
            let Some(row) = current_reader_row(reader) else {
                return string_value("");
            };
            string_value(&format!(
                "{}",
                row_value_by_index(&row, args.get(1).map(Value::as_f64).unwrap_or(0.0) as usize)
            ))
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]reader.get-name",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(reader)) = args.first() else {
                return string_value("");
            };
            let names = array_prop(reader, "__col_names");
            names
                .get(args.get(1).map(Value::as_f64).unwrap_or(0.0) as usize)
                .cloned()
                .unwrap_or_else(|| string_value(""))
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]reader.is-dbnull",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(reader)) = args.first() else {
                return Value::Bool(true);
            };
            let Some(row) = current_reader_row(reader) else {
                return Value::Bool(true);
            };
            Value::Bool(matches!(
                row_value_by_index(&row, args.get(1).map(Value::as_f64).unwrap_or(0.0) as usize),
                Value::Null
            ))
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]reader.close",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(reader)) = args.first() {
                set_prop(reader, "isclosed", Value::Bool(true));
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]reader.get-schema-table",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(reader)) = args.first() else {
                return data_table_from_rows("SchemaTable", &[], &[]);
            };
            let names = array_prop(reader, "__col_names");
            let mut rows = Vec::new();
            for (index, value) in names.iter().enumerate() {
                let mut row = Object::new();
                row.properties.insert("ColumnName".into(), value.clone());
                row.properties
                    .insert("ColumnOrdinal".into(), Value::F64(index as f64));
                rows.push(Value::Object(Arc::new(Mutex::new(row))));
            }
            let col_names = vec!["ColumnName".to_string(), "ColumnOrdinal".to_string()];
            data_table_from_rows("SchemaTable", &rows, &col_names)
        }),
    );

    vm.register_host_fn(
        "wasi:sql/types",
        "[method]adapter.fill",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(Value::Object(adapter)) = args.first() else {
                return Value::F64(0.0);
            };
            let Some(Value::Object(target)) = args.get(1) else {
                return Value::F64(0.0);
            };
            let conn_id = f64_prop(adapter, "__conn_id") as u64;
            let sql = string_prop(adapter, "selectcommand");
            let rows = match do_query(conn_id, &sql, &[]) {
                Ok(rows) => rows,
                Err(e) => {
                    eprintln!("wasi:sql/types adapter.fill: {}", e);
                    return Value::F64(0.0);
                }
            };
            let col_names = if rows.is_empty() {
                do_query_columns(conn_id, &sql, &[]).unwrap_or_default()
            } else {
                col_names_from_rows(&rows)
            };
            let target_type = string_prop(target, "__type");
            if target_type.eq_ignore_ascii_case("dataset") {
                let table_name = format!("Table{}", array_prop(target, "tables").len() + 1);
                let table = data_table_from_rows(&table_name, &rows, &col_names);
                let mut tables = array_prop(target, "tables");
                tables.push(table);
                set_prop(target, "tables", rows_array(tables));
            } else {
                set_prop(
                    target,
                    "columns",
                    rows_array(col_names.iter().map(|name| string_value(name)).collect()),
                );
                set_prop(target, "rows", rows_array(rows.clone()));
            }
            Value::F64(rows.len() as f64)
        }),
    );

    // wasi:sql/readwrite ──────────────────────────────────────────────────────

    vm.register_host_fn(
        "wasi:sql/readwrite",
        "query",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let conn_id = args.first().map(|v| wasi_id(v)).unwrap_or(0);
            let stmt_id = args.get(1).map(|v| wasi_id(v)).unwrap_or(0);
            let (sql, params) = {
                let s = state();
                let g = s.lock().unwrap();
                match g.stmts.get(&stmt_id) {
                    Some(e) => (e.query.clone(), e.params.clone()),
                    None => return rows_array(vec![]),
                }
            };
            match do_query(conn_id, &sql, &params) {
                Ok(rows) => rows_array(rows),
                Err(e) => {
                    eprintln!("wasi:sql/readwrite query: {}", e);
                    rows_array(vec![])
                }
            }
        }),
    );

    vm.register_host_fn(
        "wasi:sql/readwrite",
        "exec",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let conn_id = args.first().map(|v| wasi_id(v)).unwrap_or(0);
            let stmt_id = args.get(1).map(|v| wasi_id(v)).unwrap_or(0);
            let (sql, params) = {
                let s = state();
                let g = s.lock().unwrap();
                match g.stmts.get(&stmt_id) {
                    Some(e) => (e.query.clone(), e.params.clone()),
                    None => return Value::F64(-1.0),
                }
            };
            match do_exec(conn_id, &sql, &params) {
                Ok(n) => Value::F64(n as f64),
                Err(e) => {
                    eprintln!("wasi:sql/readwrite exec: {}", e);
                    Value::F64(-1.0)
                }
            }
        }),
    );

    register_flat_api(vm);
}

// ── Flat `wasi:sql` helpers for legacy adapter shapes ────────────────────────

fn register_flat_api(vm: &mut VM) {
    vm.register_host_fn(
        "wasi:sql",
        "connect",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let raw = arg_str(args, 0);
            if raw.is_empty() || raw == "null" {
                let mut obj = Object::new();
                obj.properties
                    .insert("__type".into(), Value::String(Arc::from("SqlConnection")));
                obj.properties.insert("__conn_id".into(), Value::F64(0.0));
                obj.properties
                    .insert("state".into(), Value::String(Arc::from("Closed")));
                return Value::Object(Arc::new(Mutex::new(obj)));
            }
            let url = normalize_conn_str_full(&raw);
            match open(&url) {
                Ok(driver) => {
                    let id = next_id();
                    state().lock().unwrap().conns.insert(
                        id,
                        ConnEntry {
                            driver,
                            url: url.clone(),
                        },
                    );
                    make_conn_obj(id, &url)
                }
                Err(e) => {
                    eprintln!("db.connect: {}", e);
                    Value::Null
                }
            }
        }),
    );

    vm.register_host_fn(
        "wasi:sql",
        "open",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let raw = obj
                    .lock()
                    .unwrap()
                    .properties
                    .get("connectionstring")
                    .map(|v| format!("{}", v))
                    .unwrap_or_default();
                if raw.is_empty() {
                    return Value::Null;
                }
                let url = normalize_conn_str_full(&raw);
                match open(&url) {
                    Ok(driver) => {
                        let id = next_id();
                        state()
                            .lock()
                            .unwrap()
                            .conns
                            .insert(id, ConnEntry { driver, url });
                        let mut o = obj.lock().unwrap();
                        o.properties
                            .insert("__conn_id".into(), Value::F64(id as f64));
                        o.properties
                            .insert("__wasi_id".into(), Value::F64(id as f64));
                        o.properties
                            .insert("state".into(), Value::String(Arc::from("Open")));
                    }
                    Err(e) => eprintln!("db.open: {}", e),
                }
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "wasi:sql",
        "close",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            state().lock().unwrap().conns.remove(&get_conn_id(args));
            Value::Null
        }),
    );

    vm.register_host_fn(
        "wasi:sql",
        "createCommand",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let conn_id = get_conn_id(args);
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("SqlCommand")));
            obj.properties
                .insert("__conn_id".into(), Value::F64(conn_id as f64));
            obj.properties
                .insert("commandtext".into(), Value::String(Arc::from("")));
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    vm.register_host_fn(
        "wasi:sql",
        "query",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let id = get_conn_id(args);
            let sql = get_sql(args);
            let params = extract_params(args, 2);
            match do_query(id, &sql, &params) {
                Ok(rows) => rows_array(rows),
                Err(e) => {
                    eprintln!("db.query: {}", e);
                    rows_array(vec![])
                }
            }
        }),
    );

    vm.register_host_fn(
        "wasi:sql",
        "execute",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let id = get_conn_id(args);
            let sql = get_sql(args);
            let params = extract_params(args, 2);
            match do_exec(id, &sql, &params) {
                Ok(n) => Value::F64(n as f64),
                Err(e) => {
                    eprintln!("db.execute: {}", e);
                    Value::F64(-1.0)
                }
            }
        }),
    );

    vm.register_host_fn(
        "wasi:sql",
        "scalar",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let id = get_conn_id(args);
            let sql = get_sql(args);
            let params = extract_params(args, 2);
            match do_query(id, &sql, &params) {
                Ok(mut rows) if !rows.is_empty() => {
                    let row = rows.swap_remove(0);
                    let col_names = col_names_from_rows(&[row.clone()]);
                    col_names
                        .first()
                        .map(|name| row_value_by_name(&row, name))
                        .unwrap_or(Value::Null)
                }
                _ => Value::Null,
            }
        }),
    );

    vm.register_host_fn(
        "wasi:sql",
        "beginTransaction",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(do_exec(get_conn_id(args), "BEGIN", &[]).is_ok())
        }),
    );

    vm.register_host_fn(
        "wasi:sql",
        "commit",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(do_exec(get_conn_id(args), "COMMIT", &[]).is_ok())
        }),
    );

    vm.register_host_fn(
        "wasi:sql",
        "rollback",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(do_exec(get_conn_id(args), "ROLLBACK", &[]).is_ok())
        }),
    );

    vm.register_host_fn(
        "wasi:sql",
        "tables",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let id = get_conn_id(args);
            let driver = {
                let s = state();
                let g = s.lock().unwrap();
                g.conns.get(&id).map(|e| Arc::clone(&e.driver))
            };
            let Some(driver) = driver else {
                return rows_array(vec![]);
            };
            match driver.query(driver.tables_sql(), &[]) {
                Ok(rows) => rows_array(
                    rows.into_iter()
                        .filter_map(|v| {
                            if let Value::Object(o) = v {
                                let row = Value::Object(Arc::clone(&o));
                                let col_names = col_names_from_rows(&[row.clone()]);
                                col_names.first().map(|name| row_value_by_name(&row, name))
                            } else {
                                None
                            }
                        })
                        .collect(),
                ),
                Err(_) => rows_array(vec![]),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:sql",
        "columns",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let id = get_conn_id(args);
            let table = arg_str(args, 1);
            let (driver, is_sqlite) = {
                let s = state();
                let g = s.lock().unwrap();
                match g.conns.get(&id) {
                    Some(e) => (Arc::clone(&e.driver), e.url.starts_with("sqlite:")),
                    None => return rows_array(vec![]),
                }
            };
            let sql = driver.columns_sql(&table);
            let col_idx = if is_sqlite { 1 } else { 0 }; // PRAGMA returns cid,name,type,...
            match driver.query(&sql, &[]) {
                Ok(rows) => rows_array(
                    rows.into_iter()
                        .filter_map(|v| {
                            if let Value::Object(o) = v {
                                let row = Value::Object(Arc::clone(&o));
                                let col_names = col_names_from_rows(&[row.clone()]);
                                col_names
                                    .get(col_idx)
                                    .map(|name| row_value_by_name(&row, name))
                            } else {
                                None
                            }
                        })
                        .collect(),
                ),
                Err(_) => rows_array(vec![]),
            }
        }),
    );
}

// ── ADO.NET / VB / PHP connection string normaliser ───────────────────────────

pub fn normalize_conn_str_full(raw: &str) -> String {
    let lower = raw.to_lowercase();

    if raw.starts_with("sqlite:")
        || raw.starts_with("postgres:")
        || raw.starts_with("postgresql:")
        || raw.starts_with("mysql:")
        || raw.starts_with("mysql2:")
        || raw.contains("://")
    {
        return raw.to_string();
    }

    if lower == ":memory:" || lower.contains("data source=:memory:") {
        return "sqlite::memory:".to_string();
    }

    if lower.starts_with("data source=") || lower.starts_with("datasource=") {
        let path = raw
            .splitn(2, '=')
            .nth(1)
            .unwrap_or("")
            .split(';')
            .next()
            .unwrap_or("")
            .trim()
            .trim_matches('"');
        return format!("sqlite:{}?mode=rwc", path);
    }

    if lower.contains("server=") || lower.contains("host=") {
        let pairs: HashMap<String, String> = raw
            .split(';')
            .filter_map(|p| {
                let mut kv = p.splitn(2, '=');
                Some((
                    kv.next()?.trim().to_lowercase(),
                    kv.next()?.trim().to_string(),
                ))
            })
            .collect();
        let host = pairs
            .get("server")
            .or(pairs.get("host"))
            .map(|s| s.as_str())
            .unwrap_or("localhost");
        let port = pairs.get("port").map(|s| s.as_str()).unwrap_or("3306");
        let db = pairs
            .get("database")
            .or(pairs.get("db"))
            .or(pairs.get("initial catalog"))
            .map(|s| s.as_str())
            .unwrap_or("");
        let user = pairs
            .get("uid")
            .or(pairs.get("user"))
            .or(pairs.get("user id"))
            .map(|s| s.as_str())
            .unwrap_or("root");
        let pass = pairs
            .get("pwd")
            .or(pairs.get("password"))
            .map(|s| s.as_str())
            .unwrap_or("");
        let scheme = if port == "5432" { "postgres" } else { "mysql" };
        return format!("{}://{}:{}@{}:{}/{}", scheme, user, pass, host, port, db);
    }

    format!("sqlite:{}?mode=rwc", raw)
}

// ── Public design-time / data-binding utilities ───────────────────────────────

pub fn test_connection_and_list_tables(conn_str: &str) -> Result<Vec<String>, String> {
    let url = normalize_conn_str_full(conn_str);
    let driver = open(&url)?;
    let sql = driver.tables_sql();
    let rows = driver
        .query(sql, &[])
        .map_err(|e| format!("Query failed: {}", e))?;
    Ok(rows
        .into_iter()
        .filter_map(|v| {
            let names = col_names_from_rows(std::slice::from_ref(&v));
            if let Some(name) = names
                .first()
                .map(|column| row_value_by_name(&v, column))
                .map(|value| format!("{}", value))
            {
                if !name.starts_with("sqlite_") && !name.starts_with("_sqlx") {
                    return Some(name);
                }
            }
            None
        })
        .collect())
}

pub fn fetch_columns_for_query(conn_str: &str, select_cmd: &str) -> Result<Vec<String>, String> {
    if conn_str.is_empty() || select_cmd.is_empty() {
        return Ok(vec![]);
    }
    let url = normalize_conn_str_full(conn_str);
    let driver = open(&url)?;
    driver.query_columns(select_cmd, &[])
}

pub fn query_rows(
    conn_str: &str,
    sql: &str,
) -> Result<(Vec<String>, Vec<HashMap<String, String>>), String> {
    let url = normalize_conn_str_full(conn_str);
    let driver = open(&url).map_err(|e| format!("Connection failed: {}", e))?;
    let rows = driver
        .query(sql, &[])
        .map_err(|e| format!("Query failed: {}", e))?;

    let columns = if rows.is_empty() {
        driver.query_columns(sql, &[]).unwrap_or_default()
    } else {
        col_names_from_rows(&rows)
    };

    let result = rows
        .into_iter()
        .filter_map(|v| {
            if matches!(v, Value::Object(_)) {
                Some(
                    columns
                        .iter()
                        .map(|column| {
                            (column.clone(), format!("{}", row_value_by_name(&v, column)))
                        })
                        .collect(),
                )
            } else {
                None
            }
        })
        .collect();

    Ok((columns, result))
}
