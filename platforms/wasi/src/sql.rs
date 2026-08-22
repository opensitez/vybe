//! `wasi:sql@0.2.0-draft` — SQL host implementation for Vybe.
//!
//! Three surfaces live here, and only the first is spec.
//!
//! **WIT** (`proposals/wasi-sql/wit/`) — five functions:
//!   - `wasi:sql/types`     → [static]connection.open, [static]statement.prepare,
//!                            [method]error.trace
//!   - `wasi:sql/readwrite` → query, exec
//! `resource connection` declares only `open` in `types.wit`; a resource's
//! teardown is its implicit destructor, so `[method]connection.close` below is
//! NOT the spec's — it belongs to the next surface.
//!
//! **The ADO.NET-shaped surface**, also under `wasi:sql/types`: connection /
//! command / reader / params / transaction / adapter members that mirror
//! `System.Data`. Vybe's own, in no proposal. Its arities come from the .NET
//! component descriptor in
//! `platforms/dotnet/src/emitter/core/component_classes_data_drawing.rs`,
//! which is what routes those calls.
//!
//! **The flat `wasi:sql` surface** for language adapters that have not moved to
//! the resource model: the PHP `mysqli`/`PDO` and Python DB-API emitters.
//!
//! Every function whose arity is single-valued carries a declared Component
//! Model signature (`sql_fn`); the ones that genuinely vary — the `*.new`
//! constructors and the optional-params `query`/`execute`/`scalar` — stay
//! undeclared, each with a comment saying why.
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

/// VM hot-reset (bucket C/D): drop all open SQL connections + prepared
/// statements. See `vmhotresetplan.md` and `vybe_host::reset_host_globals`.
pub fn reset() {
    state::reset();
}

use driver::{SqlDriver, open};
use state::{ConnEntry, StmtEntry, next_id, state};
use std::collections::HashMap;
use std::sync::{Arc, Mutex, OnceLock};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::vm::HostFnDecl;
use vybe_runtime::{FuncSig, HostContext, VM, ValType, Value};

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

// ── last error per connection (the WIT's `error.trace()`) ────────────────────
//
// `wasi:sql` models failure as `result<_, error>` where `error` is a resource
// with `trace() -> string`. Every failing call used to `eprintln!` the trace
// and hand back an empty row set or `-1`, so a guest could not tell a failed
// write from a successful one that changed no rows — `INSERT` into a read-only
// database "succeeded" in PDO, python's sqlite3 and ADO alike.
//
// The trace is recorded here instead, keyed by connection, which is the shape
// every consumer surface already wants: `sqlite3_errmsg(db)`,
// `mysql_error(conn)` and `PDO::errorInfo()` all ask a CONNECTION what went
// wrong last.

static LAST_ERROR: OnceLock<Mutex<HashMap<u64, String>>> = OnceLock::new();

fn last_error_table() -> &'static Mutex<HashMap<u64, String>> {
    LAST_ERROR.get_or_init(|| Mutex::new(HashMap::new()))
}

/// Record a failure against `conn_id`, and return the message so the caller can
/// still log it.
fn record_error(conn_id: u64, message: &str) {
    if let Ok(mut table) = last_error_table().lock() {
        table.insert(conn_id, message.to_string());
    }
}

/// Clear on success, so a stale trace never gets reported as a fresh failure.
fn clear_error(conn_id: u64) {
    if let Ok(mut table) = last_error_table().lock() {
        table.remove(&conn_id);
    }
}

fn take_error(conn_id: u64) -> String {
    last_error_table()
        .lock()
        .ok()
        .and_then(|table| table.get(&conn_id).cloned())
        .unwrap_or_default()
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
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn make_stmt_obj(id: u64) -> Value {
    let mut obj = Object::new();
    obj.properties.insert(
        "__wasi_kind".into(),
        Value::String(Arc::from("sql-statement")),
    );
    obj.properties
        .insert("__wasi_id".into(), Value::F64(id as f64));
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn make_error_obj(msg: &str) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__wasi_kind".into(), Value::String(Arc::from("sql-error")));
    obj.properties
        .insert("message".into(), Value::String(Arc::from(msg)));
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn rows_array(rows: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(rows)))
}

fn string_value(value: &str) -> Value {
    Value::String(Arc::from(value))
}

#[allow(dead_code)]
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

// ── Declaration ───────────────────────────────────────────────────────────────

/// Declare a `wasi:sql*` function's Component Model signature.
///
/// Only two of the three surfaces here are spec. `[static]connection.open`,
/// `[static]statement.prepare`, `[method]error.trace` and both
/// `wasi:sql/readwrite` functions come from `proposals/wasi-sql/wit/`, and
/// their types are quoted from it. Everything else on `wasi:sql/types` is the
/// ADO.NET-shaped surface Vybe invented for the .NET adapters, and the flat
/// `wasi:sql` interface is the pre-resource shape the PHP and Python adapters
/// still use — neither appears in any WIT, so no spec is cited for them.
///
/// For the invented names the authority for the arity is the .NET component
/// descriptor, because that is what actually routes the call: the lookup in
/// `platforms/dotnet/src/emitter/mod.rs` matches on `method.arity == arg_count`
/// and the emit in `primitives/calls.rs` pushes the receiver first, so the host
/// argc is the descriptor arity plus one.
///
/// Their parameter TYPES are mostly `Any`, and deliberately so: these calls
/// carry plain property-bag `Object`s, not resource handles, and writing
/// `Own`/`Borrow` here would claim resources that do not exist. `Any` is what
/// an untyped bag honestly declares as; the positions that really are a string
/// or an ordinal say so.
fn sql_fn(
    vm: &mut VM,
    module: &str,
    name: &str,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(HostFnDecl::new(module, name, call).with_sig(FuncSig {
        name: name.to_string(),
        params,
        results,
    }));
}

/// The receiver of an ADO-shaped `[method]` — a property-bag object, not a
/// resource handle. Named so the parameter lists read as what they are.
fn receiver() -> ValType {
    ValType::Any
}

// ── register ──────────────────────────────────────────────────────────────────

pub fn register(vm: &mut VM) {
    // wasi:sql/types ──────────────────────────────────────────────────────────

    // SPEC: `open: static func(name: string) -> result<connection, error>`
    // (`proposals/wasi-sql/wit/types.wit`).
    sql_fn(
        vm,
        "wasi:sql/types",
        "[static]connection.open",
        vec![ValType::String],
        vec![ValType::Result(
            Some(Box::new(ValType::Own("connection".to_string()))),
            Some(Box::new(ValType::Own("error".to_string()))),
        )],
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

    // SPEC: `prepare: static func(query: string, params: list<string>) ->
    // result<statement, error>` (`proposals/wasi-sql/wit/types.wit`).
    sql_fn(
        vm,
        "wasi:sql/types",
        "[static]statement.prepare",
        vec![ValType::String, ValType::List(Box::new(ValType::String))],
        vec![ValType::Result(
            Some(Box::new(ValType::Own("statement".to_string()))),
            Some(Box::new(ValType::Own("error".to_string()))),
        )],
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

    // SPEC: `trace: func() -> string` on `resource error` — a method, so the
    // borrowed receiver is the one parameter.
    sql_fn(
        vm,
        "wasi:sql/types",
        "[method]error.trace",
        vec![ValType::Borrow("error".to_string())],
        vec![ValType::String],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                if let Some(v) = obj.lock().unwrap().properties.get("message") {
                    return Value::String(Arc::from(format!("{}", v).as_str()));
                }
            }
            Value::String(Arc::from("unknown error"))
        }),
    );

    // wasi:sql/readwrite ──────────────────────────────────────────────────────

    // SPEC: `query: func(c: borrow<connection>, q: borrow<statement>) ->
    // result<list<row>, error>` (`proposals/wasi-sql/wit/readwrite.wit`).
    sql_fn(
        vm,
        "wasi:sql/readwrite",
        "query",
        vec![
            ValType::Borrow("connection".to_string()),
            ValType::Borrow("statement".to_string()),
        ],
        vec![ValType::Result(
            Some(Box::new(ValType::List(Box::new(ValType::Any)))),
            Some(Box::new(ValType::Own("error".to_string()))),
        )],
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

    // SPEC: `exec: func(c: borrow<connection>, q: borrow<statement>) ->
    // result<u32, error>` (`proposals/wasi-sql/wit/readwrite.wit`).
    sql_fn(
        vm,
        "wasi:sql/readwrite",
        "exec",
        vec![
            ValType::Borrow("connection".to_string()),
            ValType::Borrow("statement".to_string()),
        ],
        vec![ValType::Result(
            Some(Box::new(ValType::I32)),
            Some(Box::new(ValType::Own("error".to_string()))),
        )],
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
    sql_fn(
        vm,
        "wasi:sql",
        "connect",
        vec![ValType::String],
        vec![ValType::Any],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let raw = arg_str(args, 0);
            if raw.is_empty() || raw == "null" {
                let mut obj = Object::new();
                obj.properties
                    .insert("__type".into(), Value::String(Arc::from("SqlConnection")));
                obj.properties.insert("__conn_id".into(), Value::F64(0.0));
                obj.properties
                    .insert("state".into(), Value::String(Arc::from("Closed")));
                return Value::Object(vybe_runtime::heap::alloc(obj));
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

    sql_fn(
        vm,
        "wasi:sql",
        "open",
        vec![receiver()],
        vec![],
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

    sql_fn(
        vm,
        "wasi:sql",
        "close",
        vec![receiver()],
        vec![],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            state().lock().unwrap().conns.remove(&get_conn_id(args));
            Value::Null
        }),
    );

    sql_fn(
        vm,
        "wasi:sql",
        "createCommand",
        vec![receiver()],
        vec![ValType::Any],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let conn_id = get_conn_id(args);
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("SqlCommand")));
            obj.properties
                .insert("__conn_id".into(), Value::F64(conn_id as f64));
            obj.properties
                .insert("commandtext".into(), Value::String(Arc::from("")));
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    // `query`, `execute` and `scalar` stay UNDECLARED as one family: all three
    // read `extract_params(args, 2)`, so the params array is a genuinely
    // OPTIONAL trailing argument. `sql_adapter.rs` and `pdo_adapter.rs` branch
    // on `has_params` and emit 2 or 3 accordingly. The Component Model has no
    // optional parameter, so either arity would be a lie about the other, and
    // undeclared means UNKNOWN. (`scalar` happens to have only a 2-argument
    // caller today, but its closure is the same shape — declaring it at 2
    // would turn a legal 3-argument call into a false warning.)
    vm.register_host_fn(
        "wasi:sql",
        "query",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let id = get_conn_id(args);
            let sql = get_sql(args);
            let params = extract_params(args, 2);
            match do_query(id, &sql, &params) {
                Ok(rows) => {
                    clear_error(id);
                    rows_array(rows)
                }
                Err(e) => {
                    record_error(id, &e);
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
                Ok(n) => {
                    clear_error(id);
                    Value::F64(n as f64)
                }
                Err(e) => {
                    record_error(id, &e);
                    // `-1` stays the "it failed" signal every existing adapter
                    // already reads; the TRACE is what was missing.
                    Value::F64(-1.0)
                }
            }
        }),
    );

    // The flat surface's answer to `[method]error.trace` — the resource model
    // hands an `error` back from the failing call, but these adapters get a
    // plain row set or a count, so the trace is asked of the CONNECTION
    // afterwards. That is the shape `sqlite3_errmsg(db)`, `mysql_error(conn)`
    // and `PDO::errorInfo()` all want anyway. Empty string means "no failure
    // since the last successful call on this connection".
    sql_fn(
        vm,
        "wasi:sql",
        "lastError",
        vec![receiver()],
        vec![ValType::String],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::String(Arc::from(take_error(get_conn_id(args)).as_str()))
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

    // The three transaction verbs are reached through a variable NAME —
    // `emit_conn_op(.., op, ..)` in `python/sql_adapter.rs` and
    // `emit_transaction_verb(.., verb, ..)` in `php/pdo_adapter.rs` — but at a
    // fixed argc of 1, which is what the closures read.
    sql_fn(
        vm,
        "wasi:sql",
        "beginTransaction",
        vec![receiver()],
        vec![ValType::Bool],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(do_exec(get_conn_id(args), "BEGIN", &[]).is_ok())
        }),
    );

    sql_fn(
        vm,
        "wasi:sql",
        "commit",
        vec![receiver()],
        vec![ValType::Bool],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(do_exec(get_conn_id(args), "COMMIT", &[]).is_ok())
        }),
    );

    sql_fn(
        vm,
        "wasi:sql",
        "rollback",
        vec![receiver()],
        vec![ValType::Bool],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(do_exec(get_conn_id(args), "ROLLBACK", &[]).is_ok())
        }),
    );

    sql_fn(
        vm,
        "wasi:sql",
        "tables",
        vec![receiver()],
        vec![ValType::List(Box::new(ValType::String))],
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

    sql_fn(
        vm,
        "wasi:sql",
        "columns",
        vec![receiver(), ValType::String],
        vec![ValType::List(Box::new(ValType::String))],
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
        // A `file:` URI is SQLite's OWN qualified form and already carries its
        // options. Falling through to the bare-path arm below appended a
        // second query string — `file:x.db?mode=ro` became
        // `sqlite:file:x.db?mode=ro?mode=rwc`, and SQLite read the whole of
        // `ro?mode=rwc` as the access mode.
        || raw.starts_with("file:")
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
