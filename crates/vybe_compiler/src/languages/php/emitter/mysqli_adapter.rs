//! PHP `mysqli` / `mysqli_stmt` object surface (shape + runtime stubs).
//! Split out of the former `database_adapter.rs`; the PDO surface lives in
//! `pdo_adapter.rs`. Both share the same in-memory backend shape.

use crate::emitter::instructions::core_wasm;
use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use crate::emitter::collections;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn push_const(chunk: &mut Chunk, value: Value, line: u32) {
    match &value {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_op(Op::NULL, line),
        Value::BigInt(v) => chunk.emit_i64_const(v.to_i64_wrapping(), line),
        Value::String(s) => chunk.emit_string_const(&s, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),

        _ => {
            let _idx = chunk.add_constant(value);
        }
    }
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(value)), line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    chunks[current].emit_call(idx, argc, line);
}

fn struct_get_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_GET, idx, line);
}

fn struct_set_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_SET, idx, line);
    chunk.emit_op(Op::DROP, line);
}

fn global_set_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::GLOBAL_SET, idx, line);
}

fn global_get_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::GLOBAL_GET, idx, line);
}

fn emit_mark_queryish_prefix(
    chunk: &mut Chunk,
    sql_slot: u16,
    is_query_slot: u16,
    prefix: &str,
    line: u32,
) {
    lget(chunk, sql_slot, line);
    push_str(chunk, prefix, line);
    {
        let idx = chunk.add_import("ecma:string", "startsWith");
        chunk.emit_call(idx, 2, line);
    }
    chunk.emit_if(line);
    core_wasm::i32_const(chunk, line, 1);
    lset(chunk, is_query_slot, line);
    chunk.emit_end(line);
}

fn reset_mysqli_error_state(chunk: &mut Chunk, line: u32) {
    push_const(chunk, Value::F64(0.0), line);
    global_set_key(chunk, "__php_mysqli_connect_errno", line);
    push_str(chunk, "", line);
    global_set_key(chunk, "__php_mysqli_connect_error", line);
}

fn set_mysqli_error_state(chunk: &mut Chunk, errno: f64, error: &str, line: u32) {
    push_const(chunk, Value::F64(errno), line);
    global_set_key(chunk, "__php_mysqli_connect_errno", line);
    push_str(chunk, error, line);
    global_set_key(chunk, "__php_mysqli_connect_error", line);
}

fn emit_mysqli_result_fields(
    chunks: &mut [Chunk],
    current: usize,
    rows_slot: u16,
    line: u32,
) -> u16 {
    let fields_slot = alloc_local(&mut chunks[current]);

    {
        let chunk = &mut chunks[current];
        lget(chunk, rows_slot, line);
        crate::emitter::ops::emit_dyn_ne(chunk, line);
        crate::emitter::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);

        lget(chunk, rows_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        chunk.emit_op(Op::ARRAY_GET, line);
    }
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, fields_slot, line);
        chunk.emit_else(line);
    }
    collections::emit_array_new(chunks, current, 0, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, fields_slot, line);
        chunk.emit_end(line);
    }

    fields_slot
}

fn emit_mysqli_field_object(
    chunks: &mut [Chunk],
    current: usize,
    field_name_slot: u16,
    line: u32,
) -> u16 {
    call_import(chunks, current, "ecma:object", "new", 0, line);
    let field_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, field_slot, line);

    lget(chunk, field_slot, line);
    lget(chunk, field_name_slot, line);
    struct_set_key(chunk, "name", line);

    lget(chunk, field_slot, line);
    push_str(chunk, "", line);
    struct_set_key(chunk, "table", line);

    lget(chunk, field_slot, line);
    push_str(chunk, "", line);
    struct_set_key(chunk, "def", line);

    lget(chunk, field_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, "max_length", line);

    lget(chunk, field_slot, line);
    push_const(chunk, Value::Bool(false), line);
    struct_set_key(chunk, "not_null", line);

    lget(chunk, field_slot, line);
    push_const(chunk, Value::Bool(false), line);
    struct_set_key(chunk, "primary_key", line);

    lget(chunk, field_slot, line);
    push_const(chunk, Value::Bool(false), line);
    struct_set_key(chunk, "multiple_key", line);

    lget(chunk, field_slot, line);
    push_const(chunk, Value::Bool(false), line);
    struct_set_key(chunk, "unique_key", line);

    lget(chunk, field_slot, line);
    push_const(chunk, Value::Bool(false), line);
    struct_set_key(chunk, "numeric", line);

    lget(chunk, field_slot, line);
    push_const(chunk, Value::Bool(false), line);
    struct_set_key(chunk, "blob", line);

    lget(chunk, field_slot, line);
    push_str(chunk, "string", line);
    struct_set_key(chunk, "type", line);

    lget(chunk, field_slot, line);
    push_const(chunk, Value::Bool(false), line);
    struct_set_key(chunk, "unsigned", line);

    lget(chunk, field_slot, line);
    push_const(chunk, Value::Bool(false), line);
    struct_set_key(chunk, "zerofill", line);

    lget(chunk, field_slot, line);
    field_slot
}

fn emit_mysqli_result_object(
    chunks: &mut [Chunk],
    current: usize,
    rows_slot: u16,
    line: u32,
) -> u16 {
    call_import(chunks, current, "ecma:object", "new", 0, line);
    let result_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, result_slot, line);

    lget(chunk, result_slot, line);
    push_str(chunk, "mysqli_result", line);
    struct_set_key(chunk, "__type", line);

    lget(chunk, result_slot, line);
    lget(chunk, rows_slot, line);
    struct_set_key(chunk, "__rows", line);

    lget(chunk, result_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, "__cursor", line);

    lget(chunk, result_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, "__field_cursor", line);

    let fields_slot = emit_mysqli_result_fields(chunks, current, rows_slot, line);
    let chunk = &mut chunks[current];
    lget(chunk, result_slot, line);
    lget(chunk, fields_slot, line);
    struct_set_key(chunk, "__fields", line);

    lget(chunk, result_slot, line);
    result_slot
}

pub fn emit_php_mysqli_report(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_op(Op::NULL, line);
}

pub fn emit_php_mysqli_connect(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // For bootstrap capability probes, model mysqli_connect as a successful
    // constructor-shaped connection object.
    emit_php_mysqli_init(chunks, current, argc, line);
}

pub fn emit_php_mysqli_init(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }

    reset_mysqli_error_state(chunk, line);

    // Open a real `wasi:sql` connection — mysqli is a thin adapter over the
    // same backend as PDO. The test environment has no MySQL server, so we
    // back it with the offline in-memory sqlite driver; production would pass
    // a `mysql://` URL here once a server is reachable.
    push_str(chunk, "sqlite::memory:", line);
    call_import(chunks, current, "wasi:sql", "connect", 1, line);
    let conn_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, conn_slot, line);

    // Stamp the mysqli class identity + shape fields over the connection.
    lget(chunk, conn_slot, line);
    push_str(chunk, "mysqli", line);
    struct_set_key(chunk, "__type", line);

    lget(chunk, conn_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, "connect_errno", line);

    lget(chunk, conn_slot, line);
    push_str(chunk, "", line);
    struct_set_key(chunk, "connect_error", line);

    lget(chunk, conn_slot, line);
    push_str(chunk, "", line);
    struct_set_key(chunk, "error", line);

    lget(chunk, conn_slot, line);
}

pub fn emit_php_mysqli_real_connect(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];

    // mysqli_real_connect(dbh, host, user, password, database, port, socket, flags)
    // Args are in reverse order on stack: flags, socket, port, database, password, user, host, dbh
    let flags_slot = if argc >= 8 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let socket_slot = if argc >= 7 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let port_slot = if argc >= 6 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let database_slot = if argc >= 5 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let password_slot = if argc >= 4 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let user_slot = if argc >= 3 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let host_slot = if argc >= 2 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let dbh_slot = alloc_local(chunk);

    if let Some(slot) = flags_slot {
        lset(chunk, slot, line);
    }
    if let Some(slot) = socket_slot {
        lset(chunk, slot, line);
    }
    if let Some(slot) = port_slot {
        lset(chunk, slot, line);
    }
    if let Some(slot) = database_slot {
        lset(chunk, slot, line);
    }
    if let Some(slot) = password_slot {
        lset(chunk, slot, line);
    }
    if let Some(slot) = user_slot {
        lset(chunk, slot, line);
    }
    if let Some(slot) = host_slot {
        lset(chunk, slot, line);
    }
    lset(chunk, dbh_slot, line);

    // Build MySQL connection URL: mysql://user:password@host:port/database
    let url_slot = alloc_local(chunk);
    push_str(chunk, "mysql://", line);
    lset(chunk, url_slot, line);

    // Append user if provided
    if let Some(slot) = user_slot {
        lget(chunk, slot, line);
        push_str(chunk, "", line);
        crate::emitter::ops::emit_dyn_ne(chunk, line);
        chunk.emit_if(line);
        lget(chunk, url_slot, line);
        lget(chunk, slot, line);
        {
            let idx = chunk.add_import("wasm:js-string", "concat");
            chunk.emit_call(idx, 2, line);
        }
        lset(chunk, url_slot, line);
        chunk.emit_end(line);
    }

    // Append password if provided
    if let Some(pass_slot) = password_slot {
        lget(chunk, pass_slot, line);
        push_str(chunk, "", line);
        crate::emitter::ops::emit_dyn_ne(chunk, line);
        chunk.emit_if(line);
        lget(chunk, url_slot, line);
        push_str(chunk, ":", line);
        {
            let idx = chunk.add_import("wasm:js-string", "concat");
            chunk.emit_call(idx, 2, line);
        }
        lget(chunk, pass_slot, line);
        {
            let idx = chunk.add_import("wasm:js-string", "concat");
            chunk.emit_call(idx, 2, line);
        }
        lset(chunk, url_slot, line);
        chunk.emit_end(line);
    }

    // Append @host
    lget(chunk, url_slot, line);
    push_str(chunk, "@", line);
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
    if let Some(slot) = host_slot {
        lget(chunk, slot, line);
    } else {
        push_str(chunk, "localhost", line);
    }
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, url_slot, line);

    // Note: port handling skipped for simplicity; MySQL will use default port
    // (can be extended later if needed for non-standard ports)

    // Append /database if provided
    if let Some(slot) = database_slot {
        lget(chunk, slot, line);
        push_str(chunk, "", line);
        crate::emitter::ops::emit_dyn_ne(chunk, line);
        chunk.emit_if(line);
        lget(chunk, url_slot, line);
        push_str(chunk, "/", line);
        {
            let idx = chunk.add_import("wasm:js-string", "concat");
            chunk.emit_call(idx, 2, line);
        }
        lget(chunk, slot, line);
        {
            let idx = chunk.add_import("wasm:js-string", "concat");
            chunk.emit_call(idx, 2, line);
        }
        lset(chunk, url_slot, line);
        chunk.emit_end(line);
    }

    // Call wasi:sql.connect with the built URL
    lget(chunk, url_slot, line);
    call_import(chunks, current, "wasi:sql", "connect", 1, line);
    let conn_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, conn_slot, line);

    // Check if connection failed (null)
    lget(chunk, conn_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);

    // Connection failed
    set_mysqli_error_state(chunk, 1.0, "Connection failed", line);
    lget(chunk, dbh_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    struct_set_key(chunk, "connect_errno", line);
    lget(chunk, dbh_slot, line);
    push_str(chunk, "Connection failed", line);
    struct_set_key(chunk, "connect_error", line);
    lget(chunk, dbh_slot, line);
    push_str(chunk, "Connection failed", line);
    struct_set_key(chunk, "error", line);
    push_const(chunk, Value::Bool(false), line);

    chunk.emit_else(line);

    // Connection succeeded - update dbh with the connection
    reset_mysqli_error_state(chunk, line);
    lget(chunk, dbh_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, "connect_errno", line);
    lget(chunk, dbh_slot, line);
    push_str(chunk, "", line);
    struct_set_key(chunk, "connect_error", line);
    lget(chunk, dbh_slot, line);
    lget(chunk, conn_slot, line);
    struct_set_key(chunk, "__connection", line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_end(line);
}

pub fn emit_php_mysqli_connect_errno(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    global_get_key(chunk, "__php_mysqli_connect_errno", line);
}

pub fn emit_php_mysqli_connect_error(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    global_get_key(chunk, "__php_mysqli_connect_error", line);
}

pub fn emit_php_mysqli_error(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    global_get_key(chunk, "__php_mysqli_connect_error", line);
}

pub fn emit_php_mysqli_query(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let sql_slot = alloc_local(chunk);
    let dbh_slot = alloc_local(chunk);
    lset(chunk, sql_slot, line);
    lset(chunk, dbh_slot, line);
    lget(chunk, sql_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "trim");
        chunk.emit_call(idx, 1, line);
    }
    {
        let idx = chunk.add_import("ecma:string", "toLowerCase");
        chunk.emit_call(idx, 1, line);
    }
    let normalized_sql_slot = alloc_local(chunk);
    lset(chunk, normalized_sql_slot, line);

    core_wasm::i32_const(chunk, line, 0);
    let is_query_slot = alloc_local(chunk);
    lset(chunk, is_query_slot, line);
    emit_mark_queryish_prefix(chunk, normalized_sql_slot, is_query_slot, "select", line);
    emit_mark_queryish_prefix(chunk, normalized_sql_slot, is_query_slot, "pragma", line);
    emit_mark_queryish_prefix(chunk, normalized_sql_slot, is_query_slot, "show", line);
    emit_mark_queryish_prefix(chunk, normalized_sql_slot, is_query_slot, "with", line);
    emit_mark_queryish_prefix(chunk, normalized_sql_slot, is_query_slot, "describe", line);
    emit_mark_queryish_prefix(chunk, normalized_sql_slot, is_query_slot, "explain", line);

    lget(chunk, is_query_slot, line);
    chunk.emit_if_value(line);
    lget(chunk, dbh_slot, line);
    struct_get_key(chunk, "__connection", line);
    let conn_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, conn_slot, line);

    lget(chunk, conn_slot, line);
    lget(chunk, sql_slot, line);
    call_import(chunks, current, "wasi:sql", "query", 2, line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);

    lget(chunk, dbh_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, "affected_rows", line);
    lget(chunk, dbh_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, "insert_id", line);
    lget(chunk, dbh_slot, line);
    push_str(chunk, "", line);
    struct_set_key(chunk, "error", line);

    let result_slot = emit_mysqli_result_object(chunks, current, rows_slot, line);
    let chunk = &mut chunks[current];
    lget(chunk, result_slot, line);
    chunk.emit_else(line);

    lget(chunk, dbh_slot, line);
    struct_get_key(chunk, "__connection", line);
    let conn_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, conn_slot, line);

    lget(chunk, conn_slot, line);
    lget(chunk, sql_slot, line);
    call_import(chunks, current, "wasi:sql", "execute", 2, line);
    let count_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, count_slot, line);

    lget(chunk, dbh_slot, line);
    lget(chunk, count_slot, line);
    struct_set_key(chunk, "affected_rows", line);
    lget(chunk, dbh_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, "insert_id", line);
    lget(chunk, dbh_slot, line);
    push_str(chunk, "", line);
    struct_set_key(chunk, "error", line);
    push_const(chunk, Value::Bool(true), line);

    chunk.emit_end(line);
}

pub fn emit_php_mysqli_prepare(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let sql_slot = alloc_local(chunk);
    let dbh_slot = alloc_local(chunk);
    lset(chunk, sql_slot, line);
    lset(chunk, dbh_slot, line);

    // Create a statement object
    call_import(chunks, current, "ecma:object", "new", 0, line);
    let stmt_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, stmt_slot, line);

    lget(chunk, stmt_slot, line);
    push_str(chunk, "mysqli_stmt", line);
    struct_set_key(chunk, "__type", line);

    lget(chunk, stmt_slot, line);
    lget(chunk, dbh_slot, line);
    struct_set_key(chunk, "__mysqli", line);

    lget(chunk, stmt_slot, line);
    lget(chunk, sql_slot, line);
    struct_set_key(chunk, "__sql", line);

    lget(chunk, stmt_slot, line);
}

pub fn emit_php_mysqli_select_db(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let db_slot = alloc_local(chunk);
    let dbh_slot = alloc_local(chunk);
    lset(chunk, db_slot, line);
    lset(chunk, dbh_slot, line);

    lget(chunk, dbh_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);

    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);
    lget(chunk, dbh_slot, line);
    lget(chunk, db_slot, line);
    struct_set_key(chunk, "selected_db", line);
    lget(chunk, dbh_slot, line);
    lget(chunk, db_slot, line);
    struct_set_key(chunk, "database", line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_end(line);
}

pub fn emit_php_mysqli_set_charset(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let charset_slot = alloc_local(chunk);
    let dbh_slot = alloc_local(chunk);
    lset(chunk, charset_slot, line);
    lset(chunk, dbh_slot, line);

    lget(chunk, dbh_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);

    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);
    lget(chunk, dbh_slot, line);
    lget(chunk, charset_slot, line);
    struct_set_key(chunk, "charset", line);
    lget(chunk, dbh_slot, line);
    lget(chunk, charset_slot, line);
    struct_set_key(chunk, "character_set_name", line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_end(line);
}

pub fn emit_php_mysqli_ping(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let dbh_slot = alloc_local(chunk);
    lset(chunk, dbh_slot, line);

    lget(chunk, dbh_slot, line);
    struct_get_key(chunk, "__connection", line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_end(line);
}

pub fn emit_php_mysqli_errno(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    global_get_key(chunk, "__php_mysqli_connect_errno", line);
}

pub fn emit_php_mysqli_affected_rows(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let dbh_slot = alloc_local(chunk);
    lset(chunk, dbh_slot, line);
    lget(chunk, dbh_slot, line);
    struct_get_key(chunk, "affected_rows", line);
}

pub fn emit_php_mysqli_insert_id(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let dbh_slot = alloc_local(chunk);
    lset(chunk, dbh_slot, line);
    lget(chunk, dbh_slot, line);
    struct_get_key(chunk, "insert_id", line);
}

pub fn emit_php_mysqli_num_fields(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let result_slot = alloc_local(chunk);
    lset(chunk, result_slot, line);

    lget(chunk, result_slot, line);
    struct_get_key(chunk, "__fields", line);
    collections::emit_len(chunks, current, line);
}

pub fn emit_php_mysqli_fetch_field(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let result_slot = alloc_local(chunk);
    lset(chunk, result_slot, line);

    lget(chunk, result_slot, line);
    struct_get_key(chunk, "__fields", line);
    let fields_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, fields_slot, line);

    lget(chunk, result_slot, line);
    struct_get_key(chunk, "__field_cursor", line);
    let cursor_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, cursor_slot, line);

    {
        let chunk = &mut chunks[current];
        lget(chunk, fields_slot, line);
        lget(chunk, cursor_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
    }
    let field_name_slot = alloc_local(&mut chunks[current]);
    {
        let chunk = &mut chunks[current];
        lset(chunk, field_name_slot, line);
    }

    {
        let chunk = &mut chunks[current];
        lget(chunk, field_name_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if_value(line);
        chunk.emit_op(Op::NULL, line);
        chunk.emit_else(line);
    }

    {
        let chunk = &mut chunks[current];
        lget(chunk, result_slot, line);
        lget(chunk, cursor_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        struct_set_key(chunk, "__field_cursor", line);
    }

    let field_slot = emit_mysqli_field_object(chunks, current, field_name_slot, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, field_slot, line);
        chunk.emit_end(line);
    }
}

pub fn emit_php_mysqli_free_result(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let result_slot = {
        let chunk = &mut chunks[current];
        let result_slot = alloc_local(chunk);
        lset(chunk, result_slot, line);
        result_slot
    };

    {
        let chunk = &mut chunks[current];
        lget(chunk, result_slot, line);
        chunk.emit_op(Op::DROP, line);
    }
    collections::emit_array_new(chunks, current, 0, line);
    {
        let chunk = &mut chunks[current];
        struct_set_key(chunk, "__rows", line);
    }

    {
        let chunk = &mut chunks[current];
        lget(chunk, result_slot, line);
        chunk.emit_op(Op::DROP, line);
    }
    collections::emit_array_new(chunks, current, 0, line);
    {
        let chunk = &mut chunks[current];
        struct_set_key(chunk, "__fields", line);
    }

    {
        let chunk = &mut chunks[current];
        lget(chunk, result_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        struct_set_key(chunk, "__cursor", line);

        lget(chunk, result_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        struct_set_key(chunk, "__field_cursor", line);

        push_const(chunk, Value::Bool(true), line);
    }
}

pub fn emit_php_mysqli_more_results(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    push_const(chunk, Value::Bool(false), line);
}

pub fn emit_php_mysqli_next_result(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    push_const(chunk, Value::Bool(false), line);
}

pub fn emit_php_mysqli_close(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let dbh_slot = alloc_local(chunk);
    lset(chunk, dbh_slot, line);

    lget(chunk, dbh_slot, line);
    struct_get_key(chunk, "__connection", line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);

    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);

    {
        let chunk = &mut chunks[current];
        lget(chunk, dbh_slot, line);
        struct_get_key(chunk, "__connection", line);
    }
    call_import(chunks, current, "wasi:sql", "close", 1, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, dbh_slot, line);
        chunk.emit_op(Op::NULL, line);
        struct_set_key(chunk, "__connection", line);
        push_const(chunk, Value::Bool(true), line);
        chunk.emit_end(line);
    }
}

pub fn emit_php_mysqli_real_escape_string(
    chunks: &mut [Chunk],
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let data_slot = alloc_local(chunk);
    let dbh_slot = alloc_local(chunk);
    lset(chunk, data_slot, line);
    lset(chunk, dbh_slot, line);

    lget(chunk, data_slot, line);
    push_str(chunk, "\\", line);
    push_str(chunk, "\\\\", line);
    {
        let idx = chunk.add_import("ecma:string", "replaceAll");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, data_slot, line);

    lget(chunk, data_slot, line);
    push_str(chunk, "'", line);
    push_str(chunk, "\\'", line);
    {
        let idx = chunk.add_import("ecma:string", "replaceAll");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, data_slot, line);

    lget(chunk, data_slot, line);
}

pub fn emit_php_mysqli_character_set_name(
    chunks: &mut [Chunk],
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let dbh_slot = alloc_local(chunk);
    lset(chunk, dbh_slot, line);

    lget(chunk, dbh_slot, line);
    struct_get_key(chunk, "charset", line);
    let charset_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, charset_slot, line);

    lget(chunk, charset_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    push_str(chunk, "utf8mb4", line);
    chunk.emit_else(line);
    lget(chunk, charset_slot, line);
    chunk.emit_end(line);
}

pub fn emit_php_mysqli_get_client_info(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    push_str(chunk, "mysqlnd 8.0.0", line);
}

pub fn emit_php_mysqli_get_server_info(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    push_str(chunk, "8.0.0", line);
}

pub fn emit_php_mysqli_fetch_array(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let flags_slot = if argc >= 2 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let result_slot = alloc_local(chunk);
    if let Some(slot) = flags_slot {
        lset(chunk, slot, line);
    }
    lset(chunk, result_slot, line);

    lget(chunk, result_slot, line);
    struct_get_key(chunk, "__rows", line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);

    lget(chunk, result_slot, line);
    struct_get_key(chunk, "__cursor", line);
    let cursor_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, cursor_slot, line);

    lget(chunk, rows_slot, line);
    lget(chunk, cursor_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let row_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, row_slot, line);

    lget(chunk, result_slot, line);
    lget(chunk, cursor_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    struct_set_key(chunk, "__cursor", line);

    lget(chunk, row_slot, line);
}

pub fn emit_php_mysqli_fetch_assoc(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let result_slot = alloc_local(chunk);
    lset(chunk, result_slot, line);

    lget(chunk, result_slot, line);
    struct_get_key(chunk, "__rows", line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);

    lget(chunk, result_slot, line);
    struct_get_key(chunk, "__cursor", line);
    let cursor_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, cursor_slot, line);

    lget(chunk, rows_slot, line);
    lget(chunk, cursor_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let row_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, row_slot, line);

    lget(chunk, result_slot, line);
    lget(chunk, cursor_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    struct_set_key(chunk, "__cursor", line);

    lget(chunk, row_slot, line);
}

pub fn emit_php_mysqli_fetch_object(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // Same as fetch_assoc for now
    emit_php_mysqli_fetch_assoc(chunks, current, 1, line);
}

pub fn emit_php_mysqli_num_rows(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let result_slot = alloc_local(chunk);
    lset(chunk, result_slot, line);

    lget(chunk, result_slot, line);
    struct_get_key(chunk, "__rows", line);
    collections::emit_len(chunks, current, line);
}

pub fn emit_php_mysqli_fetch_all(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let result_slot = alloc_local(chunk);
    lset(chunk, result_slot, line);

    lget(chunk, result_slot, line);
    struct_get_key(chunk, "__rows", line);
}

// ── mysqli_stmt method stubs ────────────────────────────────────────────────
// The prepared-statement surface is mostly shape/state stubs over the shared
// `db_adapter` statement object. Each drops the receiver + args and returns the
// documented sentinel; row-count/id state lives as stamped properties.

/// Drop the receiver + `argc` args, then push `result` (a bool sentinel).
fn emit_stmt_sentinel(chunks: &mut [Chunk], current: usize, argc: u8, line: u32, result: bool) {
    let chunk = &mut chunks[current];
    for _ in 0..(argc as u16 + 1) {
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_bool_const(result, line);
}

/// `$stmt->bind_param`/`store_result`/`free_result`/`reset`/`close`/
/// `send_long_data`/`data_seek` → `true` (no-op stubs).
pub fn emit_mysqli_stmt_true(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_stmt_sentinel(chunks, current, argc, line, true);
}

/// `$stmt->get_warnings`/`more_results`/`next_result` → `false`.
pub fn emit_mysqli_stmt_false(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_stmt_sentinel(chunks, current, argc, line, false);
}

/// `$stmt->get_result()` → a `mysqli_result` wrapping the statement's executed
/// rows (populated by the shared `execute`). Stack: `[stmt]` → `[result]`.
pub fn emit_php_mysqli_stmt_get_result(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let stmt_slot = alloc_local(chunk);
    lset(chunk, stmt_slot, line);
    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, "__rows", line);
    let rows_slot = alloc_local(chunk);
    lset(chunk, rows_slot, line);
    let result_slot = emit_mysqli_result_object(chunks, current, rows_slot, line);
    lget(&mut chunks[current], result_slot, line);
}

/// `$stmt->attr_get($attr)` → an integer attribute value (0 stub).
pub fn emit_mysqli_stmt_attr_get(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..(argc as u16 + 1) {
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_i32_const(0, line);
}
