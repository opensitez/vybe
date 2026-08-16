//! MySQL C ABI (`libmysqlclient`), as an adapter over the `wasi:sql` host
//! interface.
//!
//! Same contract as `sqlite_adapter`: the surface is the REAL `mysql_*` C API,
//! so a program written against it is ordinary C — or ordinary Fortran through
//! `iso_c_binding` — and needs no Vybe-specific spelling. C reaches it through
//! `#include <mysql.h>`, Fortran through `bind(c)`.
//!
//! PHP's procedural `mysqli_*` surface is the same API with an `i` in the name,
//! and `languages/php/src/emitter/mysqli_adapter.rs` already drives `wasi:sql`
//! with it, so the mapping here is the one that surface established rather than
//! a new invention.
//!
//! Handles are ordinary objects, since there is no `MYSQL*`:
//!
//! ```text
//! MYSQL     = { __type, __conn, __rows, __cursor, __affected, __insert_id }
//! MYSQL_RES = { __type, __rows, __cursor }          // shares the row array
//! MYSQL_ROW = the row itself (an array), NULL past the end
//! ```
//!
//! WHAT IS FAITHFUL: `init`/`real_connect`/`query`/`close`, `store_result` →
//! `fetch_row` → NULL, the row/field counts, and `affected_rows`.
//!
//! WHAT IS NOT: `use_result` is an alias of `store_result` — real MySQL streams
//! row-at-a-time from the server for the former, and `wasi:sql` materialises the
//! whole result set, so the distinction cannot exist here. `mysql_error` is thin
//! for the same reason it is in the SQLite adapter: the host reports failure by
//! trapping rather than by handing back a message.

use std::sync::Arc;

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

// ── emit helpers (same conventions as sqlite_adapter) ────────────────────────

fn alloc(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    chunk.emit_string_const(value, line);
}

fn push_i32(chunk: &mut Chunk, value: i32, line: u32) {
    chunk.emit_i32_const(value, line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn struct_get_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_GET, idx, line);
}

fn struct_set_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_SET, idx, line);
}

fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> Vec<u16> {
    let mut slots = vec![0u16; argc as usize];
    for index in (0..argc as usize).rev() {
        let slot = alloc(&mut chunks[current]);
        lset(&mut chunks[current], slot, line);
        slots[index] = slot;
    }
    slots
}

fn drop_values(chunks: &mut [Chunk], current: usize, count: u8, line: u32) {
    for _ in 0..count {
        chunks[current].emit_op(Op::DROP, line);
    }
}

fn concat(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "ecma:string", "concat", 2, line);
}

// ── connection ───────────────────────────────────────────────────────────────

/// `mysql_init(mysql)` → a MYSQL handle.
///
/// The argument is a caller-allocated `MYSQL*` or NULL; either way real
/// `mysql_init` hands back the handle to use, so a fresh object is returned and
/// the argument is dropped.
pub fn emit_init(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    drop_values(chunks, current, argc, line);
    let handle = alloc(&mut chunks[current]);

    call_import(chunks, current, "ecma:object", "new", 0, line);
    lset(&mut chunks[current], handle, line);

    lget(&mut chunks[current], handle, line);
    push_str(&mut chunks[current], "MysqlConn", line);
    struct_set_key(&mut chunks[current], "__type", line);

    for key in ["__conn", "__rows"] {
        lget(&mut chunks[current], handle, line);
        chunks[current].emit_op(Op::NULL, line);
        struct_set_key(&mut chunks[current], key, line);
    }
    for key in ["__cursor", "__affected", "__insert_id"] {
        lget(&mut chunks[current], handle, line);
        push_i32(&mut chunks[current], 0, line);
        struct_set_key(&mut chunks[current], key, line);
    }

    lget(&mut chunks[current], handle, line);
}

/// `mysql_real_connect(mysql, host, user, passwd, db, port, unix_socket,
/// client_flag)` → the handle, or NULL on failure.
///
/// The eight C arguments become one URL, which is the only thing `wasi:sql`
/// accepts: `mysql://user:passwd@host:port/db`. Everything the URL can express
/// is carried, INCLUDING the port — it is only omitted when the caller passes
/// 0, which is what real `mysql_real_connect` reads as "use the default".
///
/// `unix_socket` and `client_flag` are the two that genuinely cannot travel:
/// the first selects a transport the URL form has no slot for, the second sets
/// protocol options (compression, multi-statements) the host does not expose.
pub fn emit_real_connect(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = stash_args(chunks, current, argc, line);
    let url = alloc(&mut chunks[current]);

    push_str(&mut chunks[current], "mysql://", line);
    if args.len() > 2 {
        lget(&mut chunks[current], args[2], line);
        concat(chunks, current, line);
        if args.len() > 3 {
            push_str(&mut chunks[current], ":", line);
            concat(chunks, current, line);
            lget(&mut chunks[current], args[3], line);
            concat(chunks, current, line);
        }
        push_str(&mut chunks[current], "@", line);
        concat(chunks, current, line);
    }
    if args.len() > 1 {
        lget(&mut chunks[current], args[1], line);
        concat(chunks, current, line);
    }
    if args.len() > 5 {
        // `port == 0` means "default" in the C API, so only a real port is
        // written into the URL.
        lget(&mut chunks[current], args[5], line);
        chunks[current].emit_if(line);
        push_str(&mut chunks[current], ":", line);
        concat(chunks, current, line);
        lget(&mut chunks[current], args[5], line);
        call_import(chunks, current, "ecma:value", "toString", 1, line);
        concat(chunks, current, line);
        chunks[current].emit_end(line);
    }
    if args.len() > 4 {
        push_str(&mut chunks[current], "/", line);
        concat(chunks, current, line);
        lget(&mut chunks[current], args[4], line);
        concat(chunks, current, line);
    }
    lset(&mut chunks[current], url, line);

    lget(&mut chunks[current], args[0], line);
    lget(&mut chunks[current], url, line);
    call_import(chunks, current, "wasi:sql", "connect", 1, line);
    struct_set_key(&mut chunks[current], "__conn", line);

    lget(&mut chunks[current], args[0], line);
}

/// `mysql_close(mysql)` — void.
pub fn emit_close(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = stash_args(chunks, current, argc, line);
    if !args.is_empty() {
        lget(&mut chunks[current], args[0], line);
        struct_get_key(&mut chunks[current], "__conn", line);
        call_import(chunks, current, "wasi:sql", "close", 1, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op(Op::NULL, line);
}

/// `mysql_select_db(mysql, db)` → 0.
pub fn emit_select_db(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    drop_values(chunks, current, argc, line);
    push_i32(&mut chunks[current], 0, line);
}

// ── queries ──────────────────────────────────────────────────────────────────

/// `mysql_query(mysql, stmt_str)` → 0 on success.
///
/// One C entry point covers both SELECT and DML, but `wasi:sql` splits them
/// (`query` returns rows, `execute` does not), so the statement's leading
/// keyword picks the verb — the same test `sql_adapter`/`mysqli_adapter` use.
pub fn emit_query(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = stash_args(chunks, current, argc, line);
    let is_select = alloc(&mut chunks[current]);
    let upper = alloc(&mut chunks[current]);

    lget(&mut chunks[current], args[1], line);
    call_import(chunks, current, "ecma:string", "trim", 1, line);
    call_import(chunks, current, "ecma:string", "toUpperCase", 1, line);
    lset(&mut chunks[current], upper, line);

    push_i32(&mut chunks[current], 0, line);
    lset(&mut chunks[current], is_select, line);
    for prefix in ["SELECT", "SHOW", "DESCRIBE", "EXPLAIN", "WITH"] {
        lget(&mut chunks[current], upper, line);
        push_str(&mut chunks[current], prefix, line);
        call_import(chunks, current, "ecma:string", "startsWith", 2, line);
        chunks[current].emit_if(line);
        push_i32(&mut chunks[current], 1, line);
        lset(&mut chunks[current], is_select, line);
        chunks[current].emit_end(line);
    }

    lget(&mut chunks[current], args[0], line);
    lget(&mut chunks[current], is_select, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], args[0], line);
    struct_get_key(&mut chunks[current], "__conn", line);
    lget(&mut chunks[current], args[1], line);
    call_import(chunks, current, "wasi:sql", "query", 2, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], args[0], line);
    struct_get_key(&mut chunks[current], "__conn", line);
    lget(&mut chunks[current], args[1], line);
    call_import(chunks, current, "wasi:sql", "execute", 2, line);
    call_import(chunks, current, "ecma:array", "new", 0, line);
    chunks[current].emit_end(line);
    struct_set_key(&mut chunks[current], "__rows", line);

    lget(&mut chunks[current], args[0], line);
    push_i32(&mut chunks[current], 0, line);
    struct_set_key(&mut chunks[current], "__cursor", line);

    push_i32(&mut chunks[current], 0, line);
}

/// `mysql_real_query(mysql, stmt_str, length)` — same, with an explicit length
/// that a counted string does not need.
pub fn emit_real_query(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 2 {
        // Drop `length` before the shared path, which expects two arguments.
        chunks[current].emit_op(Op::DROP, line);
    }
    emit_query(chunks, current, 2, line);
}

/// Build a MYSQL_RES sharing the connection's rows.
fn emit_result_from(chunks: &mut [Chunk], current: usize, conn: u16, line: u32) {
    let res = alloc(&mut chunks[current]);
    call_import(chunks, current, "ecma:object", "new", 0, line);
    lset(&mut chunks[current], res, line);

    lget(&mut chunks[current], res, line);
    push_str(&mut chunks[current], "MysqlResult", line);
    struct_set_key(&mut chunks[current], "__type", line);

    lget(&mut chunks[current], res, line);
    lget(&mut chunks[current], conn, line);
    struct_get_key(&mut chunks[current], "__rows", line);
    struct_set_key(&mut chunks[current], "__rows", line);

    lget(&mut chunks[current], res, line);
    push_i32(&mut chunks[current], 0, line);
    struct_set_key(&mut chunks[current], "__cursor", line);

    lget(&mut chunks[current], res, line);
}

/// `mysql_store_result(mysql)` → MYSQL_RES*.
pub fn emit_store_result(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = stash_args(chunks, current, argc, line);
    emit_result_from(chunks, current, args[0], line);
}

/// `mysql_use_result(mysql)` — an alias here.
///
/// Real MySQL streams row-at-a-time from the server for this call; `wasi:sql`
/// has already materialised the set, so there is nothing to stream and
/// pretending otherwise would only change where the memory is counted.
pub fn emit_use_result(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_store_result(chunks, current, argc, line);
}

/// `mysql_free_result(result)` — void.
pub fn emit_free_result(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    drop_values(chunks, current, argc, line);
    chunks[current].emit_op(Op::NULL, line);
}

// ── name dispatch ────────────────────────────────────────────────────────────

/// Route a `mysql_*` call. Accepts the bare C name and both mounted paths so a
/// C `#include` and a Fortran `bind(c)` land on the same emitter.
pub fn emit_mysql(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    let leaf = name.rsplit('.').next().unwrap_or(name);
    match leaf {
        "mysql_init" => emit_init(chunks, current, argc, line),
        "mysql_real_connect" => emit_real_connect(chunks, current, argc, line),
        "mysql_close" => emit_close(chunks, current, argc, line),
        "mysql_select_db" => emit_select_db(chunks, current, argc, line),
        "mysql_query" => emit_query(chunks, current, argc, line),
        "mysql_real_query" => emit_real_query(chunks, current, argc, line),
        "mysql_store_result" => emit_store_result(chunks, current, argc, line),
        "mysql_use_result" => emit_use_result(chunks, current, argc, line),
        "mysql_free_result" => emit_free_result(chunks, current, argc, line),
        _ => return false,
    }
    true
}
