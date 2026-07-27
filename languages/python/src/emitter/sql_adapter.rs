//! Python `sqlite3` surface, built as an ADAPTER over the existing `wasi:sql`
//! host interface — the same one PHP's `pdo_adapter` and .NET drive. NO host
//! changes: we only CALL `wasi:sql.{connect,query,execute,scalar,commit,
//! rollback,close,beginTransaction}`.
//!
//! Shapes:
//!   Connection — the object `wasi:sql.connect` returns (carries `__conn_id`);
//!                stamped `__type = "SqlConnection"` by the host.
//!   Cursor     — a plain object with `__conn` (the connection), `__rows`
//!                (last query's rows), `__cursor` (fetch index), plus the plain
//!                `lastrowid` / `rowcount` attributes Python reads directly.
//!
//! Rows come back from `wasi:sql` as column-keyed objects carrying a
//! `__col_names` array; `fetchone`/`fetchall` turn each into a **tagged tuple**
//! of values in select-column order (so `r[0]` indexes and the list reprs as
//! `[('Bob', 25), ...]`).
//!
//! The walker (see `python_sql_*`) tracks sqlite connection/cursor variables and
//! rewrites their methods to the collision-free `__sql_*` builtins routed here,
//! so generic `.close()`/`.execute()` on files/sockets are never intercepted.

use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use vybe_compiler::compiler::{collections, tuples};

// ── local emit helpers (mirror pdo_adapter conventions) ──────────────────────

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

fn call_import(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(argc, line);
}

fn struct_get_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_GET, idx, line);
}

/// `obj[key] = <value on stack>`. Stack: `[obj, value] -> []`.
fn struct_set_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_SET, idx, line);
    chunk.emit_op(Op::DROP, line);
}

/// Stash `argc` call arguments into consecutive scratch slots, arg0 at `base`.
fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

fn empty_array(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_array_new(chunks, current, 0, line);
}

// ── row → tagged tuple ───────────────────────────────────────────────────────

/// Build a tagged tuple of the row's column VALUES in select order and leave it
/// on the stack. Reads `row.__col_names` and dynamic-gets each column by name.
fn emit_row_to_tuple(chunks: &mut [Chunk], current: usize, row_slot: u16, line: u32) {
    let names = alloc(&mut chunks[current]);
    let tup = alloc(&mut chunks[current]);
    let n = alloc(&mut chunks[current]);
    let i = alloc(&mut chunks[current]);
    let name = alloc(&mut chunks[current]);
    let val = alloc(&mut chunks[current]);

    // names = row.__col_names
    lget(&mut chunks[current], row_slot, line);
    struct_get_key(&mut chunks[current], "__col_names", line);
    lset(&mut chunks[current], names, line);

    // tup = []
    empty_array(chunks, current, line);
    lset(&mut chunks[current], tup, line);

    // n = names.length ; i = 0
    lget(&mut chunks[current], names, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    lset(&mut chunks[current], n, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], i, line);

    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    lget(&mut chunks[current], i, line);
    lget(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    // name = names[i]
    lget(&mut chunks[current], names, line);
    lget(&mut chunks[current], i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(&mut chunks[current], name, line);

    // val = ecma:object.get(row, name)
    lget(&mut chunks[current], row_slot, line);
    lget(&mut chunks[current], name, line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    lset(&mut chunks[current], val, line);

    // tup.push(val)  (push leaves the new length — drop it)
    lget(&mut chunks[current], tup, line);
    lget(&mut chunks[current], val, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // i += 1
    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    // tag → tuple, leave on stack
    lget(&mut chunks[current], tup, line);
    tuples::emit_tag(chunks, current, line);
}

/// Compute `use_raw = truthy(cursor.__conn.row_factory)` into `flag_slot` — set
/// when the connection's `row_factory` is `sqlite3.Row`, so fetch returns the
/// raw column-keyed row (named + positional access) instead of a tuple.
fn emit_row_factory_flag(chunks: &mut [Chunk], current: usize, cursor: u16, flag_slot: u16, line: u32) {
    lget(&mut chunks[current], cursor, line);
    struct_get_key(&mut chunks[current], "__conn", line);
    struct_get_key(&mut chunks[current], "row_factory", line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    lset(&mut chunks[current], flag_slot, line);
}

/// Leave either the raw row (when `raw_flag` is set) or a tagged tuple on stack.
fn emit_row_result(chunks: &mut [Chunk], current: usize, row_slot: u16, raw_flag: u16, line: u32) {
    lget(&mut chunks[current], raw_flag, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], row_slot, line);
    chunks[current].emit_else(line);
    emit_row_to_tuple(chunks, current, row_slot, line);
    chunks[current].emit_end(line);
}

/// Push i32 `1` into `flag_slot` if the SQL in `sql_slot` is row-returning
/// (`SELECT`/`PRAGMA`/`WITH`/`EXPLAIN`/`VALUES`), else `0`.
fn emit_is_query_flag(chunks: &mut [Chunk], current: usize, sql_slot: u16, flag_slot: u16, line: u32) {
    let upper = alloc(&mut chunks[current]);
    // upper = sql.trim().toUpperCase()
    lget(&mut chunks[current], sql_slot, line);
    call_import(chunks, current, "ecma:string", "trim", 1, line);
    call_import(chunks, current, "ecma:string", "toUpperCase", 1, line);
    lset(&mut chunks[current], upper, line);

    // flag = 0
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], flag_slot, line);

    for prefix in ["SELECT", "PRAGMA", "WITH", "EXPLAIN", "VALUES"] {
        lget(&mut chunks[current], upper, line);
        push_str(&mut chunks[current], prefix, line);
        call_import(chunks, current, "ecma:string", "startsWith", 2, line);
        chunks[current].emit_if(line);
        chunks[current].emit_i32_const(1, line);
        lset(&mut chunks[current], flag_slot, line);
        chunks[current].emit_end(line);
    }
}

// ── builtins ─────────────────────────────────────────────────────────────────

/// `__sql_connect(path)` → connection. `:memory:` and bare paths are normalized
/// to a sqlite URL by the host's `wasi:sql.connect`.
pub fn emit_connect(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let conn = alloc(&mut chunks[current]);
    lget(&mut chunks[current], base, line);
    call_import(chunks, current, "wasi:sql", "connect", 1, line);
    lset(&mut chunks[current], conn, line);
    // Self-reference so `<conn>.__conn` resolves whether the receiver is a
    // Connection (shortcut `conn.execute(...)`) or a Cursor (`cur.__conn`).
    lget(&mut chunks[current], conn, line);
    lget(&mut chunks[current], conn, line);
    struct_set_key(&mut chunks[current], "__conn", line);
    // Python `Connection.isolation_level` defaults to "" (deferred BEGIN).
    lget(&mut chunks[current], conn, line);
    push_str(&mut chunks[current], "", line);
    struct_set_key(&mut chunks[current], "isolation_level", line);
    lget(&mut chunks[current], conn, line);
    // result: connection object
}

/// `__sql_cursor(conn)` → a fresh cursor bound to `conn`.
pub fn emit_cursor(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let cur = alloc(&mut chunks[current]);

    call_import(chunks, current, "ecma:object", "new", 0, line);
    lset(&mut chunks[current], cur, line);

    lget(&mut chunks[current], cur, line);
    push_str(&mut chunks[current], "SqlCursor", line);
    struct_set_key(&mut chunks[current], "__type", line);

    lget(&mut chunks[current], cur, line);
    lget(&mut chunks[current], base, line);
    struct_set_key(&mut chunks[current], "__conn", line);

    lget(&mut chunks[current], cur, line);
    empty_array(chunks, current, line);
    struct_set_key(&mut chunks[current], "__rows", line);

    lget(&mut chunks[current], cur, line);
    chunks[current].emit_i32_const(0, line);
    struct_set_key(&mut chunks[current], "__cursor", line);

    lget(&mut chunks[current], cur, line);
    chunks[current].emit_op(Op::NULL, line);
    struct_set_key(&mut chunks[current], "lastrowid", line);

    lget(&mut chunks[current], cur, line);
    chunks[current].emit_i32_const(-1, line);
    struct_set_key(&mut chunks[current], "rowcount", line);

    lget(&mut chunks[current], cur, line);
    // result: cursor
}

/// `__sql_execute(cursor, sql[, params])` → the cursor (Python returns it).
pub fn emit_execute(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let cursor = base;
    let sql = base + 1;
    let has_params = argc >= 3;
    let params = base + 2;

    let conn = alloc(&mut chunks[current]);
    lget(&mut chunks[current], cursor, line);
    struct_get_key(&mut chunks[current], "__conn", line);
    lset(&mut chunks[current], conn, line);

    let flag = alloc(&mut chunks[current]);
    emit_is_query_flag(chunks, current, sql, flag, line);

    lget(&mut chunks[current], flag, line);
    chunks[current].emit_if(line);
    {
        // query branch → store rows
        let rows = alloc(&mut chunks[current]);
        lget(&mut chunks[current], conn, line);
        lget(&mut chunks[current], sql, line);
        if has_params {
            lget(&mut chunks[current], params, line);
            call_import(chunks, current, "wasi:sql", "query", 3, line);
        } else {
            call_import(chunks, current, "wasi:sql", "query", 2, line);
        }
        lset(&mut chunks[current], rows, line);

        lget(&mut chunks[current], cursor, line);
        lget(&mut chunks[current], rows, line);
        struct_set_key(&mut chunks[current], "__rows", line);
        lget(&mut chunks[current], cursor, line);
        chunks[current].emit_i32_const(0, line);
        struct_set_key(&mut chunks[current], "__cursor", line);
    }
    chunks[current].emit_else(line);
    {
        // exec branch → affected count + lastrowid, reset rows
        let count = alloc(&mut chunks[current]);
        lget(&mut chunks[current], conn, line);
        lget(&mut chunks[current], sql, line);
        if has_params {
            lget(&mut chunks[current], params, line);
            call_import(chunks, current, "wasi:sql", "execute", 3, line);
        } else {
            call_import(chunks, current, "wasi:sql", "execute", 2, line);
        }
        lset(&mut chunks[current], count, line);

        lget(&mut chunks[current], cursor, line);
        lget(&mut chunks[current], count, line);
        struct_set_key(&mut chunks[current], "rowcount", line);

        lget(&mut chunks[current], cursor, line);
        empty_array(chunks, current, line);
        struct_set_key(&mut chunks[current], "__rows", line);
        lget(&mut chunks[current], cursor, line);
        chunks[current].emit_i32_const(0, line);
        struct_set_key(&mut chunks[current], "__cursor", line);

        // lastrowid = scalar("SELECT last_insert_rowid()")
        let lastid = alloc(&mut chunks[current]);
        lget(&mut chunks[current], conn, line);
        push_str(&mut chunks[current], "SELECT last_insert_rowid()", line);
        call_import(chunks, current, "wasi:sql", "scalar", 2, line);
        lset(&mut chunks[current], lastid, line);
        lget(&mut chunks[current], cursor, line);
        lget(&mut chunks[current], lastid, line);
        struct_set_key(&mut chunks[current], "lastrowid", line);
    }
    chunks[current].emit_end(line);

    lget(&mut chunks[current], cursor, line);
    // result: cursor
}

/// `__sql_executemany(cursor, sql, seq)` → the cursor. Runs `sql` once per
/// parameter tuple in `seq`.
pub fn emit_executemany(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let cursor = base;
    let sql = base + 1;
    let seq = base + 2;

    let conn = alloc(&mut chunks[current]);
    lget(&mut chunks[current], cursor, line);
    struct_get_key(&mut chunks[current], "__conn", line);
    lset(&mut chunks[current], conn, line);

    let n = alloc(&mut chunks[current]);
    let i = alloc(&mut chunks[current]);
    let param = alloc(&mut chunks[current]);
    lget(&mut chunks[current], seq, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    lset(&mut chunks[current], n, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], i, line);

    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    lget(&mut chunks[current], i, line);
    lget(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    lget(&mut chunks[current], seq, line);
    lget(&mut chunks[current], i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(&mut chunks[current], param, line);

    lget(&mut chunks[current], conn, line);
    lget(&mut chunks[current], sql, line);
    lget(&mut chunks[current], param, line);
    call_import(chunks, current, "wasi:sql", "execute", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    lget(&mut chunks[current], cursor, line);
    // result: cursor
}

/// `__sql_fetchall(cursor)` → list of tagged tuples from the current position.
pub fn emit_fetchall(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let cursor = base;

    let rows = alloc(&mut chunks[current]);
    let res = alloc(&mut chunks[current]);
    let n = alloc(&mut chunks[current]);
    let i = alloc(&mut chunks[current]);
    let row = alloc(&mut chunks[current]);
    let tup = alloc(&mut chunks[current]);
    let raw = alloc(&mut chunks[current]);
    emit_row_factory_flag(chunks, current, cursor, raw, line);

    lget(&mut chunks[current], cursor, line);
    struct_get_key(&mut chunks[current], "__rows", line);
    lset(&mut chunks[current], rows, line);

    empty_array(chunks, current, line);
    lset(&mut chunks[current], res, line);

    lget(&mut chunks[current], rows, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    lset(&mut chunks[current], n, line);

    // i = cursor.__cursor
    lget(&mut chunks[current], cursor, line);
    struct_get_key(&mut chunks[current], "__cursor", line);
    lset(&mut chunks[current], i, line);

    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    lget(&mut chunks[current], i, line);
    lget(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    lget(&mut chunks[current], rows, line);
    lget(&mut chunks[current], i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(&mut chunks[current], row, line);

    emit_row_result(chunks, current, row, raw, line);
    lset(&mut chunks[current], tup, line);

    lget(&mut chunks[current], res, line);
    lget(&mut chunks[current], tup, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    // cursor.__cursor = n (exhausted)
    lget(&mut chunks[current], cursor, line);
    lget(&mut chunks[current], n, line);
    struct_set_key(&mut chunks[current], "__cursor", line);

    lget(&mut chunks[current], res, line);
    // result: list of tuples
}

/// `__sql_fetchone(cursor)` → next row as a tagged tuple, or `None`.
pub fn emit_fetchone(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let cursor = base;

    let rows = alloc(&mut chunks[current]);
    let n = alloc(&mut chunks[current]);
    let cur = alloc(&mut chunks[current]);
    let row = alloc(&mut chunks[current]);
    let raw = alloc(&mut chunks[current]);
    emit_row_factory_flag(chunks, current, cursor, raw, line);

    lget(&mut chunks[current], cursor, line);
    struct_get_key(&mut chunks[current], "__rows", line);
    lset(&mut chunks[current], rows, line);
    lget(&mut chunks[current], rows, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    lset(&mut chunks[current], n, line);
    lget(&mut chunks[current], cursor, line);
    struct_get_key(&mut chunks[current], "__cursor", line);
    lset(&mut chunks[current], cur, line);

    // if cur < n
    lget(&mut chunks[current], cur, line);
    lget(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    {
        lget(&mut chunks[current], rows, line);
        lget(&mut chunks[current], cur, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        lset(&mut chunks[current], row, line);

        // cursor.__cursor = cur + 1
        lget(&mut chunks[current], cursor, line);
        lget(&mut chunks[current], cur, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        struct_set_key(&mut chunks[current], "__cursor", line);

        emit_row_result(chunks, current, row, raw, line);
    }
    chunks[current].emit_else(line);
    {
        chunks[current].emit_op(Op::NULL, line);
    }
    chunks[current].emit_end(line);
    // result: tuple or null
}

fn emit_conn_op(chunks: &mut [Chunk], current: usize, argc: u8, op: &str, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    lget(&mut chunks[current], base, line);
    call_import(chunks, current, "wasi:sql", op, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
    // result: None
}

/// `__sql_commit(conn)` → None.
pub fn emit_commit(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_conn_op(chunks, current, argc, "commit", line);
}

/// `__sql_rollback(conn)` → None.
pub fn emit_rollback(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_conn_op(chunks, current, argc, "rollback", line);
}

/// `__sql_close(conn)` → None.
pub fn emit_close(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_conn_op(chunks, current, argc, "close", line);
}

/// `__sql_begin(conn)` → None. Used by the `with conn:` transaction desugar.
pub fn emit_begin(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_conn_op(chunks, current, argc, "beginTransaction", line);
}
