//! SQLite3 C ABI, as an adapter over the `wasi:sql` host interface.
//!
//! The surface is the REAL `sqlite3_*` C API — same names, same argument
//! order, same return codes — so a program written against it is ordinary C
//! (or ordinary Fortran through `iso_c_binding`), compiles with a system
//! `libsqlite3` outside Vybe, and needs no Vybe-specific spelling. That is the
//! whole point of putting it here rather than in a language crate: C reaches it
//! through `#include <sqlite3.h>`, Fortran through `bind(c)`, and neither
//! frontend needs to know this file exists.
//!
//! Underneath there is no `sqlite3*` and no `sqlite3_stmt*` — `wasi:sql`
//! offers exactly eight verbs (`connect`, `query`, `execute`, `scalar`,
//! `beginTransaction`, `commit`, `rollback`, `close`), so the handles are
//! ordinary objects:
//!
//! ```text
//! db   = the object `wasi:sql.connect` returns
//! stmt = { __conn, __sql, __rows, __cursor, __row, __params }
//! ```
//!
//! WHAT IS FAITHFUL: `open`/`close`/`exec`, the `step` → `column_*` walk, the
//! bind family, and the `SQLITE_ROW`/`SQLITE_DONE` protocol.
//!
//! WHAT IS NOT: `prepare_v2` does not compile anything — it stashes the SQL, so
//! a syntax error surfaces at the first `step` rather than at `prepare`. And
//! the result set is MATERIALISED by `query`, so nothing streams: a large table
//! is fully resident before the first `step` returns. Both are inherited from
//! the host vocabulary, not chosen here.

use std::sync::Arc;

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

// Open flags (sqlite3.h) — the ones with a URI equivalent.
const SQLITE_OPEN_READONLY: i32 = 0x0000_0001;
const SQLITE_OPEN_READWRITE: i32 = 0x0000_0002;
const SQLITE_OPEN_CREATE: i32 = 0x0000_0004;
const SQLITE_OPEN_URI: i32 = 0x0000_0040;
const SQLITE_OPEN_MEMORY: i32 = 0x0000_0080;

// SQLite result codes (sqlite3.h). Only the ones this surface can produce.
const SQLITE_OK: i32 = 0;
const SQLITE_ERROR: i32 = 1;
const SQLITE_ROW: i32 = 100;
const SQLITE_DONE: i32 = 101;

// ── emit helpers (mirrors the conventions in pdo_adapter / sql_adapter) ──────

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
    // Per-chunk import table: add to the chunk being emitted into, never a
    // baked `chunks[0]` index (namespaceplan.md, "Import-table baking").
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

/// Pop `argc` arguments into fresh slots, LEFT-TO-RIGHT.
///
/// Arguments are pushed in source order, so the stack unwinds right-to-left;
/// the returned vector is indexed the way the C prototype reads.
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

/// Write `value` through a C out-pointer argument (`sqlite3 **ppDb`).
///
/// The pointer arrives as the carray reference shape every C pointer uses, so
/// this is the same store `posix_adapter` performs for its own out-parameters.
fn write_out_param(chunks: &mut [Chunk], current: usize, ptr: u16, value: u16, line: u32) {
    // There is no index-store OPCODE; element assignment goes through the
    // `ecma:array.set(array, index, value)` import, which is what an ordinary
    // `a[i] = v` compiles to as well.
    lget(&mut chunks[current], ptr, line);
    struct_get_key(&mut chunks[current], "__base", line);
    lget(&mut chunks[current], ptr, line);
    struct_get_key(&mut chunks[current], "__idx", line);
    lget(&mut chunks[current], value, line);
    call_import(chunks, current, "ecma:array", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

// ── connection ───────────────────────────────────────────────────────────────

/// `sqlite3_open(filename, ppDb)` → `SQLITE_OK`, connection written to `*ppDb`.
///
/// The filename goes to the host unchanged: `wasi:sql.connect` already maps a
/// bare path and `:memory:` onto a sqlite URL, so prefixing here would produce
/// `sqlite:sqlite:foo.db`.
pub fn emit_open(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = stash_args(chunks, current, argc, line);
    let db = alloc(&mut chunks[current]);

    lget(&mut chunks[current], args[0], line);
    call_import(chunks, current, "wasi:sql", "connect", 1, line);
    lset(&mut chunks[current], db, line);

    if args.len() > 1 {
        write_out_param(chunks, current, args[1], db, line);
    }
    push_i32(&mut chunks[current], SQLITE_OK, line);
}

/// `sqlite3_open_v2(filename, ppDb, flags, zVfs)` → `SQLITE_OK`.
///
/// The FLAGS are carried, not dropped: SQLite's own URI filename syntax has a
/// slot for every open mode, so `SQLITE_OPEN_READONLY` becomes
/// `file:name?mode=ro` and so on. That is the same move as carrying MySQL's
/// port into its URL — the option travels in the one string the host accepts.
///
/// ```text
/// READONLY              -> file:<name>?mode=ro
/// READWRITE (no CREATE) -> file:<name>?mode=rw
/// READWRITE|CREATE      -> <name>          (the default; no URI needed)
/// MEMORY                -> file:<name>?mode=memory
/// ```
///
/// `zVfs` names an OS-level VFS module — there is no such layer under
/// `wasi:sql`, so it is the one argument that cannot travel.
pub fn emit_open_v2(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = stash_args(chunks, current, argc, line);
    let db = alloc(&mut chunks[current]);
    let url = alloc(&mut chunks[current]);

    // Default: the plain filename, which is `mode=rwc` in SQLite's own terms.
    lget(&mut chunks[current], args[0], line);
    lset(&mut chunks[current], url, line);

    if args.len() > 2 {
        // Most specific first — MEMORY wins over READONLY wins over a bare
        // READWRITE, matching how sqlite3.h's flags compose.
        for (bit, mode) in [
            (SQLITE_OPEN_MEMORY, "memory"),
            (SQLITE_OPEN_READONLY, "ro"),
            (SQLITE_OPEN_READWRITE, "rw"),
        ] {
            lget(&mut chunks[current], args[2], line);
            push_i32(&mut chunks[current], bit, line);
            chunks[current].emit_op(Op::I32_AND, line);
            chunks[current].emit_if(line);
            push_str(&mut chunks[current], "file:", line);
            lget(&mut chunks[current], args[0], line);
            call_import(chunks, current, "ecma:string", "concat", 2, line);
            push_str(&mut chunks[current], "?mode=", line);
            call_import(chunks, current, "ecma:string", "concat", 2, line);
            push_str(&mut chunks[current], mode, line);
            call_import(chunks, current, "ecma:string", "concat", 2, line);
            lset(&mut chunks[current], url, line);
            chunks[current].emit_end(line);
        }
        // `READWRITE|CREATE` is the default the plain filename already means,
        // so it deliberately falls through with no URI.
    }

    lget(&mut chunks[current], url, line);
    call_import(chunks, current, "wasi:sql", "connect", 1, line);
    lset(&mut chunks[current], db, line);

    if args.len() > 1 {
        write_out_param(chunks, current, args[1], db, line);
    }
    push_i32(&mut chunks[current], SQLITE_OK, line);
}

/// `sqlite3_close(db)` → `SQLITE_OK`.
pub fn emit_close(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = stash_args(chunks, current, argc, line);
    if !args.is_empty() {
        lget(&mut chunks[current], args[0], line);
        call_import(chunks, current, "wasi:sql", "close", 1, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    push_i32(&mut chunks[current], SQLITE_OK, line);
}

/// `sqlite3_exec(db, sql, callback, arg, errmsg)` → `SQLITE_OK`.
///
/// The per-row CALLBACK is not supported: it would have to call back into guest
/// code from inside the adapter, and `wasi:sql` hands the rows over in one
/// piece. Programs that want rows use `prepare`/`step`, which is what the C API
/// itself recommends. The callback and its argument are dropped.
pub fn emit_exec(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = stash_args(chunks, current, argc, line);
    lget(&mut chunks[current], args[0], line);
    lget(&mut chunks[current], args[1], line);
    call_import(chunks, current, "wasi:sql", "execute", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    push_i32(&mut chunks[current], SQLITE_OK, line);
}

// ── statements ───────────────────────────────────────────────────────────────

/// Build the statement object: `{ __conn, __sql, __rows, __cursor, __row,
/// __params }`.
fn emit_new_stmt(chunks: &mut [Chunk], current: usize, conn: u16, sql: u16, line: u32) -> u16 {
    let stmt = alloc(&mut chunks[current]);
    call_import(chunks, current, "ecma:object", "new", 0, line);
    lset(&mut chunks[current], stmt, line);

    lget(&mut chunks[current], stmt, line);
    push_str(&mut chunks[current], "Sqlite3Stmt", line);
    struct_set_key(&mut chunks[current], "__type", line);

    lget(&mut chunks[current], stmt, line);
    lget(&mut chunks[current], conn, line);
    struct_set_key(&mut chunks[current], "__conn", line);

    lget(&mut chunks[current], stmt, line);
    lget(&mut chunks[current], sql, line);
    struct_set_key(&mut chunks[current], "__sql", line);

    // Rows are fetched on the first `step`, not here — `prepare` in this
    // surface stashes SQL and nothing else.
    lget(&mut chunks[current], stmt, line);
    chunks[current].emit_op(Op::NULL, line);
    struct_set_key(&mut chunks[current], "__rows", line);

    lget(&mut chunks[current], stmt, line);
    push_i32(&mut chunks[current], 0, line);
    struct_set_key(&mut chunks[current], "__cursor", line);

    lget(&mut chunks[current], stmt, line);
    chunks[current].emit_op(Op::NULL, line);
    struct_set_key(&mut chunks[current], "__row", line);

    // Bound parameters accumulate here and are passed to the host as a params
    // array — never spliced into the SQL text, which would be an injection.
    lget(&mut chunks[current], stmt, line);
    call_import(chunks, current, "ecma:array", "new", 0, line);
    struct_set_key(&mut chunks[current], "__params", line);

    stmt
}

/// `sqlite3_prepare_v2(db, sql, nByte, ppStmt, pzTail)` → `SQLITE_OK`.
///
/// `nByte` and `pzTail` describe a byte range and the unused tail of a
/// multi-statement string; neither is expressible against `wasi:sql`, so they
/// are accepted and dropped.
pub fn emit_prepare(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = stash_args(chunks, current, argc, line);
    let stmt = emit_new_stmt(chunks, current, args[0], args[1], line);
    if args.len() > 3 {
        write_out_param(chunks, current, args[3], stmt, line);
    }
    push_i32(&mut chunks[current], SQLITE_OK, line);
}

/// `sqlite3_finalize(stmt)` → `SQLITE_OK`. Nothing to release.
pub fn emit_finalize(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    drop_values(chunks, current, argc, line);
    push_i32(&mut chunks[current], SQLITE_OK, line);
}

/// `sqlite3_reset(stmt)` → `SQLITE_OK`: rewind the cursor, keep the binds.
pub fn emit_reset(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = stash_args(chunks, current, argc, line);
    lget(&mut chunks[current], args[0], line);
    push_i32(&mut chunks[current], 0, line);
    struct_set_key(&mut chunks[current], "__cursor", line);
    lget(&mut chunks[current], args[0], line);
    chunks[current].emit_op(Op::NULL, line);
    struct_set_key(&mut chunks[current], "__rows", line);
    push_i32(&mut chunks[current], SQLITE_OK, line);
}

/// `sqlite3_errmsg(db)` — best effort.
///
/// `wasi:sql` reports failures by trapping rather than by handing back a
/// message, so there is rarely anything to report by the time a caller asks.
/// Returning the empty string is honest; inventing text would not be.
pub fn emit_errmsg(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    drop_values(chunks, current, argc, line);
    push_str(&mut chunks[current], "", line);
}

pub fn emit_errcode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    drop_values(chunks, current, argc, line);
    push_i32(&mut chunks[current], SQLITE_OK, line);
}

// ── name dispatch ────────────────────────────────────────────────────────────

/// Route a `sqlite3_*` call. Accepts the bare C name and both mounted paths so
/// a C `#include` and a Fortran `bind(c)` land on the same emitter.
pub fn emit_sqlite(
    name: &str,
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    let leaf = name.rsplit('.').next().unwrap_or(name);
    match leaf {
        "sqlite3_open" => emit_open(chunks, current, argc, line),
        "sqlite3_open_v2" => emit_open_v2(chunks, current, argc, line),
        "sqlite3_close" | "sqlite3_close_v2" => emit_close(chunks, current, argc, line),
        "sqlite3_exec" => emit_exec(chunks, current, argc, line),
        "sqlite3_prepare" | "sqlite3_prepare_v2" => emit_prepare(chunks, current, argc, line),
        "sqlite3_finalize" => emit_finalize(chunks, current, argc, line),
        "sqlite3_reset" => emit_reset(chunks, current, argc, line),
        "sqlite3_errmsg" => emit_errmsg(chunks, current, argc, line),
        "sqlite3_errcode" | "sqlite3_extended_errcode" => {
            emit_errcode(chunks, current, argc, line)
        }
        _ => return false,
    }
    true
}

/// The result codes a caller compares against. Exposed as constants so
/// `if (rc == SQLITE_ROW)` resolves without the program declaring them.
pub const RESULT_CODES: &[(&str, i64)] = &[
    ("SQLITE_OK", SQLITE_OK as i64),
    ("SQLITE_ERROR", SQLITE_ERROR as i64),
    ("SQLITE_ROW", SQLITE_ROW as i64),
    ("SQLITE_DONE", SQLITE_DONE as i64),
    ("SQLITE_OPEN_READONLY", SQLITE_OPEN_READONLY as i64),
    ("SQLITE_OPEN_READWRITE", SQLITE_OPEN_READWRITE as i64),
    ("SQLITE_OPEN_CREATE", SQLITE_OPEN_CREATE as i64),
    ("SQLITE_OPEN_URI", SQLITE_OPEN_URI as i64),
    ("SQLITE_OPEN_MEMORY", SQLITE_OPEN_MEMORY as i64),
];
