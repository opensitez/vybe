//! .NET ADO surface (`System.Data.SqlClient` / `System.Data.OleDb` / ADODB) —
//! bytecode adapters.
//!
//! Every member here used to be a host function under `wasi:sql/types`. None of
//! them performs I/O: a `SqlDataReader` is `{ __rows, __col_names, __pos }` and
//! its entire API is a cursor over two arrays; a `SqlParameterCollection` is an
//! array push; `Commit`/`Rollback` are a statement and an exec. Those are
//! object, array and string operations, so they belong in emitted bytecode
//! running inside the VM — not behind a `CALL` import.
//!
//! Only the four functions that genuinely cross the boundary stay host-side, and
//! they are the actual `wasi:sql` WIT (`proposals/wasi-sql/wit/`):
//!   - `wasi:sql/types`     `[static]connection.open`, `[static]statement.prepare`
//!   - `wasi:sql/readwrite` `query`, `exec`
//!
//! Both `readwrite` functions resolve their connection through `wasi_id`, which
//! reads `__wasi_id` or `__conn_id` off whatever object it is handed — so any
//! ADO object carrying a connection id (command, transaction, adapter) is a
//! valid `borrow<connection>` operand without a separate lookup.
//!
//! Pattern: `convert_adapter.rs` (loops), `datatable_adapter.rs` (structs).

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::{collections, convert, ops};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::object_fields::field_slot;

const COL_NAMES_KEY: &str = "__col_names";
const CONN_ID_KEY: &str = "__conn_id";
const IS_CLOSED_KEY: &str = "isclosed";
const ITEMS_KEY: &str = "__items";
const POS_KEY: &str = "__pos";
const ROWS_KEY: &str = "__rows";

const WASI_TYPES: &str = "wasi:sql/types";
const WASI_READWRITE: &str = "wasi:sql/readwrite";
const STATEMENT_PREPARE: &str = "[static]statement.prepare";

// ── Primitive helpers ─────────────────────────────────────────────────────────

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn call_import(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// `obj.key` — reads a property off the object held in `obj_slot`.
fn get_prop(chunk: &mut Chunk, obj_slot: u16, key: &str, line: u32) {
    lget(chunk, obj_slot, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        Dest::Stack,
        line,
    );
}

/// `obj.key = <local>` — consumes nothing from the stack.
fn set_prop_local(chunk: &mut Chunk, obj_slot: u16, key: &str, val_slot: u16, line: u32) {
    lget(chunk, obj_slot, line);
    lget(chunk, val_slot, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
}

/// `DUP → <value> → STRUCT_SET key`, leaving the object on the stack. The
/// struct-literal builder: see `datatable_adapter::set_field`.
fn set_field(chunk: &mut Chunk, key: &str, val_fn: impl FnOnce(&mut Chunk, u32), line: u32) {
    chunk.emit_dup(line);
    val_fn(chunk, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
}

/// `void` return — every ADO member that answers nothing pushes a null ref, the
/// VM's always-push-one convention.
fn push_void(chunk: &mut Chunk, line: u32) {
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `0 <= idx < len(arr)` → i32 0/1. Both operands stay in locals so the guard
/// can be reused by the `if_value` arms that follow it.
fn emit_index_in_range(
    chunks: &mut [Chunk],
    current: usize,
    arr_slot: u16,
    idx_slot: u16,
    line: u32,
) {
    lget(&mut chunks[current], idx_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GE_S, line);

    lget(&mut chunks[current], idx_slot, line);
    lget(&mut chunks[current], arr_slot, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);

    chunks[current].emit_op(Op::I32_AND, line);
}

/// A fresh `SqlParameterCollection`, left in a local so the caller can hang it
/// off a struct field (a `set_field` closure only gets `&mut Chunk`, and the
/// empty array needs the import table).
fn emit_params_collection(chunks: &mut [Chunk], current: usize, line: u32) -> u16 {
    collections::emit_array_new(chunks, current, 0, line);
    let chunk = &mut chunks[current];
    let items_slot = reserve_slot(chunk);
    lset(chunk, items_slot, line);
    let slot = reserve_slot(chunk);
    class_slots::emit_class_construct(
        chunk,
        "SqlParameterCollection",
        &[
            (field_slot(ITEMS_KEY), ValueSource::Local(items_slot)),
        ],
        line,
    );
    lset(chunk, slot, line);
    slot
}

// ── SqlConnection ─────────────────────────────────────────────────────────────

/// `new SqlConnection(connStr?)`.
/// Stack in: `[connStr]` (argc 1) or `[]` (argc 0)  Stack out: `[connection]`
///
/// `provider` is left empty here and written by `Open()` from the value
/// `[static]connection.open` reports, which is the authority — the host
/// constructor guessed it by re-normalising the connection string, duplicating
/// `normalize_conn_str_full`. Nothing reads the field before `Open()`.
pub fn emit_connection_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        chunk.emit_string_const("", line);
    }
    let raw_slot = reserve_slot(chunk);
    lset(chunk, raw_slot, line);

    class_slots::emit_class_construct(
        chunk,
        "SqlConnection",
        &[
            (field_slot("__conn_id"), ValueSource::ConstF64(0.0)),
            (field_slot("connectionstring"), ValueSource::Local(raw_slot)),
            (field_slot("provider"), ValueSource::ConstStr("".to_string())),
            (field_slot("serverversion"), ValueSource::ConstStr("".to_string())),
            (field_slot("connectiontimeout"), ValueSource::ConstF64(30.0)),
            (field_slot("state"), ValueSource::ConstStr("Closed".to_string())),
        ],
        line,
    );
}

/// `connection.Open()` — open over the real WIT, then copy what the returned
/// `connection` resource reports onto the receiver.
///
/// `[static]connection.open` answers either a connection object (carrying
/// `__wasi_id`, `provider`, `serverversion`) or an `error` resource, and the
/// error's message comes from `[method]error.trace`, which is also spec. A
/// failed open leaves the receiver `Closed` and writes the trace to stderr,
/// which is what the host member did with its `eprintln!`.
///
/// Stack in: `[connection]`  Stack out: `[null]`
pub fn emit_connection_open(chunks: &mut [Chunk], current: usize, line: u32) {
    let (conn_slot, cs_slot) = {
        let chunk = &mut chunks[current];
        let conn_slot = reserve_slot(chunk);
        let cs_slot = reserve_slot(chunk);
        lset(chunk, conn_slot, line);
        get_prop(chunk, conn_slot, "connectionstring", line);
        lset(chunk, cs_slot, line);

        // Already open, or nothing to open with — the host member returned
        // early on both.
        get_prop(chunk, conn_slot, CONN_ID_KEY, line);
        chunk.emit_f64_const(0.0, line);
        ops::emit_dyn_eq(chunk, line);
        lget(chunk, cs_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_op(Op::I32_AND, line);
        lget(chunk, cs_slot, line);
        chunk.emit_string_const("", line);
        ops::emit_dyn_ne(chunk, line);
        chunk.emit_op(Op::I32_AND, line);
        chunk.emit_if(line);
        (conn_slot, cs_slot)
    };

    lget(&mut chunks[current], cs_slot, line);
    call_import(chunks, current, WASI_TYPES, "[static]connection.open", 1, line);
    let res_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        lset(chunk, slot, line);
        // An `error` resource carries no `__wasi_id`; a `connection` does.
        get_prop(chunk, slot, "__wasi_id", line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);

        let id_slot = reserve_slot(chunk);
        get_prop(chunk, slot, "__wasi_id", line);
        lset(chunk, id_slot, line);
        set_prop_local(chunk, conn_slot, CONN_ID_KEY, id_slot, line);
        set_prop_local(chunk, conn_slot, "__wasi_id", id_slot, line);

        let scratch = reserve_slot(chunk);
        get_prop(chunk, slot, "provider", line);
        lset(chunk, scratch, line);
        set_prop_local(chunk, conn_slot, "provider", scratch, line);
        get_prop(chunk, slot, "serverversion", line);
        lset(chunk, scratch, line);
        set_prop_local(chunk, conn_slot, "serverversion", scratch, line);

        let state_idx = chunk.add_constant(Value::String(Arc::from("state")));
        lget(chunk, conn_slot, line);
        chunk.emit_string_const("Open", line);
        class_slots::emit_class_set(
            chunk,
            ObjSource::Stack,
            &field_slot("state"),
            ValueSource::Stack,
            line,
        );

        chunk.emit_else(line);
        chunk.emit_string_const("wasi:sql/types connection.open: ", line);
        lget(chunk, slot, line);
        slot
    };
    call_import(chunks, current, WASI_TYPES, "[method]error.trace", 1, line);
    {
        let chunk = &mut chunks[current];
        vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    }
    crate::emitter::core::console_adapter::emit_console_error_writeline(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line); // error / success
    chunk.emit_end(line); // guard
    let _ = res_slot;
    push_void(chunk, line);
}

/// `connection.Close()` — guest-side state, plus the one host call that is real
/// teardown.
///
/// `wasi:sql/types` declares `connection` as a Component Model **resource**, and
/// a resource is destroyed by the canonical `resource.drop`, which no WIT ever
/// declares — the ABI supplies it for every resource. `wasi:sql.close` is that
/// drop: one line, `state().conns.remove(id)`. PHP's `mysqli_close` adapter
/// already works exactly this way — null the guest handle, call the host to
/// release the real connection.
///
/// Stack in: `[connection]`  Stack out: `[null]`
pub fn emit_connection_close(chunks: &mut [Chunk], current: usize, line: u32) {
    let conn_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        lset(chunk, slot, line);
        lget(chunk, slot, line);
        slot
    };
    call_import(chunks, current, "wasi:sql", "close", 1, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);

    let zero_slot = reserve_slot(chunk);
    chunk.emit_f64_const(0.0, line);
    lset(chunk, zero_slot, line);
    set_prop_local(chunk, conn_slot, CONN_ID_KEY, zero_slot, line);
    set_prop_local(chunk, conn_slot, "__wasi_id", zero_slot, line);

    let state_idx = chunk.add_constant(Value::String(Arc::from("state")));
    lget(chunk, conn_slot, line);
    chunk.emit_string_const("Closed", line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot("state"),
        ValueSource::Stack,
        line,
    );

    push_void(chunk, line);
}

/// `connection.CreateCommand()`.
/// Stack in: `[connection]`  Stack out: `[command]`
pub fn emit_connection_create_command(chunks: &mut [Chunk], current: usize, line: u32) {
    let conn_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        lset(chunk, slot, line);
        slot
    };
    emit_command_struct(chunks, current, Some(conn_slot), None, line);
}

/// `connection.BeginTransaction()` — `BEGIN` over the real WIT, then a
/// transaction object carrying the connection id. `null` on a closed
/// connection, as the host member answered.
/// Stack in: `[connection]`  Stack out: `[transaction | null]`
pub fn emit_connection_begin_transaction(chunks: &mut [Chunk], current: usize, line: u32) {
    let (conn_slot, id_slot) = {
        let chunk = &mut chunks[current];
        let conn_slot = reserve_slot(chunk);
        let id_slot = reserve_slot(chunk);
        lset(chunk, conn_slot, line);
        get_prop(chunk, conn_slot, CONN_ID_KEY, line);
        lset(chunk, id_slot, line);
        lget(chunk, id_slot, line);
        chunk.emit_f64_const(0.0, line);
        ops::emit_dyn_ne(chunk, line);
        chunk.emit_if_value(line);
        (conn_slot, id_slot)
    };

    lget(&mut chunks[current], conn_slot, line);
    chunks[current].emit_string_const("BEGIN", line);
    collections::emit_array_new(chunks, current, 0, line);
    call_import(chunks, current, WASI_TYPES, STATEMENT_PREPARE, 2, line);
    call_import(chunks, current, WASI_READWRITE, "exec", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    let chunk = &mut chunks[current];
    class_slots::emit_class_construct(
        chunk,
        "SqlTransaction",
        &[
            (field_slot(CONN_ID_KEY), ValueSource::Local(id_slot)),
            (field_slot(IS_CLOSED_KEY), ValueSource::ConstBool(false)),
        ],
        line,
    );

    chunk.emit_else(line);
    push_void(chunk, line);
    chunk.emit_end(line);
}

// ── SqlCommand ────────────────────────────────────────────────────────────────

/// The `SqlCommand` struct. `conn_slot` supplies `__conn_id` /
/// `connectionstring`; `text_slot` supplies `commandtext`. Either may be absent,
/// which is how the host constructor read a missing argument.
fn emit_command_struct(
    chunks: &mut [Chunk],
    current: usize,
    conn_slot: Option<u16>,
    text_slot: Option<u16>,
    line: u32,
) {
    let params_slot = emit_params_collection(chunks, current, line);
    let chunk = &mut chunks[current];
    class_slots::emit_class_alloc(chunk, line);
    set_field(
        chunk,
        "__type",
        |c, l| c.emit_string_const("SqlCommand", l),
        line,
    );
    set_field(
        chunk,
        CONN_ID_KEY,
        |c, l| match conn_slot {
            Some(slot) => get_prop(c, slot, CONN_ID_KEY, l),
            None => c.emit_f64_const(0.0, l),
        },
        line,
    );
    set_field(
        chunk,
        "commandtext",
        |c, l| match text_slot {
            Some(slot) => lget(c, slot, l),
            None => c.emit_string_const("", l),
        },
        line,
    );
    set_field(
        chunk,
        "commandtimeout",
        |c, l| c.emit_f64_const(30.0, l),
        line,
    );
    set_field(chunk, "commandtype", |c, l| c.emit_f64_const(1.0, l), line);
    set_field(
        chunk,
        "connectionstring",
        |c, l| match conn_slot {
            Some(slot) => get_prop(c, slot, "connectionstring", l),
            None => c.emit_string_const("", l),
        },
        line,
    );
    set_field(chunk, "parameters", |c, l| lget(c, params_slot, l), line);
}

/// `new SqlCommand(sql?, connection?)`.
/// Stack in: `[sql?, connection?]`  Stack out: `[command]`
pub fn emit_command_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (conn_slot, text_slot) = {
        let chunk = &mut chunks[current];
        let conn_slot = if argc >= 2 {
            let slot = reserve_slot(chunk);
            lset(chunk, slot, line);
            Some(slot)
        } else {
            None
        };
        let text_slot = if argc >= 1 {
            let slot = reserve_slot(chunk);
            lset(chunk, slot, line);
            Some(slot)
        } else {
            None
        };
        // Any argument past the two the host read is discarded, as it was.
        for _ in 2..argc {
            chunk.emit_op(Op::DROP, line);
        }
        (conn_slot, text_slot)
    };
    emit_command_struct(chunks, current, conn_slot, text_slot, line);
}

/// `c` is an ASCII name character — `[0-9A-Za-z_]`, the host's
/// `is_ascii_alphanumeric() || '_'`. Deliberately NOT `strings::emit_is_alnum`,
/// which is Unicode Alphabetic and would accept characters the host's scan
/// stopped at. Leaves i32 0/1.
fn emit_is_name_char(chunk: &mut Chunk, code_slot: u16, line: u32) {
    let range = |c: &mut Chunk, lo: i32, hi: i32| {
        lget(c, code_slot, line);
        c.emit_i32_const(lo, line);
        c.emit_op(Op::I32_GE_S, line);
        lget(c, code_slot, line);
        c.emit_i32_const(hi, line);
        c.emit_op(Op::I32_LE_S, line);
        c.emit_op(Op::I32_AND, line);
    };
    range(chunk, b'0' as i32, b'9' as i32);
    range(chunk, b'A' as i32, b'Z' as i32);
    chunk.emit_op(Op::I32_OR, line);
    range(chunk, b'a' as i32, b'z' as i32);
    chunk.emit_op(Op::I32_OR, line);
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(b'_' as i32, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_op(Op::I32_OR, line);
}

/// `command_sql_and_params` — the command's SQL and its bound values, with
/// `@name` placeholders rewritten to `?` in the order they occur.
///
/// This is the one member of the family that is not a struct read: a character
/// scan that tracks single-quoted literals so `'a@b'` is left alone, and that
/// resolves each `@name` against the parameter collection. `lastIndexOf` is the
/// lookup because the host built a `HashMap`, where a repeated name kept the
/// LAST value.
///
/// When no placeholder resolved, the host passed the parameter values in
/// collection order against the untouched SQL — positional `?` binding. Kept.
///
/// Returns `(sql_slot, params_slot)`.
fn emit_command_sql_and_params(
    chunks: &mut [Chunk],
    current: usize,
    cmd_slot: u16,
    line: u32,
) -> (u16, u16) {
    // ── the command's text and parameter items ────────────────────────────
    let (sql0, pobj, items) = {
        let chunk = &mut chunks[current];
        let sql0 = reserve_slot(chunk);
        let pobj = reserve_slot(chunk);
        let items = reserve_slot(chunk);
        get_prop(chunk, cmd_slot, "commandtext", line);
        lset(chunk, sql0, line);
        get_prop(chunk, cmd_slot, "parameters", line);
        lset(chunk, pobj, line);
        lget(chunk, pobj, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if(line);
        (sql0, pobj, items)
    };
    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], items, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_else(line);
        get_prop(chunk, pobj, ITEMS_KEY, line);
        lset(chunk, items, line);
        lget(chunk, items, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if(line);
    }
    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], items, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_end(line);
        chunk.emit_end(line);
    }

    // ── split the items into name / value lookup arrays ───────────────────
    let (names, values, ordered, k, n) = {
        let chunk = &mut chunks[current];
        (
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
        )
    };
    for slot in [names, values, ordered] {
        collections::emit_array_new(chunks, current, 0, line);
        lset(&mut chunks[current], slot, line);
    }
    {
        let chunk = &mut chunks[current];
        chunk.emit_i32_const(0, line);
        lset(chunk, k, line);
        lget(chunk, items, line);
    }
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_FROM_F64, line);
    lset(&mut chunks[current], n, line);

    let split = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, k, line);
        lget(chunk, n, line);
        chunk.emit_op(Op::I32_LT_S, line);
    }
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);

    let (nm, vl) = {
        let chunk = &mut chunks[current];
        let nm = reserve_slot(chunk);
        let vl = reserve_slot(chunk);
        let name_idx = chunk.add_constant(Value::String(Arc::from("name")));
        let value_idx = chunk.add_constant(Value::String(Arc::from("value")));
        let item = reserve_slot(chunk);
        lget(chunk, items, line);
        lget(chunk, k, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, item, line);
        lget(chunk, item, line);
        class_slots::emit_class_get(
            chunk,
            ObjSource::Stack,
            &field_slot("name"),
            Dest::Stack,
            line,
        );
        convert::emit_to_string(chunk, line);
        lset(chunk, nm, line);
        lget(chunk, item, line);
        class_slots::emit_class_get(
            chunk,
            ObjSource::Stack,
            &field_slot("value"),
            Dest::Stack,
            line,
        );
        lset(chunk, vl, line);

        lget(chunk, nm, line);
        chunk.emit_string_const("", line);
        ops::emit_dyn_ne(chunk, line);
        chunk.emit_if(line);
        lget(chunk, names, line);
        lget(chunk, nm, line);
        (nm, vl)
    };
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], values, line);
    lget(&mut chunks[current], vl, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], ordered, line);
    lget(&mut chunks[current], vl, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, k, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_ADD, line);
        lset(chunk, k, line);
    }
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, split, line);
    let _ = nm;

    // ── scan the SQL, rewriting resolved `@name` to `?` ───────────────────
    let (out, outp, i, slen, instr, matched, code, end, name, cond, c2) = {
        let chunk = &mut chunks[current];
        (
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
        )
    };
    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], outp, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_string_const("", line);
        lset(chunk, out, line);
        chunk.emit_i32_const(0, line);
        lset(chunk, i, line);
        chunk.emit_i32_const(0, line);
        lset(chunk, instr, line);
        lget(chunk, sql0, line);
        vybe_compiler::primitives::strings::emit_length(chunk, line);
        lset(chunk, slen, line);
    }

    let scan = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, i, line);
        lget(chunk, slen, line);
        chunk.emit_op(Op::I32_LT_S, line);
    }
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);

    {
        let chunk = &mut chunks[current];
        chunk.emit_i32_const(0, line);
        lset(chunk, matched, line);
        lget(chunk, sql0, line);
        lget(chunk, i, line);
    }
    call_import(chunks, current, "wasm:js-string", "charCodeAt", 2, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, code, line);

        // A single quote flips literal mode; the append tail still runs.
        lget(chunk, code, line);
        chunk.emit_i32_const(b'\'' as i32, line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_if(line);
        chunk.emit_i32_const(1, line);
        lget(chunk, instr, line);
        chunk.emit_op(Op::I32_SUB, line);
        lset(chunk, instr, line);
        chunk.emit_end(line);

        // `@` outside a literal starts a placeholder.
        lget(chunk, instr, line);
        chunk.emit_op(Op::I32_EQZ, line);
        lget(chunk, code, line);
        chunk.emit_i32_const(b'@' as i32, line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_op(Op::I32_AND, line);
        chunk.emit_if(line);

        chunk.emit_string_const("@", line);
        lset(chunk, name, line);
        lget(chunk, i, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_ADD, line);
        lset(chunk, end, line);
    }

    let word = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_i32_const(0, line);
        lset(chunk, cond, line);
        lget(chunk, end, line);
        lget(chunk, slen, line);
        chunk.emit_op(Op::I32_LT_S, line);
        chunk.emit_if(line);
        lget(chunk, sql0, line);
        lget(chunk, end, line);
    }
    call_import(chunks, current, "wasm:js-string", "charCodeAt", 2, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, c2, line);
        emit_is_name_char(chunk, c2, line);
        lset(chunk, cond, line);
        chunk.emit_end(line);
        lget(chunk, cond, line);
    }
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, name, line);
        lget(chunk, sql0, line);
        lget(chunk, end, line);
    }
    call_import(chunks, current, "ecma:string", "charAt", 2, line);
    {
        let chunk = &mut chunks[current];
        vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
        lset(chunk, name, line);
        lget(chunk, end, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_ADD, line);
        lset(chunk, end, line);
    }
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, word, line);

    let j = {
        let chunk = &mut chunks[current];
        let j = reserve_slot(chunk);
        lget(chunk, names, line);
        lget(chunk, name, line);
        j
    };
    collections::emit_last_index_of(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::I32_FROM_F64, line);
        lset(chunk, j, line);
        lget(chunk, j, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op(Op::I32_GE_S, line);
        chunk.emit_if(line);

        lget(chunk, out, line);
        chunk.emit_string_const("?", line);
        vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
        lset(chunk, out, line);
        lget(chunk, outp, line);
        lget(chunk, values, line);
        lget(chunk, j, line);
        chunk.emit_op(Op::ARRAY_GET, line);
    }
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, end, line);
        lset(chunk, i, line);
        chunk.emit_i32_const(1, line);
        lset(chunk, matched, line);
        chunk.emit_end(line); // resolved?
        chunk.emit_end(line); // `@` outside a literal?

        // Nothing consumed the character — copy it and step on.
        lget(chunk, matched, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        lget(chunk, out, line);
        lget(chunk, sql0, line);
        lget(chunk, i, line);
    }
    call_import(chunks, current, "ecma:string", "charAt", 2, line);
    {
        let chunk = &mut chunks[current];
        vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
        lset(chunk, out, line);
        lget(chunk, i, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_ADD, line);
        lset(chunk, i, line);
        chunk.emit_end(line);
    }
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, scan, line);

    // ── no placeholder resolved ⇒ the original SQL, bound positionally ────
    let (sql_out, params_out) = {
        let chunk = &mut chunks[current];
        (reserve_slot(chunk), reserve_slot(chunk))
    };
    lget(&mut chunks[current], outp, line);
    collections::emit_len(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::I32_FROM_F64, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_if(line);
        lget(chunk, sql0, line);
        lset(chunk, sql_out, line);
        lget(chunk, ordered, line);
        lset(chunk, params_out, line);
        chunk.emit_else(line);
        lget(chunk, out, line);
        lset(chunk, sql_out, line);
        lget(chunk, outp, line);
        lset(chunk, params_out, line);
        chunk.emit_end(line);
    }
    (sql_out, params_out)
}

/// `command.CreateParameter(name, type, direction, size, value)` — the host
/// member read only the name and the value and dropped the rest; this does the
/// same, and dropping the other three is still a separate finding.
/// Stack in: `[command, p1, p2, p3, p4, p5]`  Stack out: `[parameter]`
pub fn emit_command_create_parameter(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    // Peel the user arguments off the top; `name` is the first and `value` the
    // fourth, counting from the receiver as the host's `args[1]` / `args[4]`.
    let mut name_slot = None;
    let mut value_slot = None;
    for position in (1..=argc).rev() {
        if position == 4 {
            let slot = reserve_slot(chunk);
            lset(chunk, slot, line);
            value_slot = Some(slot);
        } else if position == 1 {
            let slot = reserve_slot(chunk);
            lset(chunk, slot, line);
            name_slot = Some(slot);
        } else {
            chunk.emit_op(Op::DROP, line);
        }
    }
    chunk.emit_op(Op::DROP, line); // receiver

    class_slots::emit_class_alloc(chunk, line);
    set_field(
        chunk,
        "__type",
        |c, l| c.emit_string_const("SqlParameter", l),
        line,
    );
    set_field(
        chunk,
        "name",
        |c, l| match name_slot {
            Some(slot) => lget(c, slot, l),
            None => c.emit_string_const("", l),
        },
        line,
    );
    set_field(
        chunk,
        "value",
        |c, l| match value_slot {
            Some(slot) => lget(c, slot, l),
            None => push_void(c, l),
        },
        line,
    );
}

/// The ordered column names of a result set: every driver stamps `__col_names`
/// on each row (`sqlite.rs`, `postgres.rs`, `mysql.rs`), so the names travel
/// with the rows and need no second round-trip.
///
/// An EMPTY result carries none — `readwrite.query` answers `list<row>` and
/// nothing else. The host members asked the driver again via `query_columns`,
/// which is not in the WIT. PHP already lives with the same limit: its
/// `columnCount()` reads the keys of row 0 and answers 0 when there are none.
fn emit_col_names_of(chunks: &mut [Chunk], current: usize, rows_slot: u16, line: u32) -> u16 {
    let (names_slot, n_slot, first_slot) = {
        let chunk = &mut chunks[current];
        (reserve_slot(chunk), reserve_slot(chunk), reserve_slot(chunk))
    };
    lget(&mut chunks[current], rows_slot, line);
    collections::emit_len(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::I32_FROM_F64, line);
        lset(chunk, n_slot, line);
        lget(chunk, n_slot, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op(Op::I32_GT_S, line);
        chunk.emit_if(line);
        lget(chunk, rows_slot, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        class_slots::emit_class_get(
            chunk,
            ObjSource::Stack,
            &field_slot(COL_NAMES_KEY),
            Dest::Stack,
            line,
        );
        lset(chunk, first_slot, line);
        lget(chunk, first_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        lget(chunk, first_slot, line);
        lset(chunk, names_slot, line);
        chunk.emit_else(line);
    }
    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], names_slot, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_end(line);
        chunk.emit_else(line);
    }
    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], names_slot, line);
    chunks[current].emit_end(line);
    names_slot
}

/// `{ __type: "DataTable", tablename, columns, rows }` — the host's
/// `data_table_from_rows`.
fn emit_data_table(
    chunk: &mut Chunk,
    name_fn: impl FnOnce(&mut Chunk, u32),
    rows_slot: u16,
    names_slot: u16,
    line: u32,
) {
    class_slots::emit_class_alloc(chunk, line);
    set_field(
        chunk,
        "__type",
        |c, l| c.emit_string_const("DataTable", l),
        line,
    );
    set_field(chunk, "tablename", name_fn, line);
    set_field(chunk, "columns", |c, l| lget(c, names_slot, l), line);
    set_field(chunk, "rows", |c, l| lget(c, rows_slot, l), line);
}

/// Leave `[borrow<connection>, borrow<statement>]` on the stack for
/// `readwrite.query` / `readwrite.exec`. The command object doubles as the
/// connection operand: `wasi_id` reads `__conn_id` off whatever it is given.
fn emit_prepared_from_command(chunks: &mut [Chunk], current: usize, cmd_slot: u16, line: u32) {
    let (sql_slot, params_slot) = emit_command_sql_and_params(chunks, current, cmd_slot, line);
    let chunk = &mut chunks[current];
    lget(chunk, cmd_slot, line);
    lget(chunk, sql_slot, line);
    lget(chunk, params_slot, line);
    let _ = chunk;
    call_import(chunks, current, WASI_TYPES, STATEMENT_PREPARE, 2, line);
}

fn take_receiver(chunks: &mut [Chunk], current: usize, line: u32) -> u16 {
    let chunk = &mut chunks[current];
    let slot = reserve_slot(chunk);
    lset(chunk, slot, line);
    slot
}

/// `command.ExecuteNonQuery()` — affected row count, `-1` on failure (which is
/// what `readwrite.exec` itself answers).
/// Stack in: `[command]`  Stack out: `[number]`
pub fn emit_command_execute_non_query(chunks: &mut [Chunk], current: usize, line: u32) {
    let cmd_slot = take_receiver(chunks, current, line);
    emit_prepared_from_command(chunks, current, cmd_slot, line);
    call_import(chunks, current, WASI_READWRITE, "exec", 2, line);
}

/// `command.ExecuteScalar()` — the first column of the first row, `null` when
/// the result is empty.
/// Stack in: `[command]`  Stack out: `[value]`
pub fn emit_command_execute_scalar(chunks: &mut [Chunk], current: usize, line: u32) {
    let cmd_slot = take_receiver(chunks, current, line);
    emit_prepared_from_command(chunks, current, cmd_slot, line);
    call_import(chunks, current, WASI_READWRITE, "query", 2, line);
    let rows_slot = take_receiver(chunks, current, line);
    let names_slot = emit_col_names_of(chunks, current, rows_slot, line);

    lget(&mut chunks[current], names_slot, line);
    collections::emit_len(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::I32_FROM_F64, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if_value(line);
    lget(chunk, rows_slot, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::ARRAY_GET, line); // [row]
    lget(chunk, names_slot, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::ARRAY_GET, line); // [row, name]
    chunk.emit_op(Op::ARRAY_GET, line); // [value]
    chunk.emit_else(line);
    push_void(chunk, line);
    chunk.emit_end(line);
}

/// `command.ExecuteReader()` — the materialised result set as a cursor object.
/// Stack in: `[command]`  Stack out: `[reader]`
pub fn emit_command_execute_reader(chunks: &mut [Chunk], current: usize, line: u32) {
    let cmd_slot = take_receiver(chunks, current, line);
    emit_prepared_from_command(chunks, current, cmd_slot, line);
    call_import(chunks, current, WASI_READWRITE, "query", 2, line);
    let rows_slot = take_receiver(chunks, current, line);
    emit_reader_struct(chunks, current, rows_slot, line);
}

/// `{ __type: "SqlDataReader", __rows, __col_names, __pos: -1, … }` — the shape
/// the reader family above reads.
fn emit_reader_struct(chunks: &mut [Chunk], current: usize, rows_slot: u16, line: u32) {
    let names_slot = emit_col_names_of(chunks, current, rows_slot, line);
    let (has_slot, count_slot) = {
        let chunk = &mut chunks[current];
        (reserve_slot(chunk), reserve_slot(chunk))
    };
    lget(&mut chunks[current], rows_slot, line);
    collections::emit_len(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::I32_FROM_F64, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op(Op::I32_GT_S, line);
        ops::emit_i32_to_bool(chunk, line);
        lset(chunk, has_slot, line);
        lget(chunk, names_slot, line);
    }
    collections::emit_len(chunks, current, line);
    lset(&mut chunks[current], count_slot, line);

    let chunk = &mut chunks[current];
    class_slots::emit_class_construct(
        chunk,
        "SqlDataReader",
        &[
            (field_slot(ROWS_KEY), ValueSource::Local(rows_slot)),
            (field_slot(COL_NAMES_KEY), ValueSource::Local(names_slot)),
            (field_slot(POS_KEY), ValueSource::ConstF64(-1.0)),
            (field_slot("hasrows"), ValueSource::Local(has_slot)),
            (field_slot("fieldcount"), ValueSource::Local(count_slot)),
            (field_slot(IS_CLOSED_KEY), ValueSource::ConstBool(false)),
        ],
        line,
    );
}

// ── SqlConnection.GetSchema ───────────────────────────────────────────────────

/// The dialect's "list every user table" query, chosen at runtime from the
/// connection's `provider`. These strings were `SqlDriver::tables_sql` on the
/// three drivers; selecting one is a string choice, not a host operation.
const SQLITE_TABLES_SQL: &str =
    "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name";
const POSTGRES_TABLES_SQL: &str = "SELECT table_name AS name FROM information_schema.tables \
     WHERE table_schema = 'public' ORDER BY table_name";
const MYSQL_TABLES_SQL: &str = "SELECT table_name AS name FROM information_schema.tables \
     WHERE table_schema = DATABASE() ORDER BY table_name";

/// `connection.GetSchema(collection)`.
///
/// Only `""` / `"Tables"` reaches a query. The descriptor declares a single
/// 1-arity overload, so `GetSchema("Columns", table)` cannot be written — the
/// host's `"columns"` branch read an `args[2]` no route supplies, built
/// `columns_sql("")`, and the query always failed back to an empty `Columns`
/// table. That empty table is produced directly here rather than by issuing
/// SQL known to be malformed.
///
/// Stack in: `[connection, collection]`  Stack out: `[DataTable]`
pub fn emit_connection_get_schema(chunks: &mut [Chunk], current: usize, line: u32) {
    let (_conn_slot, want_slot, sql_slot, kind_slot) = {
        let chunk = &mut chunks[current];
        let want_slot = reserve_slot(chunk);
        let conn_slot = reserve_slot(chunk);
        let sql_slot = reserve_slot(chunk);
        let kind_slot = reserve_slot(chunk);
        convert::emit_to_string(chunk, line);
        vybe_compiler::primitives::strings::emit_to_lower(chunk, line);
        lset(chunk, want_slot, line);
        lset(chunk, conn_slot, line);

        // kind: 0 = tables (query), 1 = columns (empty), 2 = unknown (empty)
        chunk.emit_i32_const(2, line);
        lset(chunk, kind_slot, line);
        lget(chunk, want_slot, line);
        chunk.emit_string_const("", line);
        ops::emit_dyn_eq(chunk, line);
        lget(chunk, want_slot, line);
        chunk.emit_string_const("tables", line);
        ops::emit_dyn_eq(chunk, line);
        chunk.emit_op(Op::I32_OR, line);
        chunk.emit_if(line);
        chunk.emit_i32_const(0, line);
        lset(chunk, kind_slot, line);
        chunk.emit_end(line);
        lget(chunk, want_slot, line);
        chunk.emit_string_const("columns", line);
        ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
        chunk.emit_i32_const(1, line);
        lset(chunk, kind_slot, line);
        chunk.emit_end(line);

        // The dialect query for this connection's provider.
        chunk.emit_string_const(SQLITE_TABLES_SQL, line);
        lset(chunk, sql_slot, line);
        get_prop(chunk, conn_slot, "provider", line);
        chunk.emit_string_const("postgres", line);
        ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
        chunk.emit_string_const(POSTGRES_TABLES_SQL, line);
        lset(chunk, sql_slot, line);
        chunk.emit_end(line);
        get_prop(chunk, conn_slot, "provider", line);
        chunk.emit_string_const("mysql", line);
        ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
        chunk.emit_string_const(MYSQL_TABLES_SQL, line);
        lset(chunk, sql_slot, line);
        chunk.emit_end(line);

        lget(chunk, kind_slot, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_if_value(line);
        lget(chunk, conn_slot, line);
        lget(chunk, sql_slot, line);
        (conn_slot, want_slot, sql_slot, kind_slot)
    };
    collections::emit_array_new(chunks, current, 0, line);
    call_import(chunks, current, WASI_TYPES, STATEMENT_PREPARE, 2, line);
    call_import(chunks, current, WASI_READWRITE, "query", 2, line);
    let rows_slot = take_receiver(chunks, current, line);
    let names_slot = emit_col_names_of(chunks, current, rows_slot, line);

    // An empty `Tables` result still reports the column the host fell back to.
    lget(&mut chunks[current], names_slot, line);
    collections::emit_len(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::I32_FROM_F64, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_if(line);
        chunk.emit_string_const("name", line);
        chunk.emit_array_new_fixed(0, 1, line);
        lset(chunk, names_slot, line);
        chunk.emit_end(line);
        emit_data_table(
            chunk,
            |c, l| c.emit_string_const("Tables", l),
            rows_slot,
            names_slot,
            line,
        );
        chunk.emit_else(line);
    }

    // `Columns` (unreachable overload) and anything unknown: an empty table.
    let empty_rows = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        slot
    };
    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], empty_rows, line);
    {
        let chunk = &mut chunks[current];
        let empty_names = reserve_slot(chunk);
        lget(chunk, kind_slot, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_if_value(line);
        chunk.emit_string_const("column_name", line);
        chunk.emit_array_new_fixed(0, 1, line);
        chunk.emit_else(line);
        chunk.emit_array_new_fixed(0, 0, line);
        chunk.emit_end(line);
        lset(chunk, empty_names, line);

        lget(chunk, kind_slot, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_if_value(line);
        emit_data_table(
            chunk,
            |c, l| c.emit_string_const("Columns", l),
            empty_rows,
            empty_names,
            line,
        );
        chunk.emit_else(line);
        emit_data_table(
            chunk,
            |c, l| c.emit_string_const("Schema", l),
            empty_rows,
            empty_names,
            line,
        );
        chunk.emit_end(line);
        chunk.emit_end(line);
    }
    let _ = (want_slot, sql_slot);
}

// ── SqlDataAdapter ────────────────────────────────────────────────────────────

/// `new SqlDataAdapter(selectSql?, connection?)`.
/// Stack in: `[sql?, connection?]`  Stack out: `[adapter]`
pub fn emit_data_adapter_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let conn_slot = if argc >= 2 {
        let slot = reserve_slot(chunk);
        lset(chunk, slot, line);
        Some(slot)
    } else {
        None
    };
    let sql_slot = if argc >= 1 {
        let slot = reserve_slot(chunk);
        lset(chunk, slot, line);
        Some(slot)
    } else {
        None
    };
    for _ in 2..argc {
        chunk.emit_op(Op::DROP, line);
    }

    class_slots::emit_class_alloc(chunk, line);
    set_field(
        chunk,
        "__type",
        |c, l| c.emit_string_const("SqlDataAdapter", l),
        line,
    );
    set_field(
        chunk,
        "selectcommand",
        |c, l| match sql_slot {
            Some(slot) => lget(c, slot, l),
            None => c.emit_string_const("", l),
        },
        line,
    );
    set_field(
        chunk,
        CONN_ID_KEY,
        |c, l| match conn_slot {
            Some(slot) => get_prop(c, slot, CONN_ID_KEY, l),
            None => c.emit_f64_const(0.0, l),
        },
        line,
    );
    set_field(
        chunk,
        "connectionstring",
        |c, l| match conn_slot {
            Some(slot) => get_prop(c, slot, "connectionstring", l),
            None => c.emit_string_const("", l),
        },
        line,
    );
}

/// `adapter.Fill(target)` — run `SelectCommand` and write the result into a
/// `DataTable` (columns + rows replaced) or a `DataSet` (a new table appended).
/// Answers the row count.
/// Stack in: `[adapter, target]`  Stack out: `[number]`
pub fn emit_adapter_fill(chunks: &mut [Chunk], current: usize, line: u32) {
    let (adapter_slot, target_slot) = {
        let chunk = &mut chunks[current];
        let target_slot = reserve_slot(chunk);
        let adapter_slot = reserve_slot(chunk);
        lset(chunk, target_slot, line);
        lset(chunk, adapter_slot, line);
        lget(chunk, adapter_slot, line);
        get_prop(chunk, adapter_slot, "selectcommand", line);
        (adapter_slot, target_slot)
    };
    collections::emit_array_new(chunks, current, 0, line);
    call_import(chunks, current, WASI_TYPES, STATEMENT_PREPARE, 2, line);
    call_import(chunks, current, WASI_READWRITE, "query", 2, line);
    let rows_slot = take_receiver(chunks, current, line);
    let names_slot = emit_col_names_of(chunks, current, rows_slot, line);

    let tables_slot = {
        let chunk = &mut chunks[current];
        let tables_slot = reserve_slot(chunk);
        get_prop(chunk, target_slot, "__type", line);
        convert::emit_to_string(chunk, line);
        vybe_compiler::primitives::strings::emit_to_lower(chunk, line);
        chunk.emit_string_const("dataset", line);
        ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);

        get_prop(chunk, target_slot, "tables", line);
        lset(chunk, tables_slot, line);
        lget(chunk, tables_slot, line);
        chunk.emit_string_const("Table", line);
        tables_slot
    };
    lget(&mut chunks[current], tables_slot, line);
    collections::emit_len(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::I32_FROM_F64, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_ADD, line);
        convert::emit_to_string(chunk, line);
        vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
        let name_slot = reserve_slot(chunk);
        lset(chunk, name_slot, line);
        emit_data_table(
            chunk,
            |c, l| lget(c, name_slot, l),
            rows_slot,
            names_slot,
            line,
        );
    }
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_else(line);
        set_prop_local(chunk, target_slot, "columns", names_slot, line);
        set_prop_local(chunk, target_slot, "rows", rows_slot, line);
        chunk.emit_end(line);
        lget(chunk, rows_slot, line);
    }
    collections::emit_len(chunks, current, line);
    let _ = adapter_slot;
}

// ── SqlDataReader ─────────────────────────────────────────────────────────────

/// Reader state, unpacked into locals: `(rows, col_names, pos_as_i32)`.
fn reader_state(chunks: &mut [Chunk], current: usize, reader_slot: u16, line: u32) -> (u16, u16, u16) {
    let chunk = &mut chunks[current];
    let rows_slot = reserve_slot(chunk);
    let names_slot = reserve_slot(chunk);
    let pos_slot = reserve_slot(chunk);

    get_prop(chunk, reader_slot, ROWS_KEY, line);
    lset(chunk, rows_slot, line);
    get_prop(chunk, reader_slot, COL_NAMES_KEY, line);
    lset(chunk, names_slot, line);
    get_prop(chunk, reader_slot, POS_KEY, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
    lset(chunk, pos_slot, line);

    (rows_slot, names_slot, pos_slot)
}

/// `[reader, index]` → the cursor is on a row AND `index` names a column, as an
/// i32 0/1, with the reader's state left in the returned locals.
///
/// This is `current_reader_row` + `row_value_by_index`'s bounds test, which the
/// host answered as two `Option`s. Emitted, it is one condition guarding both
/// `ARRAY_GET`s.
fn emit_field_guard(
    chunks: &mut [Chunk],
    current: usize,
    reader_slot: u16,
    idx_slot: u16,
    line: u32,
) -> (u16, u16, u16) {
    let (rows_slot, names_slot, pos_slot) = reader_state(chunks, current, reader_slot, line);
    emit_index_in_range(chunks, current, rows_slot, pos_slot, line);
    emit_index_in_range(chunks, current, names_slot, idx_slot, line);
    chunks[current].emit_op(Op::I32_AND, line);
    (rows_slot, names_slot, pos_slot)
}

/// `rows[pos][col_names[index]]` — the current row's value for a column
/// ordinal. Assumes the guard above already passed.
fn emit_field_value(
    chunk: &mut Chunk,
    rows_slot: u16,
    names_slot: u16,
    pos_slot: u16,
    idx_slot: u16,
    line: u32,
) {
    lget(chunk, rows_slot, line);
    lget(chunk, pos_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line); // [row]
    lget(chunk, names_slot, line);
    lget(chunk, idx_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line); // [row, name]
    chunk.emit_op(Op::ARRAY_GET, line); // [value]
}

/// The `[reader, index]` prologue: both operands into locals, index as i32.
fn take_reader_and_index(chunks: &mut [Chunk], current: usize, line: u32) -> (u16, u16) {
    let chunk = &mut chunks[current];
    let idx_slot = reserve_slot(chunk);
    let reader_slot = reserve_slot(chunk);
    chunk.emit_op(Op::I32_FROM_F64, line);
    lset(chunk, idx_slot, line);
    lset(chunk, reader_slot, line);
    (reader_slot, idx_slot)
}

/// `reader.Read()` — advance the cursor, answer whether it landed on a row.
///
/// Stack in: `[reader]`  Stack out: `[Bool]`
///
/// The result is a real `Bool`, not the comparison's i32: VB's `Not` is bitwise
/// on a number, so `Do While Not reader.Read()` would never terminate against
/// an i32. See `adodb_adapter::comparison_to_bool`.
pub fn emit_reader_read(chunks: &mut [Chunk], current: usize, line: u32) {
    let (reader_slot, pos_slot) = {
        let chunk = &mut chunks[current];
        (reserve_slot(chunk), reserve_slot(chunk))
    };
    {
        let chunk = &mut chunks[current];
        lset(chunk, reader_slot, line);
        get_prop(chunk, reader_slot, POS_KEY, line);
        chunk.emit_f64_const(1.0, line);
        ops::emit_dyn_add(chunk, line);
        lset(chunk, pos_slot, line);
        set_prop_local(chunk, reader_slot, POS_KEY, pos_slot, line);

        lget(chunk, pos_slot, line);
        get_prop(chunk, reader_slot, ROWS_KEY, line);
    }
    collections::emit_len(chunks, current, line);
    let chunk = &mut chunks[current];
    ops::emit_dyn_lt(chunk, line);
    ops::emit_i32_to_bool(chunk, line);
}

/// `reader.GetName(i)` — the i-th column name, `""` out of range.
/// Stack in: `[reader, index]`  Stack out: `[String]`
pub fn emit_reader_get_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let (reader_slot, idx_slot) = take_reader_and_index(chunks, current, line);
    let names_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        get_prop(chunk, reader_slot, COL_NAMES_KEY, line);
        lset(chunk, slot, line);
        slot
    };
    emit_index_in_range(chunks, current, names_slot, idx_slot, line);

    let chunk = &mut chunks[current];
    chunk.emit_if_value(line);
    lget(chunk, names_slot, line);
    lget(chunk, idx_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_else(line);
    chunk.emit_string_const("", line);
    chunk.emit_end(line);
}

/// `reader.GetValue(i)` — the current row's i-th value, `null` out of range.
/// Stack in: `[reader, index]`  Stack out: `[value]`
pub fn emit_reader_get_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let (reader_slot, idx_slot) = take_reader_and_index(chunks, current, line);
    let (rows_slot, names_slot, pos_slot) =
        emit_field_guard(chunks, current, reader_slot, idx_slot, line);

    let chunk = &mut chunks[current];
    chunk.emit_if_value(line);
    emit_field_value(chunk, rows_slot, names_slot, pos_slot, idx_slot, line);
    chunk.emit_else(line);
    push_void(chunk, line);
    chunk.emit_end(line);
}

/// `reader.GetString(i)` — the current row's i-th value as a string, `""` when
/// the cursor is off the end (the host's `format!` of a real `null` still
/// reads `"null"`, and `ecma:string.String` agrees).
/// Stack in: `[reader, index]`  Stack out: `[String]`
pub fn emit_reader_get_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let (reader_slot, idx_slot) = take_reader_and_index(chunks, current, line);
    let (rows_slot, names_slot, pos_slot) =
        emit_field_guard(chunks, current, reader_slot, idx_slot, line);

    let chunk = &mut chunks[current];
    chunk.emit_if_value(line);
    emit_field_value(chunk, rows_slot, names_slot, pos_slot, idx_slot, line);
    convert::emit_to_string(chunk, line);
    chunk.emit_else(line);
    chunk.emit_string_const("", line);
    chunk.emit_end(line);
}

/// `reader.IsDBNull(i)` — true for a SQL NULL, and true when there is no row.
/// Stack in: `[reader, index]`  Stack out: `[Bool]`
pub fn emit_reader_is_dbnull(chunks: &mut [Chunk], current: usize, line: u32) {
    let (reader_slot, idx_slot) = take_reader_and_index(chunks, current, line);
    let (rows_slot, names_slot, pos_slot) =
        emit_field_guard(chunks, current, reader_slot, idx_slot, line);

    let chunk = &mut chunks[current];
    chunk.emit_if_value(line);
    emit_field_value(chunk, rows_slot, names_slot, pos_slot, idx_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    ops::emit_i32_to_bool(chunk, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_end(line);
}

/// `reader.Close()` — marks the reader closed. The rows are already materialised
/// in the guest, so there is no cursor to release.
/// Stack in: `[reader]`  Stack out: `[null]`
pub fn emit_reader_close(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let closed_idx = chunk.add_constant(Value::String(Arc::from(IS_CLOSED_KEY)));
    chunk.emit_bool_const(true, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(IS_CLOSED_KEY),
        ValueSource::Stack,
        line,
    );
    push_void(chunk, line);
}

/// `reader.GetSchemaTable()` — a DataTable of `ColumnName` / `ColumnOrdinal`,
/// one row per column.
/// Stack in: `[reader]`  Stack out: `[DataTable]`
pub fn emit_reader_get_schema_table(chunks: &mut [Chunk], current: usize, line: u32) {
    let (reader_slot, names_slot, out_slot, i_slot, len_slot) = {
        let chunk = &mut chunks[current];
        (
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
        )
    };
    {
        let chunk = &mut chunks[current];
        lset(chunk, reader_slot, line);
        get_prop(chunk, reader_slot, COL_NAMES_KEY, line);
        lset(chunk, names_slot, line);
        chunk.emit_i32_const(0, line);
        lset(chunk, i_slot, line);
    }
    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], out_slot, line);
    lget(&mut chunks[current], names_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_FROM_F64, line);
    lset(&mut chunks[current], len_slot, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);

    lget(&mut chunks[current], out_slot, line);
    {
        let chunk = &mut chunks[current];
        class_slots::emit_class_alloc(chunk, line);
        set_field(
            chunk,
            "ColumnName",
            |c, l| {
                lget(c, names_slot, l);
                lget(c, i_slot, l);
                c.emit_op(Op::ARRAY_GET, l);
            },
            line,
        );
        set_field(
            chunk,
            "ColumnOrdinal",
            |c, l| {
                lget(c, i_slot, l);
                c.emit_op(Op::F64_FROM_I32, l);
            },
            line,
        );
    }
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    let chunk = &mut chunks[current];
    class_slots::emit_class_construct(
        chunk,
        "DataTable",
        &[
            (field_slot("tablename"), ValueSource::ConstStr("SchemaTable".to_string())),
        ],
        line,
    );
    set_field(
        chunk,
        "columns",
        |c, l| {
            c.emit_string_const("ColumnName", l);
            c.emit_string_const("ColumnOrdinal", l);
            c.emit_array_new_fixed(0, 2, l);
        },
        line,
    );
    set_field(chunk, "rows", |c, l| lget(c, out_slot, l), line);
}

// ── SqlParameterCollection ────────────────────────────────────────────────────

/// `Parameters.Add(name, value)` / `AddWithValue(name, value)`.
/// Stack in: `[params, name, value]`  Stack out: `[null]`
pub fn emit_params_add_with_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let (params_slot, name_slot, value_slot) = {
        let chunk = &mut chunks[current];
        let value_slot = reserve_slot(chunk);
        let name_slot = reserve_slot(chunk);
        let params_slot = reserve_slot(chunk);
        lset(chunk, value_slot, line);
        lset(chunk, name_slot, line);
        lset(chunk, params_slot, line);
        (params_slot, name_slot, value_slot)
    };
    {
        let chunk = &mut chunks[current];
        get_prop(chunk, params_slot, ITEMS_KEY, line);
        class_slots::emit_class_construct(
            chunk,
            "SqlParameter",
            &[
                (field_slot("name"), ValueSource::Local(name_slot)),
                (field_slot("value"), ValueSource::Local(value_slot)),
            ],
            line,
        );
    }
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    push_void(&mut chunks[current], line);
}

/// `Parameters.Clear()` — the command reads the collection through the same
/// object, so replacing `__items` is what clearing it in place was.
/// Stack in: `[params]`  Stack out: `[null]`
pub fn emit_params_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    let params_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        lset(chunk, slot, line);
        slot
    };
    collections::emit_array_new(chunks, current, 0, line);
    let items_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        lset(chunk, slot, line);
        slot
    };
    let chunk = &mut chunks[current];
    set_prop_local(chunk, params_slot, ITEMS_KEY, items_slot, line);
    push_void(chunk, line);
}

/// `Parameters.Count`.
/// Stack in: `[params]`  Stack out: `[number]`
pub fn emit_params_count(chunks: &mut [Chunk], current: usize, line: u32) {
    let items_idx = chunks[current].add_constant(Value::String(Arc::from(ITEMS_KEY)));
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Stack,
        &field_slot(ITEMS_KEY),
        Dest::Stack,
        line,
    );
    collections::emit_len(chunks, current, line);
}

// ── SqlTransaction ────────────────────────────────────────────────────────────

/// `COMMIT` / `ROLLBACK` over the real WIT: prepare the statement, exec it
/// against the connection the transaction carries, mark the transaction closed.
///
/// `readwrite.exec` answers the affected-row count, or `-1` on failure — the
/// host member answered `Bool(is_ok())`, which is `count >= 0`.
///
/// Stack in: `[tx]`  Stack out: `[Bool]`
fn emit_transaction_verb(chunks: &mut [Chunk], current: usize, sql: &str, line: u32) {
    let tx_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        lset(chunk, slot, line);
        slot
    };
    lget(&mut chunks[current], tx_slot, line); // borrow<connection>
    chunks[current].emit_string_const(sql, line);
    collections::emit_array_new(chunks, current, 0, line);
    call_import(chunks, current, WASI_TYPES, STATEMENT_PREPARE, 2, line);
    call_import(chunks, current, WASI_READWRITE, "exec", 2, line);

    let chunk = &mut chunks[current];
    let ok_slot = reserve_slot(chunk);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_ge(chunk, line);
    ops::emit_i32_to_bool(chunk, line);
    lset(chunk, ok_slot, line);

    let closed_idx = chunk.add_constant(Value::String(Arc::from(IS_CLOSED_KEY)));
    lget(chunk, tx_slot, line);
    chunk.emit_bool_const(true, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(IS_CLOSED_KEY),
        ValueSource::Stack,
        line,
    );

    lget(chunk, ok_slot, line);
}

pub fn emit_transaction_commit(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_transaction_verb(chunks, current, "COMMIT", line);
}

pub fn emit_transaction_rollback(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_transaction_verb(chunks, current, "ROLLBACK", line);
}
