//! Shared database backend for PHP's `PDO` and `mysqli` object surfaces.
//!
//! Both are thin adapters over the one `wasi:sql` component interface. This
//! module owns the operations they share (statement construction, etc.); the
//! `pdo_adapter`/`mysqli_adapter` files hold only the class-specific wrappers
//! (constructors + a couple of shape fields). A statement's class identity
//! (`PDOStatement` vs `mysqli_stmt`) is stamped from the receiver connection's
//! `__type` — the only real difference between the two.

use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    chunk.emit_string_const(value, line);
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

fn emit_empty_array(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
}

/// The statement class name for a connection whose `__type` is on top of the
/// stack: `"mysqli"` → `"mysqli_stmt"`, anything else → `"PDOStatement"`.
/// Consumes the `__type` value, leaves the label on the stack.
fn emit_stmt_label_from_type(chunk: &mut Chunk, line: u32) {
    // stack: [conn.__type]
    push_str(chunk, "mysqli", line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "mysqli_stmt", line);
    chunk.emit_else(line);
    push_str(chunk, "PDOStatement", line);
    chunk.emit_end(line);
}

/// PHP `$conn->prepare($sql)` for both PDO and mysqli. Stack: `[conn, sql]` →
/// `[stmt]`. Builds one shared statement shape; the `__type` class label is
/// derived from the receiver connection so the same op serves both classes.
pub fn emit_db_prepare(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let sql_slot = alloc_local(chunk);
    let conn_slot = alloc_local(chunk);
    lset(chunk, sql_slot, line);
    lset(chunk, conn_slot, line);

    // stmt = wasi:sql.createCommand(conn)
    lget(&mut chunks[current], conn_slot, line);
    call_import(chunks, current, "wasi:sql", "createCommand", 1, line);
    let chunk = &mut chunks[current];
    let stmt_slot = alloc_local(chunk);
    lset(chunk, stmt_slot, line);

    // label = (conn.__type == "mysqli") ? "mysqli_stmt" : "PDOStatement"
    let label_slot = alloc_local(chunk);
    lget(chunk, conn_slot, line);
    struct_get_key(chunk, "__type", line);
    emit_stmt_label_from_type(chunk, line);
    lset(chunk, label_slot, line);

    // stmt.__type = label
    lget(chunk, stmt_slot, line);
    lget(chunk, label_slot, line);
    struct_set_key(chunk, "__type", line);

    // Command text (PDO reads `commandtext`; mysqli reads `__sql`).
    for key in ["commandtext", "__prepared_commandtext", "__sql"] {
        lget(chunk, stmt_slot, line);
        lget(chunk, sql_slot, line);
        struct_set_key(chunk, key, line);
    }
    // Owning connection (PDO reads `__conn`; mysqli reads `__mysqli`).
    for key in ["__conn", "__mysqli"] {
        lget(chunk, stmt_slot, line);
        lget(chunk, conn_slot, line);
        struct_set_key(chunk, key, line);
    }
    // Cursor + bound-parameter scaffolding shared by both surfaces.
    lget(chunk, stmt_slot, line);
    chunk.emit_f64_const(0.0, line);
    struct_set_key(chunk, "__cursor", line);
    for key in [
        "__rows",
        "__bound_params",
        "__bound_named_pairs",
        "__bound_result",
    ] {
        lget(&mut chunks[current], stmt_slot, line);
        emit_empty_array(chunks, current, line);
        struct_set_key(&mut chunks[current], key, line);
    }

    // mysqli_stmt status properties (read as `$stmt->prop`). Defaults match a
    // freshly-prepared statement; `execute` updates the row-count fields.
    let chunk = &mut chunks[current];
    for (key, val) in [
        ("errno", 0.0),
        ("insert_id", 0.0),
        ("affected_rows", 0.0),
        ("num_rows", 0.0),
        ("field_count", 0.0),
    ] {
        lget(chunk, stmt_slot, line);
        chunk.emit_f64_const(val, line);
        struct_set_key(chunk, key, line);
    }
    for key in ["error", "sqlstate"] {
        lget(chunk, stmt_slot, line);
        push_str(chunk, if key == "sqlstate" { "00000" } else { "" }, line);
        struct_set_key(chunk, key, line);
    }

    // param_count = number of `?` placeholders = split(sql, "?").length - 1.
    lget(chunk, stmt_slot, line);
    lget(chunk, sql_slot, line);
    push_str(chunk, "?", line);
    {
        let idx = chunk.add_import("ecma:string", "split");
        chunk.emit_call(idx, 2, line);
    }
    crate::emitter::collections::emit_array_length(chunk, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_SUB, line);
    struct_set_key(chunk, "param_count", line);

    lget(chunk, stmt_slot, line);
}
