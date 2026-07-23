use std::sync::Arc;
use vybe_emitter::instructions::core_wasm;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use vybe_emitter::collections;

const COL_NAMES_KEY: &str = "__col_names";
const COMMAND_TYPE_KEY: &str = "commandtype";
const EOF_KEY: &str = "eof";
const ISCLOSED_KEY: &str = "isclosed";
const POS_KEY: &str = "__pos";
const RECORD_COUNT_KEY: &str = "recordcount";
const ROWS_KEY: &str = "__rows";
const VALUE_KEY: &str = "value";

// ── Private helpers ───────────────────────────────────────────────────────────

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
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
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(argc, line);
}

fn set_const_prop(chunk: &mut Chunk, key: &str, value: Value, line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    core_wasm::dup(chunk, line);
    push_const(chunk, value, line);
    chunk.emit_op_u16(Op::STRUCT_SET, key_idx, line);
    chunk.emit_op(Op::DROP, line);
}

fn set_local_prop(chunk: &mut Chunk, key: &str, local: u16, line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, local, line);
    chunk.emit_op_u16(Op::STRUCT_SET, key_idx, line);
    chunk.emit_op(Op::DROP, line);
}

fn set_object_local_prop(chunk: &mut Chunk, object_local: u16, key: &str, local: u16, line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::LOCAL_GET, object_local, line);
    chunk.emit_op_u16(Op::LOCAL_GET, local, line);
    chunk.emit_op_u16(Op::STRUCT_SET, key_idx, line);
    chunk.emit_op(Op::DROP, line);
}

fn set_object_const_prop(chunk: &mut Chunk, object_local: u16, key: &str, value: Value, line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::LOCAL_GET, object_local, line);
    push_const(chunk, value, line);
    chunk.emit_op_u16(Op::STRUCT_SET, key_idx, line);
    chunk.emit_op(Op::DROP, line);
}

fn get_prop_to_local(
    chunk: &mut Chunk,
    object_local: u16,
    key: &str,
    target_local: u16,
    line: u32,
) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::LOCAL_GET, object_local, line);
    chunk.emit_op_u16(Op::STRUCT_GET, key_idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, target_local, line);
}

fn build_field_object(chunks: &mut [Chunk], current: usize, value_local: u16, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    set_local_prop(chunk, VALUE_KEY, value_local, line);
}

/// Read `__rows` and `__col_names` from `reader_slot`, compute `eof` and
/// `recordcount`, then build a new ADODB Recordset struct and leave it on
/// the stack.  Used by `emit_adodb_connection_execute` and
/// `emit_adodb_command_execute`.
fn emit_reader_to_adodb_recordset(
    chunks: &mut [Chunk],
    current: usize,
    reader_slot: u16,
    line: u32,
) {
    let (rows_slot, cols_slot) = {
        let chunk = &mut chunks[current];
        (reserve_slot(chunk), reserve_slot(chunk))
    };
    {
        let chunk = &mut chunks[current];
        get_prop_to_local(chunk, reader_slot, ROWS_KEY, rows_slot, line);
        get_prop_to_local(chunk, reader_slot, COL_NAMES_KEY, cols_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, rows_slot, line);
    }
    collections::emit_len(chunks, current, line);
    let (count_slot, eof_slot) = {
        let chunk = &mut chunks[current];
        let count_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
        core_wasm::i32_const(chunk, line, 0);
        vybe_emitter::ops::emit_dyn_eq(chunk, line);
        let eof_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, eof_slot, line);
        (count_slot, eof_slot)
    };
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    set_local_prop(chunk, ROWS_KEY, rows_slot, line);
    set_local_prop(chunk, COL_NAMES_KEY, cols_slot, line);
    set_const_prop(chunk, POS_KEY, Value::F64(0.0), line);
    set_local_prop(chunk, RECORD_COUNT_KEY, count_slot, line);
    set_local_prop(chunk, EOF_KEY, eof_slot, line);
    set_const_prop(chunk, ISCLOSED_KEY, Value::Bool(false), line);
}

// ── ADODB.Connection ──────────────────────────────────────────────────────────

pub fn emit_adodb_connection_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    call_import(
        chunks,
        current,
        "wasi:sql/types",
        "connection.new",
        argc,
        line,
    );
}

/// `Connection.Execute(sql)` — creates a command, runs it, returns a Recordset.
pub fn emit_adodb_connection_execute(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // Stack: [conn, sql]
    let (sql_slot, conn_slot) = {
        let chunk = &mut chunks[current];
        (reserve_slot(chunk), reserve_slot(chunk))
    };
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, sql_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, conn_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, sql_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, conn_slot, line);
    }
    call_import(chunks, current, "wasi:sql/types", "command.new", 2, line);
    call_import(
        chunks,
        current,
        "wasi:sql/types",
        "[method]command.execute-reader",
        1,
        line,
    );
    let reader_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        slot
    };
    emit_reader_to_adodb_recordset(chunks, current, reader_slot, line);
}

/// `Connection.BeginTrans` — starts a transaction; the transaction object
/// returned by the host is discarded (ADODB tracks state implicitly).
pub fn emit_adodb_conn_begin_trans(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // Stack: [conn]
    call_import(
        chunks,
        current,
        "wasi:sql/types",
        "[method]connection.begin-transaction",
        1,
        line,
    );
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// `Connection.CommitTrans` — executes COMMIT on the connection.
pub fn emit_adodb_conn_commit_trans(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // Stack: [conn]
    let chunk = &mut chunks[current];
    let conn_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, conn_slot, line);
    push_const(chunk, Value::String(Arc::from("COMMIT")), line);
    chunk.emit_op_u16(Op::LOCAL_GET, conn_slot, line);
    call_import(chunks, current, "wasi:sql/types", "command.new", 2, line);
    call_import(
        chunks,
        current,
        "wasi:sql/types",
        "[method]command.execute-non-query",
        1,
        line,
    );
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// `Connection.RollbackTrans` — executes ROLLBACK on the connection.
pub fn emit_adodb_conn_rollback_trans(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // Stack: [conn]
    let chunk = &mut chunks[current];
    let conn_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, conn_slot, line);
    push_const(chunk, Value::String(Arc::from("ROLLBACK")), line);
    chunk.emit_op_u16(Op::LOCAL_GET, conn_slot, line);
    call_import(chunks, current, "wasi:sql/types", "command.new", 2, line);
    call_import(
        chunks,
        current,
        "wasi:sql/types",
        "[method]command.execute-non-query",
        1,
        line,
    );
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

// ── ADODB.Command ─────────────────────────────────────────────────────────────

pub fn emit_adodb_command_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    call_import(chunks, current, "wasi:sql/types", "command.new", argc, line);
    let chunk = &mut chunks[current];
    set_const_prop(chunk, COMMAND_TYPE_KEY, Value::F64(1.0), line);
}

/// `Command.Execute` — command already has CommandText + Parameters set;
/// calls execute-reader and wraps the result as an ADODB Recordset.
pub fn emit_adodb_command_execute(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // Stack: [cmd]
    call_import(
        chunks,
        current,
        "wasi:sql/types",
        "[method]command.execute-reader",
        1,
        line,
    );
    let reader_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        slot
    };
    emit_reader_to_adodb_recordset(chunks, current, reader_slot, line);
}

pub fn emit_adodb_command_create_parameter(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let value_slot = reserve_slot(chunk);
    let name_slot = reserve_slot(chunk);

    if argc > 0 {
        chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    } else {
        chunk.emit_op(Op::DROP, line);
        push_const(chunk, Value::Null, line);
        chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
        push_const(chunk, Value::String(Arc::from("")), line);
        chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
        build_field_object(chunks, current, value_slot, line);
        return;
    }

    for _ in 0..argc.saturating_sub(2) {
        chunk.emit_op(Op::DROP, line);
    }

    if argc >= 2 {
        chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    } else {
        push_const(chunk, Value::String(Arc::from("")), line);
        chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    }

    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    set_local_prop(chunk, "name", name_slot, line);
    set_local_prop(chunk, VALUE_KEY, value_slot, line);
}

// ── ADODB.Recordset ───────────────────────────────────────────────────────────

pub fn emit_adodb_recordset_new(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    collections::emit_array_new(chunks, current, 0, line);
    let rows_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        slot
    };
    collections::emit_array_new(chunks, current, 0, line);
    let cols_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        slot
    };
    let chunk = &mut chunks[current];
    set_local_prop(chunk, ROWS_KEY, rows_slot, line);
    set_local_prop(chunk, COL_NAMES_KEY, cols_slot, line);
    set_const_prop(chunk, POS_KEY, Value::F64(0.0), line);
    set_const_prop(chunk, EOF_KEY, Value::Bool(true), line);
    set_const_prop(chunk, RECORD_COUNT_KEY, Value::F64(0.0), line);
    set_const_prop(chunk, ISCLOSED_KEY, Value::Bool(false), line);
}

/// `Recordset.Open(sql, conn)` — populates the existing Recordset receiver
/// in-place (unlike Execute which builds a fresh struct).
pub fn emit_adodb_recordset_open(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let (conn_slot, sql_slot, rs_slot) = {
        let chunk = &mut chunks[current];
        (
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
        )
    };
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, conn_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, sql_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, rs_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, sql_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, conn_slot, line);
    }
    call_import(chunks, current, "wasi:sql/types", "command.new", 2, line);
    let cmd_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        slot
    };
    call_import(
        chunks,
        current,
        "wasi:sql/types",
        "[method]command.execute-reader",
        1,
        line,
    );
    let reader_slot = {
        let chunk = &mut chunks[current];
        let slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        slot
    };
    let (rows_slot, cols_slot, count_slot) = {
        let chunk = &mut chunks[current];
        (
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
        )
    };
    {
        let chunk = &mut chunks[current];
        get_prop_to_local(chunk, reader_slot, ROWS_KEY, rows_slot, line);
        get_prop_to_local(chunk, reader_slot, COL_NAMES_KEY, cols_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, rows_slot, line);
    }
    collections::emit_len(chunks, current, line);
    let eof_slot = {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
        set_object_local_prop(chunk, rs_slot, ROWS_KEY, rows_slot, line);
        set_object_local_prop(chunk, rs_slot, COL_NAMES_KEY, cols_slot, line);
        set_object_const_prop(chunk, rs_slot, POS_KEY, Value::F64(0.0), line);
        set_object_local_prop(chunk, rs_slot, RECORD_COUNT_KEY, count_slot, line);
        let slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
        core_wasm::i32_const(chunk, line, 0);
        vybe_emitter::ops::emit_dyn_eq(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        slot
    };
    let chunk = &mut chunks[current];
    set_object_local_prop(chunk, rs_slot, EOF_KEY, eof_slot, line);
    set_object_const_prop(chunk, rs_slot, ISCLOSED_KEY, Value::Bool(false), line);
    chunk.emit_op(Op::NULL, line);
    let _ = cmd_slot;
}

pub fn emit_adodb_recordset_move_next(chunks: &mut [Chunk], current: usize, line: u32) {
    let (rs_slot, pos_slot, rows_slot) = {
        let chunk = &mut chunks[current];
        (
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
        )
    };
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, rs_slot, line);
        get_prop_to_local(chunk, rs_slot, POS_KEY, pos_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
        chunk.emit_op(Op::I32_FROM_F64, line);
        core_wasm::i32_const(chunk, line, 1);
        chunk.emit_op(Op::I32_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_SET, pos_slot, line);
        get_prop_to_local(chunk, rs_slot, ROWS_KEY, rows_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, rows_slot, line);
    }
    collections::emit_len(chunks, current, line);
    let (len_slot, eof_slot) = {
        let chunk = &mut chunks[current];
        let len_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
        set_object_local_prop(chunk, rs_slot, POS_KEY, pos_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
        vybe_emitter::ops::emit_dyn_ge(chunk, line);
        let eof_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, eof_slot, line);
        (len_slot, eof_slot)
    };
    let chunk = &mut chunks[current];
    set_object_local_prop(chunk, rs_slot, EOF_KEY, eof_slot, line);
    chunk.emit_op(Op::NULL, line);
    let _ = len_slot;
}

pub fn emit_adodb_recordset_move_first(chunks: &mut [Chunk], current: usize, line: u32) {
    let (rs_slot, rows_slot) = {
        let chunk = &mut chunks[current];
        (reserve_slot(chunk), reserve_slot(chunk))
    };
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, rs_slot, line);
        set_object_const_prop(chunk, rs_slot, POS_KEY, Value::F64(0.0), line);
        get_prop_to_local(chunk, rs_slot, ROWS_KEY, rows_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, rows_slot, line);
    }
    collections::emit_len(chunks, current, line);
    let eof_slot = {
        let chunk = &mut chunks[current];
        let len_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
        core_wasm::i32_const(chunk, line, 0);
        vybe_emitter::ops::emit_dyn_eq(chunk, line);
        let eof_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, eof_slot, line);
        eof_slot
    };
    let chunk = &mut chunks[current];
    set_object_local_prop(chunk, rs_slot, EOF_KEY, eof_slot, line);
    chunk.emit_op(Op::NULL, line);
}

/// `Recordset.Fields(nameOrIndex)` — returns `{ value: row[key] }`.
pub fn emit_adodb_recordset_fields(chunks: &mut [Chunk], current: usize, line: u32) {
    let (key_slot, rs_slot, row_slot, value_slot) = {
        let chunk = &mut chunks[current];
        (
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
            reserve_slot(chunk),
        )
    };
    let (rows_key, pos_key) = {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, key_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, rs_slot, line);
        (
            chunk.add_constant(Value::String(Arc::from(ROWS_KEY))),
            chunk.add_constant(Value::String(Arc::from(POS_KEY))),
        )
    };
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, rs_slot, line);
        chunk.emit_op_u16(Op::STRUCT_GET, rows_key, line);
        chunk.emit_op_u16(Op::LOCAL_GET, rs_slot, line);
        chunk.emit_op_u16(Op::STRUCT_GET, pos_key, line);
        chunk.emit_op(Op::I32_FROM_F64, line);
    }
    collections::emit_get(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, row_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, row_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
    }
    collections::emit_get(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    }
    build_field_object(chunks, current, value_slot, line);
}

pub fn emit_adodb_recordset_close(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let rs_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, rs_slot, line);
    set_object_const_prop(chunk, rs_slot, ISCLOSED_KEY, Value::Bool(true), line);
    set_object_const_prop(chunk, rs_slot, EOF_KEY, Value::Bool(true), line);
    chunk.emit_op(Op::NULL, line);
}
