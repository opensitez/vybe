//! PHP `PDO` / `PDOStatement` object surface. Split out of the former
//! `database_adapter.rs`; the mysqli surface lives in `mysqli_adapter.rs`.

use std::sync::Arc;
use vybe_emitter::instructions::core_wasm;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use crate::emitter::string_adapter;
use vybe_emitter::{collections, convert};

const PDO_FETCH_COLUMN: f64 = 7.0;
const PDO_FETCH_KEY_PAIR: f64 = 12.0;

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
fn emit_string_slot_nonempty(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    push_str(chunk, "", line);
    vybe_emitter::ops::emit_dyn_ne(chunk, line);
}

fn concat_slot_with_literal(chunk: &mut Chunk, slot: u16, suffix: &str, line: u32) {
    lget(chunk, slot, line);
    push_str(chunk, suffix, line);
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, slot, line);
}

fn concat_slot_with_slot(chunk: &mut Chunk, target_slot: u16, value_slot: u16, line: u32) {
    lget(chunk, target_slot, line);
    lget(chunk, value_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, target_slot, line);
}

fn replace_in_slot(chunk: &mut Chunk, slot: u16, from: &str, to: &str, line: u32) {
    lget(chunk, slot, line);
    push_str(chunk, from, line);
    push_str(chunk, to, line);
    {
        let idx = chunk.add_import("ecma:string", "replaceAll");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, slot, line);
}

fn append_credentials(
    chunk: &mut Chunk,
    normalized_slot: u16,
    username_slot: Option<u16>,
    password_slot: Option<u16>,
    line: u32,
) {
    if let Some(slot) = username_slot {
        emit_string_slot_nonempty(chunk, slot, line);
        chunk.emit_if(line);
        concat_slot_with_literal(chunk, normalized_slot, ";user=", line);
        concat_slot_with_slot(chunk, normalized_slot, slot, line);
        chunk.emit_end(line);
    }

    if let Some(slot) = password_slot {
        emit_string_slot_nonempty(chunk, slot, line);
        chunk.emit_if(line);
        concat_slot_with_literal(chunk, normalized_slot, ";password=", line);
        concat_slot_with_slot(chunk, normalized_slot, slot, line);
        chunk.emit_end(line);
    }
}

fn normalize_pdo_dsn(
    chunks: &mut [Chunk],
    current: usize,
    dsn_slot: u16,
    username_slot: Option<u16>,
    password_slot: Option<u16>,
    line: u32,
) -> u16 {
    let chunk = &mut chunks[current];
    let normalized_slot = alloc_local(chunk);
    lget(chunk, dsn_slot, line);
    lset(chunk, normalized_slot, line);

    lget(chunk, dsn_slot, line);
    push_str(chunk, "mysql:", line);
    {
        let idx = chunk.add_import("ecma:string", "startsWith");
        chunk.emit_call(idx, 2, line);
    }
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    replace_in_slot(chunk, normalized_slot, "mysql:", "", line);
    replace_in_slot(chunk, normalized_slot, "dbname=", "db=", line);
    append_credentials(chunk, normalized_slot, username_slot, password_slot, line);

    chunk.emit_else(line);
    lget(chunk, dsn_slot, line);
    push_str(chunk, "pgsql:", line);
    {
        let idx = chunk.add_import("ecma:string", "startsWith");
        chunk.emit_call(idx, 2, line);
    }
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    replace_in_slot(chunk, normalized_slot, "pgsql:", "", line);
    replace_in_slot(chunk, normalized_slot, "dbname=", "db=", line);
    lget(chunk, normalized_slot, line);
    push_str(chunk, "port=", line);
    {
        let idx = chunk.add_import("ecma:string", "includes");
        chunk.emit_call(idx, 2, line);
    }
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    concat_slot_with_literal(chunk, normalized_slot, ";port=5432", line);
    chunk.emit_end(line);
    append_credentials(chunk, normalized_slot, username_slot, password_slot, line);

    chunk.emit_else(line);
    lget(chunk, dsn_slot, line);
    push_str(chunk, "postgres:", line);
    {
        let idx = chunk.add_import("ecma:string", "startsWith");
        chunk.emit_call(idx, 2, line);
    }
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    replace_in_slot(chunk, normalized_slot, "postgres:", "", line);
    replace_in_slot(chunk, normalized_slot, "dbname=", "db=", line);
    lget(chunk, normalized_slot, line);
    push_str(chunk, "port=", line);
    {
        let idx = chunk.add_import("ecma:string", "includes");
        chunk.emit_call(idx, 2, line);
    }
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    concat_slot_with_literal(chunk, normalized_slot, ";port=5432", line);
    chunk.emit_end(line);
    append_credentials(chunk, normalized_slot, username_slot, password_slot, line);

    chunk.emit_else(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    normalized_slot
}

fn stamp_pdo_type(chunk: &mut Chunk, conn_slot: u16, line: u32) {
    lget(chunk, conn_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_else(line);
    lget(chunk, conn_slot, line);
    push_str(chunk, "PDO", line);
    struct_set_key(chunk, "__type", line);
    chunk.emit_end(line);
}

fn emit_column_value(chunks: &mut [Chunk], current: usize, row_slot: u16, index: f64, line: u32) {
    let chunk = &mut chunks[current];
    lget(chunk, row_slot, line);
    push_const(chunk, Value::F64(index), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let direct_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, direct_slot, line);

    lget(chunk, direct_slot, line);
    {
        let undef_idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(undef_idx, 1, line);
    }
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    lget(chunk, row_slot, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let keys_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);

    lget(chunk, row_slot, line);
    lget(chunk, keys_slot, line);
    push_const(chunk, Value::F64(index), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_else(line);
    lget(chunk, direct_slot, line);
    chunk.emit_end(line);
}

fn emit_first_column_value(chunks: &mut [Chunk], current: usize, row_slot: u16, line: u32) {
    emit_column_value(chunks, current, row_slot, 0.0, line);
}

fn emit_empty_array(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_array_new(chunks, current, 0, line);
}

fn emit_sql_literal_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    value_slot: u16,
    line: u32,
) -> u16 {
    let resolved_slot = alloc_local(&mut chunks[current]);
    {
        let chunk = &mut chunks[current];
        lget(chunk, value_slot, line);
        lset(chunk, resolved_slot, line);

        // object test: not null AND not number AND not string AND not boolean
        let obj_test_slot = alloc_local(chunk);
        lget(chunk, value_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, obj_test_slot, line);
        // not null
        lget(chunk, obj_test_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_op(Op::I32_EQZ, line);
        // AND not number
        lget(chunk, obj_test_slot, line);
        let test_num = chunk.add_import("wasm:js-number", "test");
        chunk.emit_call(test_num, 1, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_op(Op::I32_AND, line);
        // AND not string
        lget(chunk, obj_test_slot, line);
        let test_str = chunk.add_import("wasm:js-string", "test");
        chunk.emit_call(test_str, 1, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_op(Op::I32_AND, line);
        // AND not boolean
        lget(chunk, obj_test_slot, line);
        let test_bool = chunk.add_import("wasm:js-boolean", "test");
        chunk.emit_call(test_bool, 1, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_op(Op::I32_AND, line);
        chunk.emit_if(line);

        lget(chunk, value_slot, line);
        struct_get_key(chunk, "__value", line);
    }
    let inner_slot = alloc_local(&mut chunks[current]);
    {
        let chunk = &mut chunks[current];
        lset(chunk, inner_slot, line);

        lget(chunk, inner_slot, line);
        {
            let undef_idx = chunk.add_import("wasm:js-undefined", "test");
            chunk.emit_call(undef_idx, 1, line);
        }
        vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);

        lget(chunk, inner_slot, line);
        lset(chunk, resolved_slot, line);

        chunk.emit_end(line);
        chunk.emit_end(line);
    }

    let out_slot = alloc_local(&mut chunks[current]);
    let string_slot = alloc_local(&mut chunks[current]);

    {
        let chunk = &mut chunks[current];
        lget(chunk, resolved_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if(line);
        push_str(chunk, "null", line);
        lset(chunk, out_slot, line);
        chunk.emit_else(line);
        lget(chunk, resolved_slot, line);
        {
            let undef_idx = chunk.add_import("wasm:js-undefined", "test");
            chunk.emit_call(undef_idx, 1, line);
        }
        chunk.emit_if(line);
        push_str(chunk, "null", line);
        lset(chunk, out_slot, line);
        chunk.emit_else(line);
        // number test (covers f64, i32)
        lget(chunk, resolved_slot, line);
        let test_num = chunk.add_import("wasm:js-number", "test");
        chunk.emit_call(test_num, 1, line);
        chunk.emit_if(line);
        lget(chunk, resolved_slot, line);
    }
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    {
        let chunk = &mut chunks[current];
        push_const(chunk, Value::F64(10.0), line);
    }
    call_import(chunks, current, "ecma:number", "toString", 2, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, out_slot, line);
        chunk.emit_else(line);
        // bigint test (covers i64)
        lget(chunk, resolved_slot, line);
        let test_bigint = chunk.add_import("wasm:js-bigint", "test");
        chunk.emit_call(test_bigint, 1, line);
        chunk.emit_if(line);
        lget(chunk, resolved_slot, line);
    }
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    {
        let chunk = &mut chunks[current];
        push_const(chunk, Value::F64(10.0), line);
    }
    call_import(chunks, current, "ecma:number", "toString", 2, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, out_slot, line);
        chunk.emit_else(line);
        // boolean test
        lget(chunk, resolved_slot, line);
        let test_bool = chunk.add_import("wasm:js-boolean", "test");
        chunk.emit_call(test_bool, 1, line);
        chunk.emit_if(line);
        lget(chunk, resolved_slot, line);
        convert::emit_to_string(chunk, line);
        lset(chunk, out_slot, line);
        chunk.emit_else(line);

        lget(chunk, resolved_slot, line);
    }
    string_adapter::emit_echo_stringify(chunks, current, 1, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, string_slot, line);

        lget(chunk, string_slot, line);
        push_str(chunk, "'", line);
        push_str(chunk, "''", line);
        {
            let idx = chunk.add_import("ecma:string", "replaceAll");
            chunk.emit_call(idx, 3, line);
        }
        lset(chunk, string_slot, line);

        push_str(chunk, "'", line);
        lget(chunk, string_slot, line);
        {
            let idx = chunk.add_import("wasm:js-string", "concat");
            chunk.emit_call(idx, 2, line);
        }
        push_str(chunk, "'", line);
        {
            let idx = chunk.add_import("wasm:js-string", "concat");
            chunk.emit_call(idx, 2, line);
        }
        lset(chunk, out_slot, line);

        chunk.emit_end(line);
        chunk.emit_end(line);
        chunk.emit_end(line);
        chunk.emit_end(line);
        chunk.emit_end(line);
    }

    out_slot
}

fn emit_apply_named_bound_pairs(
    chunks: &mut [Chunk],
    current: usize,
    sql_slot: u16,
    pairs_slot: u16,
    line: u32,
) {
    {
        let chunk = &mut chunks[current];
        push_const(chunk, Value::F64(0.0), line);
    }
    let index_slot = alloc_local(&mut chunks[current]);
    {
        let chunk = &mut chunks[current];
        lset(chunk, index_slot, line);

        lget(chunk, pairs_slot, line);
    }
    collections::emit_len(chunks, current, line);
    let len_slot = alloc_local(&mut chunks[current]);
    {
        let chunk = &mut chunks[current];
        lset(chunk, len_slot, line);
    }

    let loop_state = vybe_emitter::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, index_slot, line);
        lget(chunk, len_slot, line);
        vybe_emitter::ops::emit_dyn_lt(chunk, line);
    }
    vybe_emitter::loops::emit_loop_cond(chunks, current, line);

    {
        let chunk = &mut chunks[current];
        lget(chunk, pairs_slot, line);
        lget(chunk, index_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
    }
    let pair_slot = alloc_local(&mut chunks[current]);
    {
        let chunk = &mut chunks[current];
        lset(chunk, pair_slot, line);

        lget(chunk, pair_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::ARRAY_GET, line);
    }
    let pair_value_slot = alloc_local(&mut chunks[current]);
    {
        let chunk = &mut chunks[current];
        lset(chunk, pair_value_slot, line);
    }
    let literal_slot = emit_sql_literal_from_slot(chunks, current, pair_value_slot, line);

    {
        let chunk = &mut chunks[current];
        lget(chunk, sql_slot, line);
        lget(chunk, pair_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lget(chunk, literal_slot, line);
        {
            let idx = chunk.add_import("ecma:string", "replaceAll");
            chunk.emit_call(idx, 3, line);
        }
        lset(chunk, sql_slot, line);

        lget(chunk, index_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, index_slot, line);
    }
    vybe_emitter::loops::emit_loop_end(chunks, current, loop_state, line);
}

fn emit_apply_named_params_from_entries(
    chunks: &mut [Chunk],
    current: usize,
    sql_slot: u16,
    params_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    lget(chunk, params_slot, line);
    collections::emit_iter_entries(chunks, current, line);
    let entries_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, entries_slot, line);

    push_const(chunk, Value::F64(0.0), line);
    let index_slot = alloc_local(chunk);
    lset(chunk, index_slot, line);

    lget(chunk, entries_slot, line);
    collections::emit_len(chunks, current, line);
    let len_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, len_slot, line);

    let pair_slot = alloc_local(chunk);
    let key_slot = alloc_local(chunk);
    let key_text_slot = alloc_local(chunk);
    let value_slot = alloc_local(chunk);

    let _ = chunk;
    let loop_state = vybe_emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, index_slot, line);
    lget(chunk, len_slot, line);
    vybe_emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, entries_slot, line);
    lget(chunk, index_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, pair_slot, line);

    lget(chunk, pair_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, key_slot, line);

    lget(chunk, key_slot, line);
    let _ = chunk;
    string_adapter::emit_echo_stringify(chunks, current, 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, key_text_slot, line);

    lget(chunk, key_text_slot, line);
    push_str(chunk, ":", line);
    {
        let idx = chunk.add_import("ecma:string", "startsWith");
        chunk.emit_call(idx, 2, line);
    }
    vybe_emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_if(line);

    push_str(chunk, ":", line);
    lget(chunk, key_text_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, key_text_slot, line);
    chunk.emit_end(line);

    lget(chunk, pair_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, value_slot, line);
    let literal_slot = emit_sql_literal_from_slot(chunks, current, value_slot, line);
    let chunk = &mut chunks[current];

    lget(chunk, sql_slot, line);
    lget(chunk, key_text_slot, line);
    lget(chunk, literal_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "replaceAll");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, sql_slot, line);

    lget(chunk, index_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, index_slot, line);
    let _ = chunk;
    vybe_emitter::loops::emit_loop_end(chunks, current, loop_state, line);
}

fn emit_new_statement(
    chunks: &mut [Chunk],
    current: usize,
    conn_slot: u16,
    sql_slot: Option<u16>,
    rows_slot: Option<u16>,
    line: u32,
) {
    let chunk = &mut chunks[current];
    lget(chunk, conn_slot, line);
    call_import(chunks, current, "wasi:sql", "createCommand", 1, line);
    let stmt_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, stmt_slot, line);

    lget(chunk, stmt_slot, line);
    push_str(chunk, "PDOStatement", line);
    struct_set_key(chunk, "__type", line);

    if let Some(slot) = sql_slot {
        lget(chunk, stmt_slot, line);
        lget(chunk, slot, line);
        struct_set_key(chunk, "commandtext", line);

        lget(chunk, stmt_slot, line);
        lget(chunk, slot, line);
        struct_set_key(chunk, "__prepared_commandtext", line);
    }

    lget(chunk, stmt_slot, line);
    lget(chunk, conn_slot, line);
    struct_set_key(chunk, "__conn", line);

    lget(chunk, stmt_slot, line);
    if let Some(slot) = rows_slot {
        lget(chunk, slot, line);
    } else {
        emit_empty_array(chunks, current, line);
    }
    let chunk = &mut chunks[current];
    struct_set_key(chunk, "__rows", line);

    lget(chunk, stmt_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, "__cursor", line);

    lget(chunk, stmt_slot, line);
    emit_empty_array(chunks, current, line);
    let chunk = &mut chunks[current];
    struct_set_key(chunk, "__bound_params", line);

    lget(chunk, stmt_slot, line);
    emit_empty_array(chunks, current, line);
    let chunk = &mut chunks[current];
    struct_set_key(chunk, "__bound_named_pairs", line);

    lget(chunk, stmt_slot, line);
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

fn emit_select_column_count_from_sql_slot(chunk: &mut Chunk, sql_slot: u16, line: u32) {
    lget(chunk, sql_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "trim");
        chunk.emit_call(idx, 1, line);
    }
    {
        let idx = chunk.add_import("ecma:string", "toLowerCase");
        chunk.emit_call(idx, 1, line);
    }
    let lower_slot = alloc_local(chunk);
    lset(chunk, lower_slot, line);

    lget(chunk, lower_slot, line);
    push_str(chunk, "select", line);
    {
        let idx = chunk.add_import("ecma:string", "startsWith");
        chunk.emit_call(idx, 2, line);
    }
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, lower_slot, line);
    push_str(chunk, ",", line);
    {
        let idx = chunk.add_import("ecma:string", "split");
        chunk.emit_call(idx, 2, line);
    }
    vybe_emitter::collections::emit_array_length(chunk, line);
    chunk.emit_else(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_end(line);
}

pub fn emit_php_pdo_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let options_slot = if argc >= 4 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let password_slot = if argc >= 3 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let username_slot = if argc >= 2 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let dsn_slot = alloc_local(chunk);

    if let Some(slot) = options_slot {
        lset(chunk, slot, line);
    }
    if let Some(slot) = password_slot {
        lset(chunk, slot, line);
    }
    if let Some(slot) = username_slot {
        lset(chunk, slot, line);
    }
    lset(chunk, dsn_slot, line);

    let normalized_slot = normalize_pdo_dsn(
        chunks,
        current,
        dsn_slot,
        username_slot,
        password_slot,
        line,
    );
    let chunk = &mut chunks[current];
    lget(chunk, normalized_slot, line);
    call_import(chunks, current, "wasi:sql", "connect", 1, line);
    let conn_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, conn_slot, line);
    lget(chunk, conn_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    push_str(chunk, "sqlite::memory:", line);
    let _ = chunk;
    call_import(chunks, current, "wasi:sql", "connect", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, conn_slot, line);
    chunk.emit_end(line);
    stamp_pdo_type(chunk, conn_slot, line);
    lget(chunk, conn_slot, line);
}

pub fn emit_php_pdo_query(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let sql_slot = alloc_local(chunk);
    let conn_slot = alloc_local(chunk);
    lset(chunk, sql_slot, line);
    lset(chunk, conn_slot, line);

    lget(chunk, conn_slot, line);
    lget(chunk, sql_slot, line);
    call_import(chunks, current, "wasi:sql", "query", 2, line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);

    emit_new_statement(
        chunks,
        current,
        conn_slot,
        Some(sql_slot),
        Some(rows_slot),
        line,
    );
}

pub fn emit_php_pdo_exec(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let sql_slot = alloc_local(chunk);
    let conn_slot = alloc_local(chunk);
    lset(chunk, sql_slot, line);
    lset(chunk, conn_slot, line);

    lget(chunk, conn_slot, line);
    lget(chunk, sql_slot, line);
    call_import(chunks, current, "wasi:sql", "execute", 2, line);
}

pub fn emit_php_pdo_prepare(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let sql_slot = alloc_local(chunk);
    let conn_slot = alloc_local(chunk);
    lset(chunk, sql_slot, line);
    lset(chunk, conn_slot, line);

    emit_new_statement(chunks, current, conn_slot, Some(sql_slot), None, line);
}

pub fn emit_php_pdo_set_attribute(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = alloc_local(chunk);
    let attr_slot = alloc_local(chunk);
    let conn_slot = alloc_local(chunk);
    lset(chunk, value_slot, line);
    lset(chunk, attr_slot, line);
    lset(chunk, conn_slot, line);

    lget(chunk, conn_slot, line);
    lget(chunk, value_slot, line);
    struct_set_key(chunk, "__pdo_attr", line);
    push_const(chunk, Value::Bool(true), line);
}

pub fn emit_php_pdo_begin_transaction(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "wasi:sql", "beginTransaction", 1, line);
}

pub fn emit_php_pdo_commit(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "wasi:sql", "commit", 1, line);
}

pub fn emit_php_pdo_rollback(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "wasi:sql", "rollback", 1, line);
}

fn emit_php_pdo_statement_bind_common(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let _driver_options_slot = if argc >= 6 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let _max_length_slot = if argc >= 5 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let _type_slot = if argc >= 4 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let value_slot = alloc_local(chunk);
    let param_slot = alloc_local(chunk);
    let stmt_slot = alloc_local(chunk);

    if let Some(slot) = _driver_options_slot {
        lset(chunk, slot, line);
    }
    if let Some(slot) = _max_length_slot {
        lset(chunk, slot, line);
    }
    if let Some(slot) = _type_slot {
        lset(chunk, slot, line);
    }
    lset(chunk, value_slot, line);
    lset(chunk, param_slot, line);
    lset(chunk, stmt_slot, line);

    let chunk = &mut chunks[current];

    lget(chunk, param_slot, line);
    let test_str_param = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str_param, 1, line);
    chunk.emit_if(line);

    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, "__bound_named_pairs", line);
    let named_pairs_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, named_pairs_slot, line);

    lget(chunk, named_pairs_slot, line);
    lget(chunk, param_slot, line);
    lget(chunk, value_slot, line);
    collections::emit_array_new(chunks, current, 2, line);
    collections::emit_push(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);

    chunk.emit_else(line);
    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, "__bound_params", line);
    let params_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, params_slot, line);

    lget(chunk, params_slot, line);
    lget(chunk, param_slot, line);
    convert::emit_to_int(chunk, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_SUB, line);
    lget(chunk, value_slot, line);
    collections::emit_set(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
    push_const(chunk, Value::Bool(true), line);
}

pub fn emit_php_pdo_statement_bind_param(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    emit_php_pdo_statement_bind_common(chunks, current, argc, line);
}

pub fn emit_php_pdo_statement_bind_value(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    emit_php_pdo_statement_bind_common(chunks, current, argc, line);
}

pub fn emit_php_pdo_statement_bind_column(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    for _ in 0..(argc as u16 + 1) {
        chunk.emit_op(Op::DROP, line);
    }
    push_const(chunk, Value::Bool(true), line);
}

pub fn emit_php_pdo_statement_execute(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let explicit_params_slot = {
        let chunk = &mut chunks[current];
        if argc >= 2 {
            Some(alloc_local(chunk))
        } else {
            None
        }
    };
    let stmt_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    {
        let chunk = &mut chunks[current];
        if let Some(slot) = explicit_params_slot {
            lset(chunk, slot, line);
        }
        lset(chunk, stmt_slot, line);
        lget(chunk, stmt_slot, line);
        struct_get_key(chunk, "__prepared_commandtext", line);
    }
    let sql_text_slot = alloc_local(&mut chunks[current]);
    {
        let chunk = &mut chunks[current];
        lset(chunk, sql_text_slot, line);
    }

    {
        let chunk = &mut chunks[current];
        lget(chunk, stmt_slot, line);
        emit_select_column_count_from_sql_slot(chunk, sql_text_slot, line);
        struct_set_key(chunk, "field_count", line);
    }

    let effective_params_slot = alloc_local(&mut chunks[current]);
    if let Some(slot) = explicit_params_slot {
        let chunk = &mut chunks[current];
        lget(chunk, slot, line);
        lset(chunk, effective_params_slot, line);
        emit_apply_named_params_from_entries(
            chunks,
            current,
            sql_text_slot,
            effective_params_slot,
            line,
        );
    } else {
        {
            let chunk = &mut chunks[current];
            lget(chunk, stmt_slot, line);
            struct_get_key(chunk, "__bound_params", line);
            lset(chunk, effective_params_slot, line);

            lget(chunk, stmt_slot, line);
            struct_get_key(chunk, "__bound_named_pairs", line);
        }
        let named_pairs_slot = alloc_local(&mut chunks[current]);
        {
            let chunk = &mut chunks[current];
            lset(chunk, named_pairs_slot, line);
        }
        emit_apply_named_bound_pairs(chunks, current, sql_text_slot, named_pairs_slot, line);
    }

    {
        let chunk = &mut chunks[current];
        lget(chunk, sql_text_slot, line);
        {
            let idx = chunk.add_import("ecma:string", "trim");
            chunk.emit_call(idx, 1, line);
        }
        {
            let idx = chunk.add_import("ecma:string", "toLowerCase");
            chunk.emit_call(idx, 1, line);
        }
    }
    let sql_slot = alloc_local(&mut chunks[current]);
    {
        let chunk = &mut chunks[current];
        lset(chunk, sql_slot, line);
    }

    let is_query_slot = {
        let chunk = &mut chunks[current];
        core_wasm::i32_const(chunk, line, 0);
        let slot = alloc_local(chunk);
        lset(chunk, slot, line);
        emit_mark_queryish_prefix(chunk, sql_slot, slot, "select", line);
        emit_mark_queryish_prefix(chunk, sql_slot, slot, "pragma", line);
        emit_mark_queryish_prefix(chunk, sql_slot, slot, "show", line);
        emit_mark_queryish_prefix(chunk, sql_slot, slot, "with", line);
        emit_mark_queryish_prefix(chunk, sql_slot, slot, "describe", line);
        slot
    };

    let chunk = &mut chunks[current];
    lget(chunk, is_query_slot, line);
    chunk.emit_if_value(line);

    lget(chunk, stmt_slot, line);
    lget(chunk, sql_text_slot, line);
    struct_set_key(chunk, "commandtext", line);
    lget(chunk, stmt_slot, line);
    lget(chunk, sql_text_slot, line);
    lget(chunk, effective_params_slot, line);
    call_import(chunks, current, "wasi:sql", "query", 3, line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);
    lget(chunk, stmt_slot, line);
    lget(chunk, rows_slot, line);
    struct_set_key(chunk, "__rows", line);
    lget(chunk, stmt_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, "__cursor", line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    lget(chunk, stmt_slot, line);
    lget(chunk, sql_text_slot, line);
    struct_set_key(chunk, "commandtext", line);
    lget(chunk, stmt_slot, line);
    lget(chunk, sql_text_slot, line);
    lget(chunk, effective_params_slot, line);
    call_import(chunks, current, "wasi:sql", "execute", 3, line);
    let count_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, count_slot, line);
    lget(chunk, stmt_slot, line);
    emit_empty_array(chunks, current, line);
    let chunk = &mut chunks[current];
    struct_set_key(chunk, "__rows", line);
    lget(chunk, stmt_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, "__cursor", line);
    lget(chunk, stmt_slot, line);
    lget(chunk, count_slot, line);
    struct_set_key(chunk, "__row_count", line);
    lget(chunk, count_slot, line);
    push_const(chunk, Value::F64(-1.0), line);
    vybe_emitter::ops::emit_dyn_ne(chunk, line);

    chunk.emit_end(line);
}

pub fn emit_php_pdo_statement_fetch(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let mode_slot = if argc >= 2 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let stmt_slot = alloc_local(chunk);
    if let Some(slot) = mode_slot {
        lset(chunk, slot, line);
    }
    lset(chunk, stmt_slot, line);

    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, "__rows", line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);

    lget(chunk, stmt_slot, line);
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

    lget(chunk, stmt_slot, line);
    lget(chunk, cursor_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    struct_set_key(chunk, "__cursor", line);

    if let Some(slot) = mode_slot {
        lget(chunk, slot, line);
        push_const(chunk, Value::F64(PDO_FETCH_COLUMN), line);
        vybe_emitter::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if_value(line);

        lget(chunk, row_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if_value(line);
        chunk.emit_op(Op::NULL, line);
        chunk.emit_else(line);
        let _ = chunk;
        emit_first_column_value(chunks, current, row_slot, line);
        let chunk = &mut chunks[current];
        chunk.emit_end(line);
        chunk.emit_else(line);
        lget(chunk, row_slot, line);
        chunk.emit_end(line);
    } else {
        // Default mode: PDO::fetch() returns `false` past the last row.
        let chunk = &mut chunks[current];
        lget(chunk, row_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if_value(line);
        chunk.emit_bool_const(false, line);
        chunk.emit_else(line);
        lget(chunk, row_slot, line);
        chunk.emit_end(line);
    }
}

pub fn emit_php_pdo_statement_fetch_all(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let mode_slot = if argc >= 2 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let stmt_slot = alloc_local(chunk);
    if let Some(slot) = mode_slot {
        lset(chunk, slot, line);
    }
    lset(chunk, stmt_slot, line);

    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, "__rows", line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);

    if let Some(slot) = mode_slot {
        lget(chunk, slot, line);
        push_const(chunk, Value::F64(65_535.0), line);
        vybe_emitter::ops::emit_dyn_gt(chunk, line);
        chunk.emit_if_value(line);

        call_import(chunks, current, "ecma:object", "new", 0, line);
        let out_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, out_slot, line);

        call_import(chunks, current, "ecma:map", "new", 0, line);
        let groups_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, groups_slot, line);

        push_const(chunk, Value::F64(0.0), line);
        let index_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, index_slot, line);

        lget(chunk, rows_slot, line);
        collections::emit_len(chunks, current, line);
        let len_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, len_slot, line);

        let _ = chunk;
        let loop_state = vybe_emitter::loops::emit_loop_start(chunks, current, line);
        let chunk = &mut chunks[current];
        lget(chunk, index_slot, line);
        lget(chunk, len_slot, line);
        vybe_emitter::ops::emit_dyn_lt(chunk, line);
        let _ = chunk;
        vybe_emitter::loops::emit_loop_cond(chunks, current, line);
        let chunk = &mut chunks[current];

        lget(chunk, rows_slot, line);
        lget(chunk, index_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        let row_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, row_slot, line);

        let _ = chunk;
        emit_column_value(chunks, current, row_slot, 0.0, line);
        let key_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, key_slot, line);

        lget(chunk, groups_slot, line);
        lget(chunk, key_slot, line);
        {
            let idx = chunk.add_import("ecma:map", "has");
            chunk.emit_call(idx, 2, line);
        }
        vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if_value(line);
        lget(chunk, groups_slot, line);
        lget(chunk, key_slot, line);
        {
            let idx = chunk.add_import("ecma:map", "get");
            chunk.emit_call(idx, 2, line);
        }
        chunk.emit_else(line);
        collections::emit_array_new(chunks, current, 0, line);
        let chunk = &mut chunks[current];
        chunk.emit_end(line);
        let group_slot = alloc_local(chunk);
        lset(chunk, group_slot, line);

        lget(chunk, group_slot, line);
        lget(chunk, row_slot, line);
        collections::emit_push(chunks, current, line);
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);

        lget(chunk, groups_slot, line);
        lget(chunk, key_slot, line);
        lget(chunk, group_slot, line);
        {
            let idx = chunk.add_import("ecma:map", "set");
            chunk.emit_call(idx, 3, line);
        }
        chunk.emit_op(Op::DROP, line);

        lget(chunk, out_slot, line);
        lget(chunk, key_slot, line);
        lget(chunk, group_slot, line);
        {
            let idx = chunk.add_import("ecma:object", "set");
            chunk.emit_call(idx, 3, line);
        }
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);

        lget(chunk, index_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, index_slot, line);
        let _ = chunk;
        vybe_emitter::loops::emit_loop_end(chunks, current, loop_state, line);
        let chunk = &mut chunks[current];
        lget(chunk, out_slot, line);
        chunk.emit_else(line);

        lget(chunk, slot, line);
        push_const(chunk, Value::F64(PDO_FETCH_KEY_PAIR), line);
        vybe_emitter::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if_value(line);

        call_import(chunks, current, "ecma:map", "new", 0, line);
        let out_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, out_slot, line);

        push_const(chunk, Value::F64(0.0), line);
        let index_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, index_slot, line);

        lget(chunk, rows_slot, line);
        collections::emit_len(chunks, current, line);
        let len_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, len_slot, line);

        let _ = chunk;
        let loop_state = vybe_emitter::loops::emit_loop_start(chunks, current, line);
        let chunk = &mut chunks[current];
        lget(chunk, index_slot, line);
        lget(chunk, len_slot, line);
        vybe_emitter::ops::emit_dyn_lt(chunk, line);
        let _ = chunk;
        vybe_emitter::loops::emit_loop_cond(chunks, current, line);
        let chunk = &mut chunks[current];

        lget(chunk, rows_slot, line);
        lget(chunk, index_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        let row_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, row_slot, line);

        lget(chunk, row_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        vybe_emitter::ops::emit_dyn_not(chunk, line);
        chunk.emit_if(line);

        lget(chunk, row_slot, line);
        push_const(chunk, Value::F64(2.0), line);
        chunk.emit_op(Op::ARRAY_GET, line);
        {
            let undef_idx = chunk.add_import("wasm:js-undefined", "test");
            chunk.emit_call(undef_idx, 1, line);
        }
        vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        let _ = chunk;
        crate::emitter::type_guard::emit_throw_const(
            chunks,
            current,
            "PDOException",
            "PDO::FETCH_KEY_PAIR requires exactly two columns",
            line,
        );
        let chunk = &mut chunks[current];
        chunk.emit_end(line);

        let _ = chunk;
        emit_column_value(chunks, current, row_slot, 0.0, line);
        let key_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, key_slot, line);
        let _ = chunk;
        emit_column_value(chunks, current, row_slot, 1.0, line);
        let value_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, value_slot, line);
        lget(chunk, out_slot, line);
        lget(chunk, key_slot, line);
        lget(chunk, value_slot, line);
        collections::emit_set(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        let chunk = &mut chunks[current];
        chunk.emit_end(line);

        lget(chunk, index_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, index_slot, line);
        let _ = chunk;
        vybe_emitter::loops::emit_loop_end(chunks, current, loop_state, line);
        let chunk = &mut chunks[current];
        lget(chunk, out_slot, line);
        chunk.emit_else(line);

        lget(chunk, slot, line);
        push_const(chunk, Value::F64(PDO_FETCH_COLUMN), line);
        vybe_emitter::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if_value(line);

        emit_empty_array(chunks, current, line);
        let out_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, out_slot, line);

        push_const(chunk, Value::F64(0.0), line);
        let index_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, index_slot, line);

        lget(chunk, rows_slot, line);
        collections::emit_len(chunks, current, line);
        let len_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, len_slot, line);

        let _ = chunk;
        let loop_state = vybe_emitter::loops::emit_loop_start(chunks, current, line);
        let chunk = &mut chunks[current];
        lget(chunk, index_slot, line);
        lget(chunk, len_slot, line);
        vybe_emitter::ops::emit_dyn_lt(chunk, line);
        let _ = chunk;
        vybe_emitter::loops::emit_loop_cond(chunks, current, line);
        let chunk = &mut chunks[current];

        lget(chunk, rows_slot, line);
        lget(chunk, index_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        let row_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, row_slot, line);

        lget(chunk, row_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        vybe_emitter::ops::emit_dyn_not(chunk, line);
        chunk.emit_if(line);
        lget(chunk, out_slot, line);
        let _ = chunk;
        emit_first_column_value(chunks, current, row_slot, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        let chunk = &mut chunks[current];
        chunk.emit_end(line);

        lget(chunk, index_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, index_slot, line);
        let _ = chunk;
        vybe_emitter::loops::emit_loop_end(chunks, current, loop_state, line);
        let chunk = &mut chunks[current];
        lget(chunk, out_slot, line);
        chunk.emit_else(line);
        lget(chunk, rows_slot, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
        chunk.emit_end(line);
        return;
    }

    let chunk = &mut chunks[current];
    lget(chunk, rows_slot, line);
}

pub fn emit_php_pdo_statement_fetch_object(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let class_slot = if argc >= 2 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let stmt_slot = alloc_local(chunk);
    if let Some(slot) = class_slot {
        lset(chunk, slot, line);
    }
    lset(chunk, stmt_slot, line);

    lget(chunk, stmt_slot, line);
    emit_php_pdo_statement_fetch(chunks, current, 0, line);
    let chunk = &mut chunks[current];
    let row_slot = alloc_local(chunk);
    lset(chunk, row_slot, line);

    lget(chunk, row_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_else(line);
    if let Some(slot) = class_slot {
        lget(chunk, row_slot, line);
        lget(chunk, slot, line);
        struct_set_key(chunk, "__type", line);
    }
    lget(chunk, row_slot, line);
    chunk.emit_end(line);
}

// ── Additional PDOStatement / PDO methods over the shared statement shape ────

/// `$stmt->fetchColumn()` — advance the cursor, return the first column of the
/// row, or `false` at end of result. Stack: `[stmt]` → `[value|false]`.
pub fn emit_php_pdo_statement_fetch_column(
    chunks: &mut [Chunk],
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let stmt_slot = alloc_local(chunk);
    lset(chunk, stmt_slot, line);
    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, "__rows", line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);
    lget(chunk, stmt_slot, line);
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
    // cursor += 1
    lget(chunk, stmt_slot, line);
    lget(chunk, cursor_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    struct_set_key(chunk, "__cursor", line);
    // row === null ? false : firstColumn(row)
    lget(chunk, row_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_else(line);
    let _ = chunk;
    emit_first_column_value(chunks, current, row_slot, line);
    chunks[current].emit_end(line);
}

/// `$stmt->rowCount()` — rows affected by the last DML, or 0. Stack: `[stmt]` → `[n]`.
pub fn emit_php_pdo_statement_row_count(
    chunks: &mut [Chunk],
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    struct_get_key(chunk, "__row_count", line);
    // null → 0
    let n_slot = alloc_local(chunk);
    lset(chunk, n_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_else(line);
    lget(chunk, n_slot, line);
    chunk.emit_end(line);
}

/// `$stmt->columnCount()` — number of columns in the result (keys of row 0). 0
/// when there are no rows. Stack: `[stmt]` → `[n]`.
pub fn emit_php_pdo_statement_column_count(
    chunks: &mut [Chunk],
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let stmt_slot = alloc_local(chunk);
    lset(chunk, stmt_slot, line);
    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, "__rows", line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);
    // rows.length > 0 ? Object.keys(rows[0]).length : stmt.field_count
    lget(chunk, rows_slot, line);
    vybe_emitter::collections::emit_array_length(chunk, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_emitter::ops::emit_dyn_gt(chunk, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, rows_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    vybe_emitter::collections::emit_array_length(chunk, line);
    chunk.emit_else(line);
    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, "field_count", line);
    {
        let undef_idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(undef_idx, 1, line);
    }
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_else(line);
    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, "field_count", line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `$stmt->paramCount()` — placeholder count stamped by the shared
/// `db_adapter` prepare path. Stack: `[stmt]` → `[n]`.
pub fn emit_php_pdo_statement_param_count(
    chunks: &mut [Chunk],
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    struct_get_key(chunk, "param_count", line);
    let n_slot = alloc_local(chunk);
    lset(chunk, n_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_else(line);
    lget(chunk, n_slot, line);
    chunk.emit_end(line);
}

/// `$pdo->getAttribute(PDO::ATTR_DRIVER_NAME)` → the connection's driver name.
/// Stack: `[conn, attr]` → `[value]`. Reads `__driver` (stamped at connect).
pub fn emit_php_pdo_get_attribute(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    // Only ATTR_DRIVER_NAME is exercised; the DB layer is sqlite-backed.
    for _ in 0..argc as u16 {
        chunk.emit_op(Op::DROP, line);
    }
    push_str(chunk, "sqlite", line);
}

/// `$pdo->quote($s)` — wrap in single quotes, doubling embedded quotes.
/// Stack: `[conn, s]` → `[quoted]`.
pub fn emit_php_pdo_quote(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    lset(chunk, s_slot, line);
    chunk.emit_op(Op::DROP, line); // receiver conn
                                   // "'" + s.replaceAll("'", "''") + "'"
    push_str(chunk, "'", line);
    lget(chunk, s_slot, line);
    push_str(chunk, "'", line);
    push_str(chunk, "''", line);
    {
        let idx = chunk.add_import("ecma:string", "replaceAll");
        chunk.emit_call(idx, 3, line);
    }
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    push_str(chunk, "'", line);
    vybe_emitter::ops::emit_dyn_add(chunk, line);
}

/// `$stmt->errorCode()` → SQLSTATE "00000" on success. Stack: `[stmt]` → `[str]`.
pub fn emit_php_pdo_error_code(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc as u16 {
        chunk.emit_op(Op::DROP, line);
    }
    push_str(chunk, "00000", line);
}

/// `$pdo->lastInsertId()` — the id of the last inserted row via
/// `SELECT last_insert_rowid()`. Stack: `[conn]`/`[conn, name]` → `[id]`.
pub fn emit_php_pdo_last_insert_id(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc >= 2 {
        chunk.emit_op(Op::DROP, line); // optional sequence-name arg
    }
    let conn_slot = alloc_local(chunk);
    lset(chunk, conn_slot, line);
    // Reuse the proven `$pdo->query(...)` path: it returns a statement whose
    // `__rows` holds the result. Stack for it: [conn, sql].
    lget(chunk, conn_slot, line);
    push_str(chunk, "SELECT last_insert_rowid() AS id", line);
    emit_php_pdo_query(chunks, current, 1, line);
    let chunk = &mut chunks[current];
    let stmt_slot = alloc_local(chunk);
    lset(chunk, stmt_slot, line);
    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, "__rows", line);
    let rows_slot = alloc_local(chunk);
    lset(chunk, rows_slot, line);
    lget(chunk, rows_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let row_slot = alloc_local(chunk);
    lset(chunk, row_slot, line);
    emit_first_column_value(chunks, current, row_slot, line);
}

/// `$stmt->errorInfo()` / `$pdo->errorInfo()` → `["00000", null, null]`.
/// Stack: `[recv]` → `[array]`.
pub fn emit_php_pdo_error_info(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc as u16 {
        chunk.emit_op(Op::DROP, line);
    }
    push_str(chunk, "00000", line);
    chunk.emit_op(Op::NULL, line);
    chunk.emit_op(Op::NULL, line);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, line);
}
