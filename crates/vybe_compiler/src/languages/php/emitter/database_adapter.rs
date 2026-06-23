use crate::emitter::instructions::core_wasm;
use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use crate::emitter::{collections, convert};
use crate::languages::php::emitter::string_adapter;

const PDO_FETCH_COLUMN: f64 = 7.0;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let slot = chunk.local_count;
    chunk.local_count = slot + 1;
    slot
}

fn push_const(chunk: &mut Chunk, value: Value, line: u32) {
    match &value {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_op(Op::NULL, line),
        Value::BigInt(v) => chunk.emit_i64_const(*v, line),
        Value::String(s) => chunk.emit_string_const(&s, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),
        
        _ => {
            let idx = chunk.add_constant(value);
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
    chunk.emit_op(Op::DROP, line);
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
    chunk.emit_op(Op::DROP, line);
}

fn global_get_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::GLOBAL_GET, idx, line);
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

fn emit_string_slot_nonempty(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    push_str(chunk, "", line);
    crate::emitter::ops::emit_dyn_ne(chunk, line);
}

fn concat_slot_with_literal(chunk: &mut Chunk, slot: u16, suffix: &str, line: u32) {
    lget(chunk, slot, line);
    push_str(chunk, suffix, line);
    { let idx = chunk.add_import("wasm:js-string", "concat"); chunk.emit_call(idx, 2, line); }
    lset(chunk, slot, line);
}

fn concat_slot_with_slot(chunk: &mut Chunk, target_slot: u16, value_slot: u16, line: u32) {
    lget(chunk, target_slot, line);
    lget(chunk, value_slot, line);
    { let idx = chunk.add_import("wasm:js-string", "concat"); chunk.emit_call(idx, 2, line); }
    lset(chunk, target_slot, line);
}

fn replace_in_slot(chunk: &mut Chunk, slot: u16, from: &str, to: &str, line: u32) {
    lget(chunk, slot, line);
    push_str(chunk, from, line);
    push_str(chunk, to, line);
    { let idx = chunk.add_import("ecma:string", "replaceAll"); chunk.emit_call(idx, 3, line); }
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
    { let idx = chunk.add_import("ecma:string", "startsWith"); chunk.emit_call(idx, 2, line); }
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    replace_in_slot(chunk, normalized_slot, "mysql:", "", line);
    replace_in_slot(chunk, normalized_slot, "dbname=", "db=", line);
    append_credentials(chunk, normalized_slot, username_slot, password_slot, line);

    chunk.emit_else(line);
    lget(chunk, dsn_slot, line);
    push_str(chunk, "pgsql:", line);
    { let idx = chunk.add_import("ecma:string", "startsWith"); chunk.emit_call(idx, 2, line); }
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    replace_in_slot(chunk, normalized_slot, "pgsql:", "", line);
    replace_in_slot(chunk, normalized_slot, "dbname=", "db=", line);
    lget(chunk, normalized_slot, line);
    push_str(chunk, "port=", line);
    { let idx = chunk.add_import("ecma:string", "includes"); chunk.emit_call(idx, 2, line); }
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    concat_slot_with_literal(chunk, normalized_slot, ";port=5432", line);
    chunk.emit_end(line);
    append_credentials(chunk, normalized_slot, username_slot, password_slot, line);

    chunk.emit_else(line);
    lget(chunk, dsn_slot, line);
    push_str(chunk, "postgres:", line);
    { let idx = chunk.add_import("ecma:string", "startsWith"); chunk.emit_call(idx, 2, line); }
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    replace_in_slot(chunk, normalized_slot, "postgres:", "", line);
    replace_in_slot(chunk, normalized_slot, "dbname=", "db=", line);
    lget(chunk, normalized_slot, line);
    push_str(chunk, "port=", line);
    { let idx = chunk.add_import("ecma:string", "includes"); chunk.emit_call(idx, 2, line); }
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
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

fn emit_first_column_value(chunks: &mut [Chunk], current: usize, row_slot: u16, line: u32) {
    let chunk = &mut chunks[current];
    lget(chunk, row_slot, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let keys_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);

    lget(chunk, row_slot, line);
    lget(chunk, keys_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_GET, line);
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
        chunk.emit_op(Op::DROP, line);
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
        { let undef_idx = chunk.add_import("wasm:js-undefined", "test"); chunk.emit_call(undef_idx, 1, line); }
        crate::emitter::ops::emit_dyn_to_bool(chunk, line);
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
        { let undef_idx = chunk.add_import("wasm:js-undefined", "test"); chunk.emit_call(undef_idx, 1, line); }
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
        { let idx = chunk.add_import("ecma:string", "replaceAll"); chunk.emit_call(idx, 3, line); }
        lset(chunk, string_slot, line);

        push_str(chunk, "'", line);
        lget(chunk, string_slot, line);
        { let idx = chunk.add_import("wasm:js-string", "concat"); chunk.emit_call(idx, 2, line); }
        push_str(chunk, "'", line);
        { let idx = chunk.add_import("wasm:js-string", "concat"); chunk.emit_call(idx, 2, line); }
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

    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, index_slot, line);
        lget(chunk, len_slot, line);
        crate::emitter::ops::emit_dyn_lt(chunk, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);

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
        { let idx = chunk.add_import("ecma:string", "replaceAll"); chunk.emit_call(idx, 3, line); }
        lset(chunk, sql_slot, line);

        lget(chunk, index_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, index_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
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
    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, index_slot, line);
    lget(chunk, len_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
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
    { let idx = chunk.add_import("ecma:string", "startsWith"); chunk.emit_call(idx, 2, line); }
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_if(line);

    push_str(chunk, ":", line);
    lget(chunk, key_text_slot, line);
    { let idx = chunk.add_import("wasm:js-string", "concat"); chunk.emit_call(idx, 2, line); }
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
    { let idx = chunk.add_import("ecma:string", "replaceAll"); chunk.emit_call(idx, 3, line); }
    lset(chunk, sql_slot, line);

    lget(chunk, index_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, index_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
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
    { let idx = chunk.add_import("ecma:string", "startsWith"); chunk.emit_call(idx, 2, line); }
    chunk.emit_if(line);
    core_wasm::i32_const(chunk, line, 1);
    lset(chunk, is_query_slot, line);
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
        { let idx = chunk.add_import("ecma:string", "trim"); chunk.emit_call(idx, 1, line); }
        { let idx = chunk.add_import("ecma:string", "toLowerCase"); chunk.emit_call(idx, 1, line); }
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
    crate::emitter::ops::emit_dyn_ne(chunk, line);

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
        crate::emitter::ops::emit_dyn_eq(chunk, line);
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
        let chunk = &mut chunks[current];
        lget(chunk, row_slot, line);
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
        push_const(chunk, Value::F64(PDO_FETCH_COLUMN), line);
        crate::emitter::ops::emit_dyn_eq(chunk, line);
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
        let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
        let chunk = &mut chunks[current];
        lget(chunk, index_slot, line);
        lget(chunk, len_slot, line);
        crate::emitter::ops::emit_dyn_lt(chunk, line);
        let _ = chunk;
        crate::emitter::loops::emit_loop_cond(chunks, current, line);
        let chunk = &mut chunks[current];

        lget(chunk, rows_slot, line);
        lget(chunk, index_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        let row_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, row_slot, line);

        lget(chunk, row_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        crate::emitter::ops::emit_dyn_not(chunk, line);
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
        crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
        let chunk = &mut chunks[current];
        lget(chunk, out_slot, line);
        chunk.emit_else(line);
        lget(chunk, rows_slot, line);
        chunk.emit_end(line);
        return;
    }

    let chunk = &mut chunks[current];
    lget(chunk, rows_slot, line);
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

    call_import(chunks, current, "ecma:object", "new", 0, line);
    let conn_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, conn_slot, line);

    // Stamp a mysqli-like marker and a default zero errno field so
    // property probes in `wpdb::db_connect()` have expected shape.
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
        { let idx = chunk.add_import("wasm:js-string", "concat"); chunk.emit_call(idx, 2, line); }
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
        { let idx = chunk.add_import("wasm:js-string", "concat"); chunk.emit_call(idx, 2, line); }
        lget(chunk, pass_slot, line);
        { let idx = chunk.add_import("wasm:js-string", "concat"); chunk.emit_call(idx, 2, line); }
        lset(chunk, url_slot, line);
        chunk.emit_end(line);
    }

    // Append @host
    lget(chunk, url_slot, line);
    push_str(chunk, "@", line);
    { let idx = chunk.add_import("wasm:js-string", "concat"); chunk.emit_call(idx, 2, line); }
    if let Some(slot) = host_slot {
        lget(chunk, slot, line);
    } else {
        push_str(chunk, "localhost", line);
    }
    { let idx = chunk.add_import("wasm:js-string", "concat"); chunk.emit_call(idx, 2, line); }
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
        { let idx = chunk.add_import("wasm:js-string", "concat"); chunk.emit_call(idx, 2, line); }
        lget(chunk, slot, line);
        { let idx = chunk.add_import("wasm:js-string", "concat"); chunk.emit_call(idx, 2, line); }
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
    { let idx = chunk.add_import("ecma:string", "trim"); chunk.emit_call(idx, 1, line); }
    { let idx = chunk.add_import("ecma:string", "toLowerCase"); chunk.emit_call(idx, 1, line); }
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
    { let idx = chunk.add_import("ecma:string", "replaceAll"); chunk.emit_call(idx, 3, line); }
    lset(chunk, data_slot, line);

    lget(chunk, data_slot, line);
    push_str(chunk, "'", line);
    push_str(chunk, "\\'", line);
    { let idx = chunk.add_import("ecma:string", "replaceAll"); chunk.emit_call(idx, 3, line); }
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
