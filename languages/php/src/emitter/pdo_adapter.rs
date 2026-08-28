//! PHP `PDO` / `PDOStatement` object surface. Split out of the former
//! `database_adapter.rs`; the mysqli surface lives in `mysqli_adapter.rs`.

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ValueSource,
};
use vybe_compiler::primitives::instructions::core_wasm;

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use crate::emitter::string_adapter;
use vybe_compiler::primitives::{collections, convert};

const PDO_FETCH_COLUMN: f64 = 7.0;
const PDO_FETCH_CLASS: f64 = 8.0;
const PDO_FETCH_INTO: f64 = 9.0;
const PDO_FETCH_FUNC: f64 = 10.0;
const PDO_FETCH_KEY_PAIR: f64 = 12.0;
const PDO_FETCH_GROUP_BIT: f64 = 65_536.0;

/// A PDO fetch mode is a base style in the low 16 bits OR-ed with modifier bits
/// (`FETCH_GROUP`, `FETCH_UNIQUE`, `FETCH_CLASSTYPE`, `FETCH_PROPS_LATE`, …).
/// Testing `mode > 65535` conflates them — `FETCH_CLASS|FETCH_PROPS_LATE` is
/// 1048584 and is not a grouped fetch — so isolate the one bit that was asked
/// for: `floor(mode / bit) - 2 * floor(mode / 2bit) == 1`. Leaves an **i32**.
fn emit_mode_has_bit(chunk: &mut Chunk, mode_slot: u16, bit: f64, line: u32) {
    lget(chunk, mode_slot, line);
    push_const(chunk, Value::F64(bit), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);

    lget(chunk, mode_slot, line);
    push_const(chunk, Value::F64(bit * 2.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::F64_MUL, line);

    chunk.emit_op(Op::F64_SUB, line);
    push_const(chunk, Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
}

/// The base fetch style — the mode with every modifier bit cleared, i.e. the
/// low 16 bits. Leaves an f64 on the stack.
fn emit_mode_base(chunk: &mut Chunk, mode_slot: u16, line: u32) {
    lget(chunk, mode_slot, line);
    lget(chunk, mode_slot, line);
    push_const(chunk, Value::F64(PDO_FETCH_GROUP_BIT), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    push_const(chunk, Value::F64(PDO_FETCH_GROUP_BIT), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
}

/// `base_mode == want` as an i32, for chaining the style arms.
fn emit_base_mode_is(chunk: &mut Chunk, mode_slot: u16, want: f64, line: u32) {
    emit_mode_base(chunk, mode_slot, line);
    push_const(chunk, Value::F64(want), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
}

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn push_const(chunk: &mut Chunk, value: Value, line: u32) {
    match &value {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
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

fn struct_get_key(chunk: &mut Chunk, key: &ClassSlot, line: u32) {
    let slot = class_slots::resolve(key, &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &slot, Dest::Stack, line);
}

fn struct_set_key(chunk: &mut Chunk, key: &ClassSlot, line: u32) {
    let slot = class_slots::resolve(key, &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &slot, ValueSource::Stack, line);
}

#[allow(dead_code)]
fn global_set_key(chunk: &mut Chunk, key: &str, line: u32) {
    vybe_compiler::primitives::globals::emit_write(chunk, key, line);
}

#[allow(dead_code)]
fn global_get_key(chunk: &mut Chunk, key: &str, line: u32) {
    vybe_compiler::primitives::globals::emit_read(chunk, key, line);
}
fn emit_string_slot_nonempty(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    push_str(chunk, "", line);
    vybe_compiler::primitives::ops::emit_dyn_ne(chunk, line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    replace_in_slot(chunk, normalized_slot, "pgsql:", "", line);
    replace_in_slot(chunk, normalized_slot, "dbname=", "db=", line);
    lget(chunk, normalized_slot, line);
    push_str(chunk, "port=", line);
    {
        let idx = chunk.add_import("ecma:string", "includes");
        chunk.emit_call(idx, 2, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    replace_in_slot(chunk, normalized_slot, "postgres:", "", line);
    replace_in_slot(chunk, normalized_slot, "dbname=", "db=", line);
    lget(chunk, normalized_slot, line);
    push_str(chunk, "port=", line);
    {
        let idx = chunk.add_import("ecma:string", "includes");
        chunk.emit_call(idx, 2, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
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
    let cs_id = class_slots::resolve(&ClassSlot::TypeIdentity, &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &cs_id, ValueSource::Stack, line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
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

/// Same as [`emit_column_value`] with the index taken from a slot, for
/// `fetchColumn($n)` where `$n` is only known at runtime.
fn emit_column_value_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    row_slot: u16,
    index_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    lget(chunk, row_slot, line);
    lget(chunk, index_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let direct_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, direct_slot, line);

    lget(chunk, direct_slot, line);
    {
        let undef_idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(undef_idx, 1, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    lget(chunk, row_slot, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let keys_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);

    lget(chunk, row_slot, line);
    lget(chunk, keys_slot, line);
    lget(chunk, index_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_else(line);
    lget(chunk, direct_slot, line);
    chunk.emit_end(line);
}

fn emit_empty_array(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_array_new(chunks, current, 0, line);
}

/// Read THROUGH a bound reference, if that is what the slot holds.
///
/// `bindParam`/`bind_param` take their variable by reference — php spells no
/// `&`, the binder declares it, and the php walker supplies it (see
/// `mark_php_bound_variable_args`). So what lands in `__bound_params` is a
/// reference CELL, `{__ref_kind, __value}`, not the value; reading the variable
/// happens here, at `execute()`, which is exactly when php reads it.
///
/// The test is structural, not a kind check: anything that is not null, a
/// number, a string or a boolean is asked for `__value`, and a `__value` that
/// comes back undefined means it was an ordinary object and the original stands.
/// A cell is indistinguishable from any other object at this layer and does not
/// need to be distinguished — a bound plain object has no meaning for a
/// `list<string>` parameter channel either way.
fn emit_resolve_bound_reference(
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
        struct_get_key(chunk, &ClassSlot::internal("__value"), line);
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
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);

        lget(chunk, inner_slot, line);
        lset(chunk, resolved_slot, line);

        chunk.emit_end(line);
        chunk.emit_end(line);
    }

    resolved_slot
}

fn emit_sql_literal_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    value_slot: u16,
    line: u32,
) -> u16 {
    let resolved_slot = emit_resolve_bound_reference(chunks, current, value_slot, line);

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

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, index_slot, line);
        lget(chunk, len_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    }
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

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
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
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
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, index_slot, line);
    lget(chunk, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
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
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
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
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
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
    let cs_id = class_slots::resolve(&ClassSlot::TypeIdentity, &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &cs_id, ValueSource::Stack, line);

    if let Some(slot) = sql_slot {
        lget(chunk, stmt_slot, line);
        lget(chunk, slot, line);
        struct_set_key(chunk, &ClassSlot::internal("commandtext"), line);

        lget(chunk, stmt_slot, line);
        lget(chunk, slot, line);
        struct_set_key(chunk, &ClassSlot::internal("__prepared_commandtext"), line);
    }

    lget(chunk, stmt_slot, line);
    lget(chunk, conn_slot, line);
    struct_set_key(chunk, &ClassSlot::internal("__conn"), line);

    lget(chunk, stmt_slot, line);
    if let Some(slot) = rows_slot {
        lget(chunk, slot, line);
    } else {
        emit_empty_array(chunks, current, line);
    }
    let chunk = &mut chunks[current];
    struct_set_key(chunk, &ClassSlot::internal("__rows"), line);

    lget(chunk, stmt_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, &ClassSlot::internal("__cursor"), line);

    lget(chunk, stmt_slot, line);
    emit_empty_array(chunks, current, line);
    let chunk = &mut chunks[current];
    struct_set_key(chunk, &ClassSlot::internal("__bound_params"), line);

    lget(chunk, stmt_slot, line);
    emit_empty_array(chunks, current, line);
    let chunk = &mut chunks[current];
    struct_set_key(chunk, &ClassSlot::internal("__bound_named_pairs"), line);

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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, lower_slot, line);
    push_str(chunk, ",", line);
    {
        let idx = chunk.add_import("ecma:string", "split");
        chunk.emit_call(idx, 2, line);
    }
    vybe_compiler::primitives::collections::emit_array_length(chunk, line);
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

/// After a `wasi:sql` call, act on whether it FAILED.
///
/// The host used to print the failure and hand back `-1` or an empty row set,
/// so `$db->exec("INSERT …")` on a read-only database looked like a success
/// that changed no rows. `wasi:sql.lastError(conn)` returns the trace (empty
/// when the last call succeeded).
///
/// The trace is stashed as `__pdo_error` for `errorCode()` to report, and under
/// `ERRMODE_EXCEPTION` (2) it is also RAISED — `ERRMODE_SILENT` (0) and
/// `ERRMODE_WARNING` (1) leave the caller to ask, exactly as PHP does.
///
/// Throwing goes through `errors::emit_throw`, which imports the exception TAG
/// and emits it as the opcode's two operand bytes. A bare `emit_op(Op::THROW)`
/// reads the following two bytes of adapter code as the tag and dies with
/// `catch label 0 out of range` — the failure looks structural and is not.
/// Stack: unchanged.
fn emit_record_failure(chunks: &mut [Chunk], current: usize, conn_slot: u16, line: u32) {
    let msg_slot = alloc_local(&mut chunks[current]);
    lget(&mut chunks[current], conn_slot, line);
    call_import(chunks, current, "wasi:sql", "lastError", 1, line);
    lset(&mut chunks[current], msg_slot, line);

    lget(&mut chunks[current], conn_slot, line);
    lget(&mut chunks[current], msg_slot, line);
    struct_set_key(&mut chunks[current], &ClassSlot::internal("__pdo_error"), line);

    emit_string_slot_nonempty(&mut chunks[current], msg_slot, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], conn_slot, line);
    struct_get_key(&mut chunks[current], &ClassSlot::internal("__pdo_attr"), line);
    push_const(&mut chunks[current], Value::F64(2.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    // PHP's message shape, so `$e->getMessage()` reads the way a PDO user
    // expects rather than exposing the bare driver text.
    let text_slot = alloc_local(&mut chunks[current]);
    push_str(&mut chunks[current], "SQLSTATE[HY000]: General error: ", line);
    lset(&mut chunks[current], text_slot, line);
    concat_slot_with_slot(&mut chunks[current], text_slot, msg_slot, line);
    class_slots::emit_class_alloc(&mut chunks[current], line);
    chunks[current].emit_dup(line);
    lget(&mut chunks[current], text_slot, line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        &mut chunks[current],
        "PDOException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
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
    let affected_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], affected_slot, line);
    emit_record_failure(chunks, current, conn_slot, line);
    lget(&mut chunks[current], affected_slot, line);
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
    struct_set_key(chunk, &ClassSlot::internal("__pdo_attr"), line);
    push_const(chunk, Value::Bool(true), line);
}

/// Run a `wasi:sql` transaction verb and leave `__in_tx` on the connection to
/// match — `inTransaction()` is a guest-side flag because the host surface has
/// no way to ask. Stack: `[conn]` → `[result]`.
fn emit_transaction_verb(chunks: &mut [Chunk], current: usize, verb: &str, in_tx: bool, line: u32) {
    let chunk = &mut chunks[current];
    let conn_slot = alloc_local(chunk);
    lset(chunk, conn_slot, line);
    lget(chunk, conn_slot, line);
    call_import(chunks, current, "wasi:sql", verb, 1, line);
    let chunk = &mut chunks[current];
    let result_slot = alloc_local(chunk);
    lset(chunk, result_slot, line);

    lget(chunk, conn_slot, line);
    push_const(chunk, Value::Bool(in_tx), line);
    struct_set_key(chunk, &ClassSlot::internal("__in_tx"), line);
    lget(chunk, result_slot, line);
}

pub fn emit_php_pdo_begin_transaction(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_transaction_verb(chunks, current, "beginTransaction", true, line);
}

pub fn emit_php_pdo_commit(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_transaction_verb(chunks, current, "commit", false, line);
}

pub fn emit_php_pdo_rollback(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_transaction_verb(chunks, current, "rollback", false, line);
}

/// `$pdo->inTransaction()` — false until `beginTransaction()`, false again
/// after `commit()`/`rollBack()`. Stack: `[conn]` → `[bool]`.
pub fn emit_php_pdo_in_transaction(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    struct_get_key(chunk, &ClassSlot::internal("__in_tx"), line);
    let flag_slot = alloc_local(chunk);
    lset(chunk, flag_slot, line);
    lget(chunk, flag_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_else(line);
    lget(chunk, flag_slot, line);
    chunk.emit_end(line);
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
    struct_get_key(chunk, &ClassSlot::internal("__bound_named_pairs"), line);
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
    struct_get_key(chunk, &ClassSlot::internal("__bound_params"), line);
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

/// Read through every bound reference in the POSITIONAL parameter list.
///
/// This has to run FIRST, before stringifying and before null-inlining: both of
/// those ask what a parameter IS — is it null, is it a bool — and a cell answers
/// for the wrapper, not for the variable. A cell holding null is an object, so
/// null-inlining would miss it; stringifying it would render the wrapper.
/// Resolving once here means neither pass, nor the mysqli binder that writes
/// into this same `__bound_params`, needs to know references exist.
///
/// Returns a NEW array. A keyed map is left alone — its values were already
/// substituted into the SQL text through `emit_sql_literal_from_slot`, which
/// resolves references on its own.
fn emit_resolve_bound_references(
    chunks: &mut [Chunk],
    current: usize,
    params_slot: u16,
    line: u32,
) -> u16 {
    let result_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lget(chunk, params_slot, line);
    lset(chunk, result_slot, line);

    lget(chunk, params_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    let _ = chunk;
    emit_empty_array(chunks, current, line);
    let out_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, out_slot, line);

    push_const(chunk, Value::F64(0.0), line);
    let i_slot = alloc_local(chunk);
    lset(chunk, i_slot, line);

    lget(chunk, params_slot, line);
    let _ = chunk;
    collections::emit_len(chunks, current, line);
    let n_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, n_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, params_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let v_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, v_slot, line);

    let _ = chunk;
    let resolved_slot = emit_resolve_bound_reference(chunks, current, v_slot, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    lget(chunk, resolved_slot, line);
    let _ = chunk;
    collections::emit_push(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    lset(chunk, result_slot, line);
    chunk.emit_end(line);

    result_slot
}

/// `wasi:sql`'s parameter channel is `Vec<String>` — the host flattens every
/// bound value with `format!("{}", v)`. That is vybe's generic rendering, not
/// php's cast, so a bound `true`/`false` reached sqlite as the TEXT
/// `'true'`/`'false'` instead of php's `'1'`/`''`. The coercion is php's to
/// make, so make it here, through the SAME emitter the
/// `[builtin_slots.string] to_string` binding uses — `emit_echo_stringify` is
/// php's one value → string rule and already knows `true` is `"1"`, an array
/// is `"Array"`, and a bigint keeps its digits. Returns a NEW array; the
/// statement's own `__bound_params` must not be rewritten.
///
/// NULL is deliberately left alone. php binds it as SQL NULL, and a
/// `Vec<String>` cannot say that; stringifying it here would store `''` and
/// merely hide the host limitation behind a second wrong answer.
fn emit_php_stringify_params(
    chunks: &mut [Chunk],
    current: usize,
    params_slot: u16,
    line: u32,
) -> u16 {
    // Only a POSITIONAL list is rewritten. `execute([':m' => 'boot'])` hands
    // over a KEYED map whose values are already substituted into the SQL text,
    // and walking that with `len`/`ARRAY_GET` replaced it with an EMPTY array —
    // which silently inserted nothing at all.
    let result_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lget(chunk, params_slot, line);
    lset(chunk, result_slot, line);

    lget(chunk, params_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    let _ = chunk;
    emit_empty_array(chunks, current, line);
    let out_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, out_slot, line);

    push_const(chunk, Value::F64(0.0), line);
    let i_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, i_slot, line);

    lget(chunk, params_slot, line);
    let _ = chunk;
    collections::emit_len(chunks, current, line);
    let n_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, n_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, params_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let v_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, v_slot, line);

    lget(chunk, out_slot, line);
    lget(chunk, v_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    lget(chunk, v_slot, line);
    chunk.emit_else(line);
    lget(chunk, v_slot, line);
    let _ = chunk;
    string_adapter::emit_echo_stringify(chunks, current, 1, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    let _ = chunk;
    collections::emit_push(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    lset(chunk, result_slot, line);
    chunk.emit_end(line);

    result_slot
}

/// Inline null-valued positional parameters into the SQL as the keyword NULL.
///
/// `wasi:sql`'s parameter channel is `list<string>` **in the spec**
/// (`proposals/wasi-sql/wit/types.wit`: `prepare(query, params: list<string>)`),
/// so there is no spec-conformant way to BIND a NULL — the `null` arm of
/// `data-type` describes a value coming back in a `row`, not one going in.
/// Widening the host's param type would deviate from the spec, so the null is
/// resolved on this side instead, exactly as the NAMED path already does via
/// `emit_sql_literal_from_slot`.
///
/// Rewrites `sql_slot` in place and returns a slot holding the surviving
/// params. Scanning is QUOTE-AWARE: a `?` inside a string literal is data, and
/// the drivers parse it correctly today, so it must not be substituted. A
/// doubled `''` escape falls out of toggling on every quote. When no param is
/// null nothing is rewritten and the driver keeps binding as before.
fn emit_inline_null_positional_params(
    chunks: &mut [Chunk],
    current: usize,
    sql_slot: u16,
    params_slot: u16,
    line: u32,
) -> u16 {
    let result_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lget(chunk, params_slot, line);
    lset(chunk, result_slot, line);

    // Only positional lists; a keyed map is already substituted into the text.
    lget(chunk, params_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    // out = "" ; kept = [] ; i = 0 ; k = 0 ; in_quote = 0
    push_str(chunk, "", line);
    let out_slot = alloc_local(chunk);
    lset(chunk, out_slot, line);
    let _ = chunk;
    emit_empty_array(chunks, current, line);
    let kept_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, kept_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    let i_slot = alloc_local(chunk);
    lset(chunk, i_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    let k_slot = alloc_local(chunk);
    lset(chunk, k_slot, line);
    core_wasm::i32_const(chunk, line, 0);
    let quote_slot = alloc_local(chunk);
    lset(chunk, quote_slot, line);
    core_wasm::i32_const(chunk, line, 0);
    let hit_slot = alloc_local(chunk);
    lset(chunk, hit_slot, line);

    lget(chunk, sql_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    let n_slot = alloc_local(chunk);
    lset(chunk, n_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, sql_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    let c_slot = alloc_local(chunk);
    lset(chunk, c_slot, line);

    // c == "'" → flip the quote state
    lget(chunk, c_slot, line);
    push_str(chunk, "'", line);
    {
        let idx = chunk.add_import("wasm:js-string", "equals");
        chunk.emit_call(idx, 2, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, quote_slot, line);
    chunk.emit_op(Op::I32_EQZ, line);
    lset(chunk, quote_slot, line);
    chunk.emit_end(line);

    // placeholder = (c == "?") && !in_quote
    lget(chunk, c_slot, line);
    push_str(chunk, "?", line);
    {
        let idx = chunk.add_import("wasm:js-string", "equals");
        chunk.emit_call(idx, 2, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    lget(chunk, quote_slot, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);

    lget(chunk, params_slot, line);
    lget(chunk, k_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let pv_slot = alloc_local(chunk);
    lset(chunk, pv_slot, line);

    lget(chunk, pv_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    // null → the SQL keyword, and the param is consumed
    concat_slot_with_literal(chunk, out_slot, "NULL", line);
    core_wasm::i32_const(chunk, line, 1);
    lset(chunk, hit_slot, line);
    chunk.emit_else(line);
    concat_slot_with_literal(chunk, out_slot, "?", line);
    lget(chunk, kept_slot, line);
    lget(chunk, pv_slot, line);
    let _ = chunk;
    collections::emit_push(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);

    lget(chunk, k_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, k_slot, line);

    chunk.emit_else(line);
    concat_slot_with_slot(chunk, out_slot, c_slot, line);
    chunk.emit_end(line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    // Leave everything untouched unless a null was actually inlined.
    lget(chunk, hit_slot, line);
    chunk.emit_if(line);
    lget(chunk, out_slot, line);
    lset(chunk, sql_slot, line);
    lget(chunk, kept_slot, line);
    lset(chunk, result_slot, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    result_slot
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
        struct_get_key(chunk, &ClassSlot::internal("__prepared_commandtext"), line);
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
        struct_set_key(chunk, &ClassSlot::internal("field_count"), line);
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
            struct_get_key(chunk, &ClassSlot::internal("__bound_params"), line);
            lset(chunk, effective_params_slot, line);

            lget(chunk, stmt_slot, line);
            struct_get_key(chunk, &ClassSlot::internal("__bound_named_pairs"), line);
        }
        let named_pairs_slot = alloc_local(&mut chunks[current]);
        {
            let chunk = &mut chunks[current];
            lset(chunk, named_pairs_slot, line);
        }
        emit_apply_named_bound_pairs(chunks, current, sql_text_slot, named_pairs_slot, line);
    }

    // A bound variable arrives as a reference; php reads it HERE, at execute.
    // Before anything asks what a parameter is, make it be the value.
    let deref_params_slot =
        emit_resolve_bound_references(chunks, current, effective_params_slot, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, deref_params_slot, line);
        lset(chunk, effective_params_slot, line);
    }

    // Hand the host php's own string for each param, not vybe's generic one.
    let normalized_params_slot =
        emit_php_stringify_params(chunks, current, effective_params_slot, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, normalized_params_slot, line);
        lset(chunk, effective_params_slot, line);
    }
    // Nulls survive the stringify pass untouched so they are still visible here.
    let kept_params_slot = emit_inline_null_positional_params(
        chunks,
        current,
        sql_text_slot,
        effective_params_slot,
        line,
    );
    {
        let chunk = &mut chunks[current];
        lget(chunk, kept_params_slot, line);
        lset(chunk, effective_params_slot, line);
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
    struct_set_key(chunk, &ClassSlot::internal("commandtext"), line);
    lget(chunk, stmt_slot, line);
    lget(chunk, sql_text_slot, line);
    lget(chunk, effective_params_slot, line);
    call_import(chunks, current, "wasi:sql", "query", 3, line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);
    lget(chunk, stmt_slot, line);
    lget(chunk, rows_slot, line);
    struct_set_key(chunk, &ClassSlot::internal("__rows"), line);
    lget(chunk, stmt_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, &ClassSlot::internal("__cursor"), line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    lget(chunk, stmt_slot, line);
    lget(chunk, sql_text_slot, line);
    struct_set_key(chunk, &ClassSlot::internal("commandtext"), line);
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
    struct_set_key(chunk, &ClassSlot::internal("__rows"), line);
    lget(chunk, stmt_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, &ClassSlot::internal("__cursor"), line);
    lget(chunk, stmt_slot, line);
    lget(chunk, count_slot, line);
    struct_set_key(chunk, &ClassSlot::internal("__row_count"), line);
    lget(chunk, count_slot, line);
    push_const(chunk, Value::F64(-1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ne(chunk, line);

    chunk.emit_end(line);
}

/// `$stmt->setFetchMode($mode, $arg)` — remember the mode (and `FETCH_INTO`'s
/// target object or `FETCH_CLASS`'s class name) so a later argument-less
/// `fetch()` honours it. Stack: `[stmt, mode, arg?]` → `[true]`.
pub fn emit_php_pdo_statement_set_fetch_mode(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    for _ in 3..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let arg_slot = alloc_local(chunk);
    if argc >= 3 {
        lset(chunk, arg_slot, line);
    } else {
        push_const(chunk, Value::Null, line);
        lset(chunk, arg_slot, line);
    }
    let mode_slot = alloc_local(chunk);
    lset(chunk, mode_slot, line);
    let stmt_slot = alloc_local(chunk);
    lset(chunk, stmt_slot, line);

    lget(chunk, stmt_slot, line);
    lget(chunk, mode_slot, line);
    struct_set_key(chunk, &ClassSlot::internal("__fetch_mode"), line);
    lget(chunk, stmt_slot, line);
    lget(chunk, arg_slot, line);
    struct_set_key(chunk, &ClassSlot::internal("__fetch_arg"), line);
    chunk.emit_bool_const(true, line);
}

/// Copy a row's NAMED columns onto an existing object — `PDO::FETCH_INTO`.
/// The row carries each column twice, once positionally and once by name, so
/// the positional keys are skipped via php's own key rule (a canonical decimal
/// integer string normalises to a number; a column name does not).
fn emit_copy_named_columns_into(
    chunks: &mut [Chunk],
    current: usize,
    row_slot: u16,
    target_slot: u16,
    line: u32,
) {
    lget(&mut chunks[current], row_slot, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let keys_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);

    push_const(chunk, Value::F64(0.0), line);
    let i_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, i_slot, line);

    lget(chunk, keys_slot, line);
    let _ = chunk;
    collections::emit_len(chunks, current, line);
    let n_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, n_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, keys_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let key_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, key_slot, line);

    // A key that normalises to a NUMBER is the positional duplicate.
    let _ = chunk;
    crate::emitter::array_adapter::emit_php_array_key(chunks, current, key_slot, line);
    let chunk = &mut chunks[current];
    {
        let num_idx = chunk.add_import("wasm:js-number", "test");
        chunk.emit_call(num_idx, 1, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);

    lget(chunk, target_slot, line);
    lget(chunk, key_slot, line);
    lget(chunk, row_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "set", 3, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
}

pub fn emit_php_pdo_statement_fetch(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 2..argc {
        chunk.emit_op(Op::DROP, line);
    }
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
    struct_get_key(chunk, &ClassSlot::internal("__rows"), line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);

    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, &ClassSlot::internal("__cursor"), line);
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
    struct_set_key(chunk, &ClassSlot::internal("__cursor"), line);

    // The effective mode is the explicit argument, else whatever a previous
    // `setFetchMode()` stored on the statement. A missing one is 0, which
    // matches no style and falls through to the plain row.
    let eff_slot = alloc_local(chunk);
    match mode_slot {
        Some(slot) => {
            lget(chunk, slot, line);
            lset(chunk, eff_slot, line);
        }
        None => {
            lget(chunk, stmt_slot, line);
            struct_get_key(chunk, &ClassSlot::internal("__fetch_mode"), line);
            let chunk = &mut chunks[current];
            lset(chunk, eff_slot, line);
        }
    }
    let chunk = &mut chunks[current];
    lget(chunk, eff_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, eff_slot, line);
    chunk.emit_end(line);

    // Past the last row `fetch()` is `false` in every mode.
    lget(chunk, row_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_else(line);

    emit_base_mode_is(chunk, eff_slot, PDO_FETCH_COLUMN, line);
    chunk.emit_if_value(line);
    let _ = chunk;
    emit_first_column_value(chunks, current, row_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);

    // FETCH_INTO — fill the object handed to setFetchMode() and hand it back.
    emit_base_mode_is(chunk, eff_slot, PDO_FETCH_INTO, line);
    chunk.emit_if_value(line);
    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, &ClassSlot::internal("__fetch_arg"), line);
    let target_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, target_slot, line);
    let _ = chunk;
    emit_copy_named_columns_into(chunks, current, row_slot, target_slot, line);
    let chunk = &mut chunks[current];
    lget(chunk, target_slot, line);
    chunk.emit_else(line);

    // FETCH_CLASS — stamp the row with the requested class, the same shape
    // `fetchObject()` already produces. FETCH_PROPS_LATE only orders the
    // constructor against the property writes, and no constructor runs here.
    emit_base_mode_is(chunk, eff_slot, PDO_FETCH_CLASS, line);
    chunk.emit_if_value(line);
    lget(chunk, row_slot, line);
    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, &ClassSlot::internal("__fetch_arg"), line);
    let chunk = &mut chunks[current];
    let cs_id = class_slots::resolve(&ClassSlot::TypeIdentity, &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &cs_id, ValueSource::Stack, line);
    lget(chunk, row_slot, line);
    chunk.emit_else(line);

    lget(chunk, row_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_php_pdo_statement_fetch_all(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    // `fetchAll($mode, $arg)` — FETCH_FUNC's callback, FETCH_CLASS's class name
    // and FETCH_COLUMN's index all arrive as a THIRD stack entry. Popping only
    // two landed the ARGUMENT in the mode slot and the MODE in the statement
    // slot, so every read afterwards was against a number.
    for _ in 3..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let arg_slot = if argc >= 3 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let mode_slot = if argc >= 2 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let stmt_slot = alloc_local(chunk);
    if let Some(slot) = arg_slot {
        lset(chunk, slot, line);
    }
    if let Some(slot) = mode_slot {
        lset(chunk, slot, line);
    }
    lset(chunk, stmt_slot, line);

    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, &ClassSlot::internal("__rows"), line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);

    if let Some(slot) = mode_slot {
        emit_mode_has_bit(chunk, slot, PDO_FETCH_GROUP_BIT, line);
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
        let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
        let chunk = &mut chunks[current];
        lget(chunk, index_slot, line);
        lget(chunk, len_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
        let _ = chunk;
        vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
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
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
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
        // `FETCH_GROUP|FETCH_COLUMN` collects the SECOND column keyed by the
        // first one's value; a plain grouped fetch collects the whole row.
        emit_base_mode_is(chunk, slot, PDO_FETCH_COLUMN, line);
        chunk.emit_if_value(line);
        let _ = chunk;
        emit_column_value(chunks, current, row_slot, 1.0, line);
        let chunk = &mut chunks[current];
        chunk.emit_else(line);
        lget(chunk, row_slot, line);
        chunk.emit_end(line);
        let _ = chunk;
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
        vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
        let chunk = &mut chunks[current];
        lget(chunk, out_slot, line);
        chunk.emit_else(line);

        lget(chunk, slot, line);
        push_const(chunk, Value::F64(PDO_FETCH_KEY_PAIR), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
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
        let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
        let chunk = &mut chunks[current];
        lget(chunk, index_slot, line);
        lget(chunk, len_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
        let _ = chunk;
        vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
        let chunk = &mut chunks[current];

        lget(chunk, rows_slot, line);
        lget(chunk, index_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        let row_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, row_slot, line);

        lget(chunk, row_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
        chunk.emit_if(line);

        lget(chunk, row_slot, line);
        push_const(chunk, Value::F64(2.0), line);
        chunk.emit_op(Op::ARRAY_GET, line);
        {
            let undef_idx = chunk.add_import("wasm:js-undefined", "test");
            chunk.emit_call(undef_idx, 1, line);
        }
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
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
        vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
        let chunk = &mut chunks[current];
        lget(chunk, out_slot, line);
        chunk.emit_else(line);

        lget(chunk, slot, line);
        push_const(chunk, Value::F64(PDO_FETCH_COLUMN), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
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
        let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
        let chunk = &mut chunks[current];
        lget(chunk, index_slot, line);
        lget(chunk, len_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
        let _ = chunk;
        vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
        let chunk = &mut chunks[current];

        lget(chunk, rows_slot, line);
        lget(chunk, index_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        let row_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, row_slot, line);

        lget(chunk, row_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
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
        vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
        let chunk = &mut chunks[current];
        lget(chunk, out_slot, line);
        chunk.emit_else(line);

        // ── FETCH_FUNC ──────────────────────────────────────────────────────
        // The callback takes one argument per column, so the arity is a
        // RUNTIME property of the row and fixed-arity callable invoke cannot express it. Collect
        // the row's positional columns into an array and go through
        // `Reflect.apply`, which spreads them for any width.
        //
        // Only reachable when a callback was actually passed — `fetchAll($mode)`
        // with two args has no `arg_slot`, and this arm is emitted for every
        // mode, so it must not be built at all in that case.
        let Some(cb_slot) = arg_slot else {
            lget(chunk, rows_slot, line);
            chunk.emit_end(line);
            chunk.emit_end(line);
            chunk.emit_end(line);
            return;
        };
        emit_base_mode_is(chunk, slot, PDO_FETCH_FUNC, line);
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
        let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
        let chunk = &mut chunks[current];
        lget(chunk, index_slot, line);
        lget(chunk, len_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
        let _ = chunk;
        vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
        let chunk = &mut chunks[current];

        lget(chunk, rows_slot, line);
        lget(chunk, index_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        let row_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, row_slot, line);

        // args = [] ; while row[j] is defined: args.push(row[j])
        let _ = chunk;
        emit_empty_array(chunks, current, line);
        let args_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, args_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        let col_slot = alloc_local(&mut chunks[current]);
        let chunk = &mut chunks[current];
        lset(chunk, col_slot, line);

        let _ = chunk;
        let col_loop = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
        let chunk = &mut chunks[current];
        lget(chunk, row_slot, line);
        lget(chunk, col_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        {
            let undef_idx = chunk.add_import("wasm:js-undefined", "test");
            chunk.emit_call(undef_idx, 1, line);
        }
        // A NULL column is a legitimate value; only `undefined` ends the row.
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        let _ = chunk;
        vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
        let chunk = &mut chunks[current];

        lget(chunk, args_slot, line);
        lget(chunk, row_slot, line);
        lget(chunk, col_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        let _ = chunk;
        collections::emit_push(chunks, current, line);
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);

        lget(chunk, col_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, col_slot, line);
        let _ = chunk;
        vybe_compiler::primitives::loops::emit_loop_end(chunks, current, col_loop, line);
        let chunk = &mut chunks[current];

        lget(chunk, out_slot, line);
        lget(chunk, cb_slot, line);
        push_const(chunk, Value::Null, line);
        lget(chunk, args_slot, line);
        let _ = chunk;
        call_import(chunks, current, "ecma:reflect", "apply", 3, line);
        collections::emit_push(chunks, current, line);
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);

        lget(chunk, index_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, index_slot, line);
        let _ = chunk;
        vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
        let chunk = &mut chunks[current];
        lget(chunk, out_slot, line);

        chunk.emit_else(line);
        lget(chunk, rows_slot, line);
        chunk.emit_end(line);
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
        let cs_id = class_slots::resolve(&ClassSlot::TypeIdentity, &PlainNames);
        class_slots::emit_class_set(chunk, ObjSource::Stack, &cs_id, ValueSource::Stack, line);
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
    argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    // `$argc` was ignored, so `fetchColumn(0)` popped the COLUMN INDEX into the
    // statement slot and every read after it was against a number.
    let index_slot = alloc_local(chunk);
    if argc >= 2 {
        lset(chunk, index_slot, line);
    } else {
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, index_slot, line);
    }
    let stmt_slot = alloc_local(chunk);
    lset(chunk, stmt_slot, line);
    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, &ClassSlot::internal("__rows"), line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);
    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, &ClassSlot::internal("__cursor"), line);
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
    struct_set_key(chunk, &ClassSlot::internal("__cursor"), line);
    // row === null ? false : firstColumn(row)
    lget(chunk, row_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_else(line);
    let _ = chunk;
    emit_column_value_from_slot(chunks, current, row_slot, index_slot, line);
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
    struct_get_key(chunk, &ClassSlot::internal("__row_count"), line);
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
    struct_get_key(chunk, &ClassSlot::internal("__rows"), line);
    let rows_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, rows_slot, line);
    // rows.length > 0 ? Object.keys(rows[0]).length : stmt.field_count
    lget(chunk, rows_slot, line);
    vybe_compiler::primitives::collections::emit_array_length(chunk, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, rows_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::collections::emit_array_length(chunk, line);
    chunk.emit_else(line);
    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, &ClassSlot::internal("field_count"), line);
    {
        let undef_idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(undef_idx, 1, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_else(line);
    lget(chunk, stmt_slot, line);
    struct_get_key(chunk, &ClassSlot::internal("field_count"), line);
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
    struct_get_key(chunk, &ClassSlot::internal("param_count"), line);
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
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    push_str(chunk, "'", line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
}

/// `$stmt->errorCode()` → SQLSTATE "00000" on success. Stack: `[stmt]` → `[str]`.
/// `$pdo->errorCode()` — `00000` when the last call succeeded, `HY000` when it
/// did not. Previously hard-coded to success, which made a failed write
/// indistinguishable from a successful one.
pub fn emit_php_pdo_error_code(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc as u16 - 1 {
        chunk.emit_op(Op::DROP, line);
    }
    let conn_slot = alloc_local(chunk);
    lset(chunk, conn_slot, line);
    let msg_slot = alloc_local(&mut chunks[current]);
    lget(&mut chunks[current], conn_slot, line);
    struct_get_key(&mut chunks[current], &ClassSlot::internal("__pdo_error"), line);
    lset(&mut chunks[current], msg_slot, line);

    let code_slot = alloc_local(&mut chunks[current]);
    push_str(&mut chunks[current], "00000", line);
    lset(&mut chunks[current], code_slot, line);
    emit_string_slot_nonempty(&mut chunks[current], msg_slot, line);
    chunks[current].emit_if(line);
    push_str(&mut chunks[current], "HY000", line);
    lset(&mut chunks[current], code_slot, line);
    chunks[current].emit_end(line);
    lget(&mut chunks[current], code_slot, line);
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
    struct_get_key(chunk, &ClassSlot::internal("__rows"), line);
    let rows_slot = alloc_local(chunk);
    lset(chunk, rows_slot, line);
    lget(chunk, rows_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let row_slot = alloc_local(chunk);
    lset(chunk, row_slot, line);
    emit_first_column_value(chunks, current, row_slot, line);
    // PDO::lastInsertId() is documented to return a STRING (the drivers hand
    // back text), so `is_string($pdo->lastInsertId())` is true even for an
    // integer primary key.
    vybe_compiler::primitives::convert::emit_to_string(&mut chunks[current], line);
}

/// `$stmt->errorInfo()` / `$pdo->errorInfo()` → `["00000", null, null]`.
/// Stack: `[recv]` → `[array]`.
pub fn emit_php_pdo_error_info(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc as u16 {
        chunk.emit_op(Op::DROP, line);
    }
    push_str(chunk, "00000", line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_array_new_fixed(0, 3, line);
}
