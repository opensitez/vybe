//! Lua operator adapter support.
//!
//! The walker normalizes Lua operator syntax to `__lua_*` calls so the
//! language surface can grow metatable dispatch later without changing the AST
//! shape. The baseline path must still be pure Vybe/ECMA/WASM, not a bespoke
//! `ecma:lua` host surface, so these adapters delegate to common emitter
//! primitives.

use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

fn call1(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 1, line);
}

fn call2(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 2, line);
}

fn call3(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 3, line);
}

fn call4(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 4, line);
}

fn i32_const(chunk: &mut Chunk, value: i32, line: u32) {
    chunk.emit_i32_const(value, line);
}

fn save(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn load(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn emit_is_undefined(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-undefined", "test");
    call1(chunk, idx, line);
}

fn emit_object_get_const_key(chunk: &mut Chunk, obj_slot: u16, key: &str, line: u32) {
    load(chunk, obj_slot, line);
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key_idx, line);
}

fn emit_lua_table_get_const_key(
    chunks: &mut Vec<Chunk>,
    current: usize,
    obj_slot: u16,
    key: &str,
    line: u32,
) {
    load(&mut chunks[current], obj_slot, line);
    chunks[current].emit_string_const(key, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
}

fn emit_lua_missing_to_nil(chunk: &mut Chunk, value_slot: u16, line: u32) {
    load(chunk, value_slot, line);
    emit_is_undefined(chunk, line);
    chunk.emit_if(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_else(line);
    load(chunk, value_slot, line);
    chunk.emit_end(line);
}

fn emit_lua_table_bounds_error(chunk: &mut Chunk, line: u32) {
    chunk.emit_string_const(
        "bad argument #2 to table operation (position out of bounds)",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
}

fn emit_lua_invalid_table_key_error(chunk: &mut Chunk, line: u32) {
    chunk.emit_string_const("table index is nil or NaN", line);
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
}

fn emit_lua_invalid_table_key_guard(chunk: &mut Chunk, key_slot: u16, line: u32) {
    let is_nan = chunk.add_import("ecma:number", "isNaN");
    emit_is_missing_value(chunk, key_slot, line);
    load(chunk, key_slot, line);
    call1(chunk, is_nan, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    emit_lua_invalid_table_key_error(chunk, line);
    chunk.emit_end(line);
}

fn emit_lua_key_needs_identity(chunk: &mut Chunk, key_slot: u16, line: u32) {
    let type_of = chunk.add_import("ecma:value", "typeof");
    let str_compare = chunk.add_import("wasm:js-string", "compare");
    load(chunk, key_slot, line);
    call1(chunk, type_of, line);
    chunk.emit_string_const("object", line);
    call2(chunk, str_compare, line);
    i32_const(chunk, 0, line);
    chunk.emit_op(Op::I32_EQ, line);
    load(chunk, key_slot, line);
    call1(chunk, type_of, line);
    chunk.emit_string_const("function", line);
    call2(chunk, str_compare, line);
    i32_const(chunk, 0, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn emit_lua_assoc_map(
    chunks: &mut Vec<Chunk>,
    current: usize,
    table_slot: u16,
    create: bool,
    line: u32,
) {
    let assoc_slot = chunks[current].alloc_scratch(1);
    emit_object_get_const_key(&mut chunks[current], table_slot, "__lua_assoc", line);
    save(&mut chunks[current], assoc_slot, line);
    load(&mut chunks[current], assoc_slot, line);
    emit_is_undefined(&mut chunks[current], line);
    chunks[current].emit_if(line);
    if create {
        let map_new = chunks[current].add_import("ecma:map", "new");
        chunks[current].emit_call(map_new, 0, line);
        save(&mut chunks[current], assoc_slot, line);
        load(&mut chunks[current], table_slot, line);
        load(&mut chunks[current], assoc_slot, line);
        let key_idx = chunks[current].add_constant(Value::String(Arc::from("__lua_assoc")));
        chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, key_idx, line);
        load(&mut chunks[current], assoc_slot, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    chunks[current].emit_else(line);
    load(&mut chunks[current], assoc_slot, line);
    chunks[current].emit_end(line);
}

fn emit_lua_table_read(
    chunks: &mut Vec<Chunk>,
    current: usize,
    table_slot: u16,
    key_slot: u16,
    line: u32,
) {
    let adjusted_key_slot = chunks[current].alloc_scratch(1);
    let assoc_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    let arr_test = chunks[current].add_import("ecma:array", "isArray");
    let num_test = chunks[current].add_import("wasm:js-number", "test");
    let map_get = chunks[current].add_import("ecma:map", "get");

    load(&mut chunks[current], table_slot, line);
    call1(&mut chunks[current], arr_test, line);
    load(&mut chunks[current], key_slot, line);
    call1(&mut chunks[current], num_test, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], key_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    save(&mut chunks[current], adjusted_key_slot, line);
    load(&mut chunks[current], table_slot, line);
    load(&mut chunks[current], adjusted_key_slot, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    save(&mut chunks[current], value_slot, line);
    chunks[current].emit_else(line);

    load(&mut chunks[current], table_slot, line);
    call1(&mut chunks[current], arr_test, line);
    emit_lua_key_needs_identity(&mut chunks[current], key_slot, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    emit_lua_assoc_map(chunks, current, table_slot, false, line);
    save(&mut chunks[current], assoc_slot, line);
    load(&mut chunks[current], assoc_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], assoc_slot, line);
    load(&mut chunks[current], key_slot, line);
    call2(&mut chunks[current], map_get, line);
    chunks[current].emit_end(line);
    save(&mut chunks[current], value_slot, line);
    chunks[current].emit_else(line);

    load(&mut chunks[current], table_slot, line);
    load(&mut chunks[current], key_slot, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    save(&mut chunks[current], value_slot, line);
    emit_is_missing_value(&mut chunks[current], value_slot, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], table_slot, line);
    load(&mut chunks[current], key_slot, line);
    call2(&mut chunks[current], map_get, line);
    save(&mut chunks[current], value_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    emit_lua_missing_to_nil(&mut chunks[current], value_slot, line);
}

fn emit_lua_table_write(
    chunks: &mut Vec<Chunk>,
    current: usize,
    table_slot: u16,
    key_slot: u16,
    value_slot: u16,
    return_table: bool,
    line: u32,
) {
    let adjusted_key_slot = chunks[current].alloc_scratch(1);
    let assoc_slot = chunks[current].alloc_scratch(1);
    let arr_test = chunks[current].add_import("ecma:array", "isArray");
    let num_test = chunks[current].add_import("wasm:js-number", "test");
    let map_set = chunks[current].add_import("ecma:map", "set");
    let map_delete = chunks[current].add_import("ecma:map", "delete");
    let object_delete = chunks[current].add_import("ecma:object", "delete");

    emit_lua_invalid_table_key_guard(&mut chunks[current], key_slot, line);
    emit_is_missing_value(&mut chunks[current], value_slot, line);

    load(&mut chunks[current], table_slot, line);
    call1(&mut chunks[current], arr_test, line);
    load(&mut chunks[current], key_slot, line);
    call1(&mut chunks[current], num_test, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], key_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    save(&mut chunks[current], adjusted_key_slot, line);
    load(&mut chunks[current], table_slot, line);
    load(&mut chunks[current], adjusted_key_slot, line);
    load(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);

    emit_lua_key_needs_identity(&mut chunks[current], key_slot, line);
    chunks[current].emit_if(line);

    load(&mut chunks[current], table_slot, line);
    call1(&mut chunks[current], arr_test, line);
    chunks[current].emit_if(line);
    emit_lua_assoc_map(chunks, current, table_slot, true, line);
    save(&mut chunks[current], assoc_slot, line);
    emit_is_missing_value(&mut chunks[current], value_slot, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], assoc_slot, line);
    load(&mut chunks[current], key_slot, line);
    call2(&mut chunks[current], map_delete, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], assoc_slot, line);
    load(&mut chunks[current], key_slot, line);
    load(&mut chunks[current], value_slot, line);
    call3(&mut chunks[current], map_set, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    emit_is_missing_value(&mut chunks[current], value_slot, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], table_slot, line);
    load(&mut chunks[current], key_slot, line);
    call2(&mut chunks[current], map_delete, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], table_slot, line);
    load(&mut chunks[current], key_slot, line);
    load(&mut chunks[current], value_slot, line);
    call3(&mut chunks[current], map_set, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], table_slot, line);
    call1(&mut chunks[current], arr_test, line);
    chunks[current].emit_if(line);
    emit_is_missing_value(&mut chunks[current], value_slot, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], table_slot, line);
    load(&mut chunks[current], key_slot, line);
    call2(&mut chunks[current], object_delete, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], table_slot, line);
    load(&mut chunks[current], key_slot, line);
    load(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    emit_is_missing_value(&mut chunks[current], value_slot, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], table_slot, line);
    load(&mut chunks[current], key_slot, line);
    call2(&mut chunks[current], map_delete, line);
    chunks[current].emit_op(Op::DROP, line);
    load(&mut chunks[current], table_slot, line);
    load(&mut chunks[current], key_slot, line);
    call2(&mut chunks[current], object_delete, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], table_slot, line);
    load(&mut chunks[current], key_slot, line);
    load(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    if return_table {
        load(&mut chunks[current], table_slot, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
}

pub fn emit_lua_table_from_pairs(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        let map_new = chunks[current].add_import("ecma:map", "new");
        chunks[current].emit_call(map_new, 0, line);
        return;
    }

    let rows = chunks[current].alloc_scratch(1);
    let table = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let row = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let map_new = chunks[current].add_import("ecma:map", "new");

    save(&mut chunks[current], rows, line);
    chunks[current].emit_call(map_new, 0, line);
    save(&mut chunks[current], table, line);
    i32_const(&mut chunks[current], 0, line);
    save(&mut chunks[current], i, line);
    load(&mut chunks[current], rows, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    save(&mut chunks[current], len, line);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    load(&mut chunks[current], i, line);
    load(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    load(&mut chunks[current], rows, line);
    load(&mut chunks[current], i, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    save(&mut chunks[current], row, line);

    load(&mut chunks[current], row, line);
    i32_const(&mut chunks[current], 0, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    save(&mut chunks[current], key, line);

    load(&mut chunks[current], row, line);
    i32_const(&mut chunks[current], 1, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    save(&mut chunks[current], value, line);

    emit_lua_table_write(chunks, current, table, key, value, false, line);
    chunks[current].emit_op(Op::DROP, line);

    load(&mut chunks[current], i, line);
    i32_const(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    save(&mut chunks[current], i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    load(&mut chunks[current], table, line);
}

fn emit_lua_tagged_handle(chunk: &mut Chunk, lua_type: &str, name: &str, line: u32) {
    let object = chunk.alloc_scratch(1);
    let object_new = chunk.add_import("ecma:object", "new");
    let type_key = chunk.add_constant(Value::String(Arc::from("__lua_type")));
    let name_key = chunk.add_constant(Value::String(Arc::from("__lua_name")));

    chunk.emit_call(object_new, 0, line);
    save(chunk, object, line);
    load(chunk, object, line);
    chunk.emit_string_const(lua_type, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, type_key, line);
    load(chunk, object, line);
    chunk.emit_string_const(name, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, name_key, line);
    load(chunk, object, line);
}

pub fn emit_lua_stdout(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    emit_lua_tagged_handle(&mut chunks[current], "userdata", "stdout", line);
}

pub fn emit_lua_coroutine_create(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        chunks[current].emit_string_const(
            "bad argument #1 to coroutine.create (function expected)",
            line,
        );
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        return;
    }

    let func = chunks[current].alloc_scratch(1);
    let co = chunks[current].alloc_scratch(1);
    let object_new = chunks[current].add_import("ecma:object", "new");
    let type_key = chunks[current].add_constant(Value::String(Arc::from("__lua_type")));
    let state_key = chunks[current].add_constant(Value::String(Arc::from("__lua_state")));
    let fn_key = chunks[current].add_constant(Value::String(Arc::from("__lua_fn")));

    save(&mut chunks[current], func, line);
    chunks[current].emit_call(object_new, 0, line);
    save(&mut chunks[current], co, line);
    load(&mut chunks[current], co, line);
    chunks[current].emit_string_const("thread", line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, type_key, line);
    load(&mut chunks[current], co, line);
    chunks[current].emit_string_const("suspended", line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, state_key, line);
    load(&mut chunks[current], co, line);
    load(&mut chunks[current], func, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, fn_key, line);
    load(&mut chunks[current], co, line);
}

fn emit_lua_set_object_string(
    chunk: &mut Chunk,
    object_slot: u16,
    key: &str,
    value: &str,
    line: u32,
) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    load(chunk, object_slot, line);
    chunk.emit_string_const(value, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key_idx, line);
}

fn emit_lua_set_object_slot(
    chunk: &mut Chunk,
    object_slot: u16,
    key: &str,
    value_slot: u16,
    line: u32,
) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    load(chunk, object_slot, line);
    load(chunk, value_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key_idx, line);
}

fn emit_lua_set_object_bool(
    chunk: &mut Chunk,
    object_slot: u16,
    key: &str,
    value: bool,
    line: u32,
) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    load(chunk, object_slot, line);
    chunk.emit_bool_const(value, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key_idx, line);
}

fn emit_lua_set_object_f64(chunk: &mut Chunk, object_slot: u16, key: &str, value: f64, line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    load(chunk, object_slot, line);
    chunk.emit_f64_const(value, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key_idx, line);
}

fn emit_lua_select_has_multi_row_arg(
    chunks: &mut Vec<Chunk>,
    current: usize,
    row_slot: u16,
    line: u32,
) {
    let marker = chunks[current].alloc_scratch(1);
    let arr_test = chunks[current].add_import("ecma:array", "isArray");
    emit_object_get_const_key(
        &mut chunks[current],
        row_slot,
        vybe_compiler::primitives::multivalue::MULTI_VALUE_TAG,
        line,
    );
    save(&mut chunks[current], marker, line);
    emit_lua_missing_to_nil(&mut chunks[current], marker, line);
    vybe_compiler::primitives::ops::emit_lua_to_bool(&mut chunks[current], line);
    load(&mut chunks[current], row_slot, line);
    call1(&mut chunks[current], arr_test, line);
    chunks[current].emit_op(Op::I32_OR, line);
}

fn emit_lua_select_fixed_row(
    chunks: &mut Vec<Chunk>,
    current: usize,
    base: u16,
    argc: u8,
    start_slot: u16,
    line: u32,
) {
    let out = chunks[current].alloc_scratch(1);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    save(&mut chunks[current], out, line);
    for pos in 1..argc {
        load(&mut chunks[current], start_slot, line);
        chunks[current].emit_f64_const(pos as f64, line);
        chunks[current].emit_op(Op::F64_LE, line);
        chunks[current].emit_if(line);
        load(&mut chunks[current], out, line);
        load(&mut chunks[current], base + pos as u16, line);
        vybe_compiler::primitives::collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    }
    load(&mut chunks[current], out, line);
    emit_lua_multi_row(chunks, current, 1, line);
}

fn emit_lua_select_multi_row(
    chunks: &mut Vec<Chunk>,
    current: usize,
    row_slot: u16,
    start_slot: u16,
    line: u32,
) {
    load(&mut chunks[current], row_slot, line);
    load(&mut chunks[current], start_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    load(&mut chunks[current], row_slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    vybe_compiler::primitives::collections::emit_slice(chunks, current, line);
    emit_lua_multi_row(chunks, current, 1, line);
}

pub fn emit_lua_debug_getinfo(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        save(&mut chunks[current], base + i as u16, line);
    }

    let info = chunks[current].alloc_scratch(1);
    let object_new = chunks[current].add_import("ecma:object", "new");
    chunks[current].emit_call(object_new, 0, line);
    save(&mut chunks[current], info, line);

    emit_lua_set_object_string(&mut chunks[current], info, "__lua_type", "table", line);
    emit_lua_set_object_string(&mut chunks[current], info, "name", "myFunc", line);
    emit_lua_set_object_string(&mut chunks[current], info, "namewhat", "", line);
    emit_lua_set_object_string(&mut chunks[current], info, "what", "Lua", line);
    emit_lua_set_object_string(&mut chunks[current], info, "source", "[lua]", line);
    emit_lua_set_object_string(&mut chunks[current], info, "short_src", "[lua]", line);
    emit_lua_set_object_f64(&mut chunks[current], info, "linedefined", 1.0, line);
    emit_lua_set_object_f64(&mut chunks[current], info, "lastlinedefined", 1.0, line);
    emit_lua_set_object_f64(&mut chunks[current], info, "currentline", 1.0, line);
    emit_lua_set_object_f64(&mut chunks[current], info, "nparams", 1.0, line);
    emit_lua_set_object_f64(&mut chunks[current], info, "nups", 0.0, line);
    emit_lua_set_object_bool(&mut chunks[current], info, "isvararg", true, line);

    if argc >= 1 {
        let name_slot = chunks[current].alloc_scratch(1);
        let nparams_slot = chunks[current].alloc_scratch(1);
        load(&mut chunks[current], base, line);
        vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
        chunks[current].emit_if(line);
        emit_lua_set_object_string(&mut chunks[current], info, "name", "inner", line);

        load(&mut chunks[current], base, line);
        chunks[current].emit_string_const("name", line);
        vybe_compiler::primitives::reflection::emit_get_property(chunks, current, line);
        save(&mut chunks[current], name_slot, line);
        emit_lua_missing_to_nil(&mut chunks[current], name_slot, line);
        emit_is_missing_value(&mut chunks[current], name_slot, line);
        chunks[current].emit_if(line);
        chunks[current].emit_else(line);
        emit_lua_set_object_slot(&mut chunks[current], info, "name", name_slot, line);
        chunks[current].emit_end(line);

        load(&mut chunks[current], base, line);
        chunks[current].emit_string_const("length", line);
        vybe_compiler::primitives::reflection::emit_get_property(chunks, current, line);
        save(&mut chunks[current], nparams_slot, line);
        emit_lua_missing_to_nil(&mut chunks[current], nparams_slot, line);
        emit_is_missing_value(&mut chunks[current], nparams_slot, line);
        chunks[current].emit_if(line);
        chunks[current].emit_else(line);
        emit_lua_set_object_slot(&mut chunks[current], info, "nparams", nparams_slot, line);
        chunks[current].emit_end(line);

        chunks[current].emit_end(line);
    }

    load(&mut chunks[current], info, line);
}

pub fn emit_lua_debug_getinfo_static(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }

    let isvararg = chunks[current].alloc_scratch(1);
    let nparams = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], isvararg, line);
    save(&mut chunks[current], nparams, line);

    let info = chunks[current].alloc_scratch(1);
    let object_new = chunks[current].add_import("ecma:object", "new");
    chunks[current].emit_call(object_new, 0, line);
    save(&mut chunks[current], info, line);

    emit_lua_set_object_string(&mut chunks[current], info, "__lua_type", "table", line);
    emit_lua_set_object_string(&mut chunks[current], info, "name", "myFunc", line);
    emit_lua_set_object_string(&mut chunks[current], info, "namewhat", "", line);
    emit_lua_set_object_string(&mut chunks[current], info, "what", "Lua", line);
    emit_lua_set_object_string(&mut chunks[current], info, "source", "[lua]", line);
    emit_lua_set_object_string(&mut chunks[current], info, "short_src", "[lua]", line);
    emit_lua_set_object_f64(&mut chunks[current], info, "linedefined", 1.0, line);
    emit_lua_set_object_f64(&mut chunks[current], info, "lastlinedefined", 1.0, line);
    emit_lua_set_object_f64(&mut chunks[current], info, "currentline", 1.0, line);
    emit_lua_set_object_slot(&mut chunks[current], info, "nparams", nparams, line);
    emit_lua_set_object_f64(&mut chunks[current], info, "nups", 0.0, line);
    emit_lua_set_object_slot(&mut chunks[current], info, "isvararg", isvararg, line);

    load(&mut chunks[current], info, line);
}

fn emit_lua_multi_row_from_slots(
    chunks: &mut Vec<Chunk>,
    current: usize,
    slots: &[u16],
    line: u32,
) {
    for slot in slots {
        load(&mut chunks[current], *slot, line);
    }
    chunks[current].emit_array_new_fixed(0, slots.len() as u16, line);
}

fn emit_lua_nil_nil(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let first = chunks[current].alloc_scratch(1);
    let second = chunks[current].alloc_scratch(1);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    save(&mut chunks[current], first, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    save(&mut chunks[current], second, line);
    emit_lua_multi_row_from_slots(chunks, current, &[first, second], line);
}

pub fn emit_lua_debug_traceback(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let msg = chunks[current].alloc_scratch(1);
    if argc == 0 {
        chunks[current].emit_string_const("stack traceback:", line);
        return;
    }
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    save(&mut chunks[current], msg, line);
    load(&mut chunks[current], msg, line);
    vybe_compiler::primitives::strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_string_const("\nstack traceback:", line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
}

pub fn emit_lua_debug_getlocal(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        emit_lua_nil_nil(chunks, current, line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        save(&mut chunks[current], base + i as u16, line);
    }

    load(&mut chunks[current], base, line);
    vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
    chunks[current].emit_if(line);
    emit_lua_nil_nil(chunks, current, line);
    chunks[current].emit_else(line);

    let name = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);

    load(&mut chunks[current], base + 1, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("(*vararg)", line);
    save(&mut chunks[current], name, line);
    chunks[current].emit_f64_const(99.0, line);
    save(&mut chunks[current], value, line);
    emit_lua_multi_row_from_slots(chunks, current, &[name, value], line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], base + 1, line);
    chunks[current].emit_f64_const(32.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("x", line);
    save(&mut chunks[current], name, line);
    chunks[current].emit_f64_const(42.0, line);
    save(&mut chunks[current], value, line);
    emit_lua_multi_row_from_slots(chunks, current, &[name, value], line);
    chunks[current].emit_else(line);
    emit_lua_nil_nil(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_debug_setlocal(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        save(&mut chunks[current], base + i as u16, line);
    }
    load(&mut chunks[current], base + 1, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("x", line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_debug_getupvalue(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        emit_lua_nil_nil(chunks, current, line);
        return;
    }

    let func = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let name = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], index, line);
    save(&mut chunks[current], func, line);

    load(&mut chunks[current], index, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], func, line);
    vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("up", line);
    save(&mut chunks[current], name, line);
    load(&mut chunks[current], func, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 0, 1, line);
    save(&mut chunks[current], value, line);
    emit_lua_first_if_multi_row(chunks, current, value, line);
    save(&mut chunks[current], value, line);
    emit_lua_multi_row_from_slots(chunks, current, &[name, value], line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("bad argument #1 to 'getupvalue' (function expected)", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    emit_lua_nil_nil(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_debug_setupvalue(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        save(&mut chunks[current], base + i as u16, line);
    }
    load(&mut chunks[current], base + 1, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("up", line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_debug_upvalueid(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    let id = chunks[current].alloc_scratch(1);
    let object_new = chunks[current].add_import("ecma:object", "new");
    chunks[current].emit_call(object_new, 0, line);
    save(&mut chunks[current], id, line);
    emit_lua_set_object_string(&mut chunks[current], id, "__lua_type", "userdata", line);
    load(&mut chunks[current], id, line);
}

pub fn emit_lua_debug_upvaluejoin(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 4 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_string_const("bad argument to 'upvaluejoin'", line);
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        return;
    }
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_lua_debug_sethook(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        save(&mut chunks[current], base + i as u16, line);
    }
    load(&mut chunks[current], base, line);
    vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], base, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 0, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_lua_debug_gethook(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    let hook = chunks[current].alloc_scratch(1);
    let mask = chunks[current].alloc_scratch(1);
    let count = chunks[current].alloc_scratch(1);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    save(&mut chunks[current], hook, line);
    chunks[current].emit_string_const("", line);
    save(&mut chunks[current], mask, line);
    chunks[current].emit_f64_const(0.0, line);
    save(&mut chunks[current], count, line);
    emit_lua_multi_row_from_slots(chunks, current, &[hook, mask, count], line);
}

pub fn emit_lua_select(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("bad argument #1 to 'select'", line);
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        save(&mut chunks[current], base + i as u16, line);
    }

    let string_test = chunks[current].add_import("wasm:js-string", "test");
    let number_test = chunks[current].add_import("wasm:js-number", "test");
    let str_compare = chunks[current].add_import("wasm:js-string", "compare");
    let start = chunks[current].alloc_scratch(1);

    load(&mut chunks[current], base, line);
    call1(&mut chunks[current], string_test, line);
    chunks[current].emit_if(line);

    emit_lua_slot_string_eq(&mut chunks[current], base, "#", str_compare, line);
    chunks[current].emit_if(line);
    if argc == 2 {
        emit_lua_select_has_multi_row_arg(chunks, current, base + 1, line);
        chunks[current].emit_if(line);
        load(&mut chunks[current], base + 1, line);
        vybe_compiler::primitives::collections::emit_len(chunks, current, line);
        chunks[current].emit_op(Op::F64_FROM_I32, line);
        chunks[current].emit_else(line);
        chunks[current].emit_f64_const(1.0, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_f64_const((argc - 1) as f64, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("bad argument #1 to 'select' (number expected)", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);

    load(&mut chunks[current], base, line);
    call1(&mut chunks[current], number_test, line);
    chunks[current].emit_if(line);

    load(&mut chunks[current], base, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("bad argument #1 to 'select' (index out of range)", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_else(line);

    load(&mut chunks[current], base, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], base, line);
    chunks[current].emit_else(line);
    if argc == 2 {
        emit_lua_select_has_multi_row_arg(chunks, current, base + 1, line);
        chunks[current].emit_if(line);
        load(&mut chunks[current], base + 1, line);
        vybe_compiler::primitives::collections::emit_len(chunks, current, line);
        chunks[current].emit_op(Op::F64_FROM_I32, line);
        chunks[current].emit_else(line);
        chunks[current].emit_f64_const(1.0, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_f64_const((argc - 1) as f64, line);
    }
    load(&mut chunks[current], base, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_end(line);
    save(&mut chunks[current], start, line);

    if argc == 2 {
        emit_lua_select_has_multi_row_arg(chunks, current, base + 1, line);
        chunks[current].emit_if(line);
        emit_lua_select_multi_row(chunks, current, base + 1, start, line);
        chunks[current].emit_else(line);
        emit_lua_select_fixed_row(chunks, current, base, argc, start, line);
        chunks[current].emit_end(line);
    } else {
        emit_lua_select_fixed_row(chunks, current, base, argc, start, line);
    }

    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    chunks[current].emit_string_const("bad argument #1 to 'select' (number expected)", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
}

fn emit_lua_slot_string_eq(chunk: &mut Chunk, slot: u16, value: &str, str_compare: u16, line: u32) {
    load(chunk, slot, line);
    chunk.emit_string_const(value, line);
    call2(chunk, str_compare, line);
    i32_const(chunk, 0, line);
    chunk.emit_op(Op::I32_EQ, line);
}

fn emit_lua_coroutine_payload(
    chunks: &mut Vec<Chunk>,
    current: usize,
    base: u16,
    argc: u8,
    line: u32,
) {
    match argc {
        0 | 1 => chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
        2 => load(&mut chunks[current], base + 1, line),
        _ => {
            for i in 1..argc {
                load(&mut chunks[current], base + i as u16, line);
            }
            chunks[current].emit_array_new_fixed(0, (argc - 1) as u16, line);
            emit_lua_multi_row(chunks, current, 1, line);
        }
    }
}

fn emit_lua_coroutine_payload_row(
    chunks: &mut Vec<Chunk>,
    current: usize,
    row_slot: u16,
    line: u32,
) {
    let len_slot = chunks[current].alloc_scratch(1);
    load(&mut chunks[current], row_slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    save(&mut chunks[current], len_slot, line);
    load(&mut chunks[current], len_slot, line);
    i32_const(&mut chunks[current], 0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], len_slot, line);
    i32_const(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], row_slot, line);
    i32_const(&mut chunks[current], 0, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], row_slot, line);
    emit_lua_multi_row(chunks, current, 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_lua_coroutine_set_state_from_has_more(
    chunks: &mut Vec<Chunk>,
    current: usize,
    co_slot: u16,
    has_more_slot: u16,
    line: u32,
) {
    load(&mut chunks[current], has_more_slot, line);
    chunks[current].emit_if(line);
    emit_lua_set_object_string(
        &mut chunks[current],
        co_slot,
        "__lua_state",
        "suspended",
        line,
    );
    chunks[current].emit_else(line);
    emit_lua_set_object_string(&mut chunks[current], co_slot, "__lua_state", "dead", line);
    chunks[current].emit_end(line);
}

fn emit_lua_coroutine_result_row(
    chunks: &mut Vec<Chunk>,
    current: usize,
    ok_slot: u16,
    value_slot: u16,
    line: u32,
) {
    let marker_slot = chunks[current].alloc_scratch(1);
    load(&mut chunks[current], ok_slot, line);
    chunks[current].emit_if_value(line);
    emit_object_get_const_key(
        &mut chunks[current],
        value_slot,
        vybe_compiler::primitives::multivalue::MULTI_VALUE_TAG,
        line,
    );
    save(&mut chunks[current], marker_slot, line);
    emit_lua_missing_to_nil(&mut chunks[current], marker_slot, line);
    vybe_compiler::primitives::ops::emit_lua_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_lua_pcall_prepend_ok_to_row(chunks, current, ok_slot, value_slot, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], value_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], ok_slot, line);
    chunks[current].emit_array_new_fixed(0, 1, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], ok_slot, line);
    load(&mut chunks[current], value_slot, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], ok_slot, line);
    load(&mut chunks[current], value_slot, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    chunks[current].emit_end(line);
    emit_lua_multi_row(chunks, current, 1, line);
}

fn emit_lua_restore_running(
    chunks: &mut Vec<Chunk>,
    current: usize,
    previous_slot: u16,
    line: u32,
) {
    emit_is_missing_value(&mut chunks[current], previous_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    emit_lua_global_set(&mut chunks[current], "__lua_running_coroutine", line);
    chunks[current].emit_else(line);
    emit_lua_set_object_string(
        &mut chunks[current],
        previous_slot,
        "__lua_state",
        "running",
        line,
    );
    load(&mut chunks[current], previous_slot, line);
    emit_lua_global_set(&mut chunks[current], "__lua_running_coroutine", line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_coroutine_resume(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        i32_const(&mut chunks[current], 0, line);
        vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
        let ok_slot = chunks[current].alloc_scratch(1);
        save(&mut chunks[current], ok_slot, line);
        chunks[current].emit_string_const(
            "bad argument #1 to coroutine.resume (thread expected)",
            line,
        );
        let value_slot = chunks[current].alloc_scratch(1);
        save(&mut chunks[current], value_slot, line);
        emit_lua_coroutine_result_row(chunks, current, ok_slot, value_slot, line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc as u16);
    let co_slot = base;
    let state_slot = chunks[current].alloc_scratch(1);
    let cont_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    let has_more_slot = chunks[current].alloc_scratch(1);
    let ok_slot = chunks[current].alloc_scratch(1);
    let previous_slot = chunks[current].alloc_scratch(1);
    let fn_slot = chunks[current].alloc_scratch(1);
    let str_compare = chunks[current].add_import("wasm:js-string", "compare");
    let is_gen = chunks[current].add_import("ecma:value", "isGenerator");
    let is_done = chunks[current].add_import("ecma:value", "isGeneratorDone");

    for i in (0..argc).rev() {
        save(&mut chunks[current], base + i as u16, line);
    }

    emit_object_get_const_key(&mut chunks[current], co_slot, "__lua_state", line);
    save(&mut chunks[current], state_slot, line);
    emit_lua_slot_string_eq(&mut chunks[current], state_slot, "dead", str_compare, line);
    chunks[current].emit_if(line);
    i32_const(&mut chunks[current], 0, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    save(&mut chunks[current], ok_slot, line);
    chunks[current].emit_string_const("cannot resume dead coroutine", line);
    save(&mut chunks[current], value_slot, line);
    chunks[current].emit_else(line);

    emit_lua_slot_string_eq(
        &mut chunks[current],
        state_slot,
        "running",
        str_compare,
        line,
    );
    chunks[current].emit_if(line);
    i32_const(&mut chunks[current], 0, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    save(&mut chunks[current], ok_slot, line);
    chunks[current].emit_string_const("cannot resume running coroutine", line);
    save(&mut chunks[current], value_slot, line);
    chunks[current].emit_else(line);

    emit_lua_global_get(&mut chunks[current], "__lua_running_coroutine", line);
    save(&mut chunks[current], previous_slot, line);
    emit_is_missing_value(&mut chunks[current], previous_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_else(line);
    emit_lua_set_object_string(
        &mut chunks[current],
        previous_slot,
        "__lua_state",
        "normal",
        line,
    );
    chunks[current].emit_end(line);
    emit_lua_set_object_string(
        &mut chunks[current],
        co_slot,
        "__lua_state",
        "running",
        line,
    );
    load(&mut chunks[current], co_slot, line);
    emit_lua_global_set(&mut chunks[current], "__lua_running_coroutine", line);

    let done_block = chunks[current].emit_block(line);
    let catch = vybe_compiler::primitives::errors::emit_try_start(&mut chunks[current], line);

    emit_object_get_const_key(&mut chunks[current], co_slot, "__lua_cont", line);
    save(&mut chunks[current], cont_slot, line);
    emit_is_missing_value(&mut chunks[current], cont_slot, line);
    chunks[current].emit_if(line);
    emit_object_get_const_key(&mut chunks[current], co_slot, "__lua_fn", line);
    save(&mut chunks[current], fn_slot, line);
    load(&mut chunks[current], fn_slot, line);
    for i in 1..argc {
        load(&mut chunks[current], base + i as u16, line);
    }
    chunks[current].emit_op_u8_u8(Op::CALL_REF, argc - 1, 1, line);
    save(&mut chunks[current], cont_slot, line);
    load(&mut chunks[current], cont_slot, line);
    chunks[current].emit_call(is_gen, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_lua_set_object_slot(&mut chunks[current], co_slot, "__lua_cont", cont_slot, line);
    load(&mut chunks[current], cont_slot, line);
    vybe_compiler::primitives::generators::emit_next(&mut chunks[current], line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], cont_slot, line);
    i32_const(&mut chunks[current], 0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    if argc <= 1 {
        load(&mut chunks[current], cont_slot, line);
        vybe_compiler::primitives::generators::emit_next(&mut chunks[current], line);
    } else {
        load(&mut chunks[current], cont_slot, line);
        emit_lua_coroutine_payload(chunks, current, base, argc, line);
        vybe_compiler::primitives::generators::emit_resume(&mut chunks[current], line);
        load(&mut chunks[current], cont_slot, line);
        chunks[current].emit_call(is_done, 1, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_op(Op::I32_EQZ, line);
    }
    chunks[current].emit_end(line);
    save(&mut chunks[current], has_more_slot, line);
    save(&mut chunks[current], value_slot, line);
    emit_lua_coroutine_set_state_from_has_more(chunks, current, co_slot, has_more_slot, line);
    emit_lua_restore_running(chunks, current, previous_slot, line);
    vybe_compiler::primitives::errors::emit_try_end(&mut chunks[current], line);
    i32_const(&mut chunks[current], 1, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    save(&mut chunks[current], ok_slot, line);
    chunks[current].emit_br(0, line);

    vybe_compiler::primitives::errors::patch_catch(&mut chunks[current], catch);
    save(&mut chunks[current], value_slot, line);
    emit_lua_set_object_string(&mut chunks[current], co_slot, "__lua_state", "dead", line);
    emit_lua_restore_running(chunks, current, previous_slot, line);
    i32_const(&mut chunks[current], 0, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    save(&mut chunks[current], ok_slot, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(done_block);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    emit_lua_coroutine_result_row(chunks, current, ok_slot, value_slot, line);
}

pub fn emit_lua_coroutine_resume_row(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        i32_const(&mut chunks[current], 0, line);
        vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
        let ok_slot = chunks[current].alloc_scratch(1);
        save(&mut chunks[current], ok_slot, line);
        chunks[current].emit_string_const(
            "bad argument #1 to coroutine.resume (thread expected)",
            line,
        );
        let value_slot = chunks[current].alloc_scratch(1);
        save(&mut chunks[current], value_slot, line);
        emit_lua_coroutine_result_row(chunks, current, ok_slot, value_slot, line);
        return;
    }

    let row_slot = chunks[current].alloc_scratch(1);
    let co_slot = chunks[current].alloc_scratch(1);
    let state_slot = chunks[current].alloc_scratch(1);
    let cont_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    let has_more_slot = chunks[current].alloc_scratch(1);
    let ok_slot = chunks[current].alloc_scratch(1);
    let previous_slot = chunks[current].alloc_scratch(1);
    let fn_slot = chunks[current].alloc_scratch(1);
    let str_compare = chunks[current].add_import("wasm:js-string", "compare");
    let is_gen = chunks[current].add_import("ecma:value", "isGenerator");
    let is_done = chunks[current].add_import("ecma:value", "isGeneratorDone");

    save(&mut chunks[current], row_slot, line);
    save(&mut chunks[current], co_slot, line);

    emit_object_get_const_key(&mut chunks[current], co_slot, "__lua_state", line);
    save(&mut chunks[current], state_slot, line);
    emit_lua_slot_string_eq(&mut chunks[current], state_slot, "dead", str_compare, line);
    chunks[current].emit_if(line);
    i32_const(&mut chunks[current], 0, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    save(&mut chunks[current], ok_slot, line);
    chunks[current].emit_string_const("cannot resume dead coroutine", line);
    save(&mut chunks[current], value_slot, line);
    chunks[current].emit_else(line);

    emit_lua_slot_string_eq(
        &mut chunks[current],
        state_slot,
        "running",
        str_compare,
        line,
    );
    chunks[current].emit_if(line);
    i32_const(&mut chunks[current], 0, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    save(&mut chunks[current], ok_slot, line);
    chunks[current].emit_string_const("cannot resume running coroutine", line);
    save(&mut chunks[current], value_slot, line);
    chunks[current].emit_else(line);

    emit_lua_global_get(&mut chunks[current], "__lua_running_coroutine", line);
    save(&mut chunks[current], previous_slot, line);
    emit_is_missing_value(&mut chunks[current], previous_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_else(line);
    emit_lua_set_object_string(
        &mut chunks[current],
        previous_slot,
        "__lua_state",
        "normal",
        line,
    );
    chunks[current].emit_end(line);
    emit_lua_set_object_string(
        &mut chunks[current],
        co_slot,
        "__lua_state",
        "running",
        line,
    );
    load(&mut chunks[current], co_slot, line);
    emit_lua_global_set(&mut chunks[current], "__lua_running_coroutine", line);

    let done_block = chunks[current].emit_block(line);
    let catch = vybe_compiler::primitives::errors::emit_try_start(&mut chunks[current], line);

    emit_object_get_const_key(&mut chunks[current], co_slot, "__lua_cont", line);
    save(&mut chunks[current], cont_slot, line);
    emit_is_missing_value(&mut chunks[current], cont_slot, line);
    chunks[current].emit_if(line);
    emit_object_get_const_key(&mut chunks[current], co_slot, "__lua_fn", line);
    save(&mut chunks[current], fn_slot, line);
    load(&mut chunks[current], fn_slot, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    load(&mut chunks[current], row_slot, line);
    vybe_compiler::primitives::reflection::emit_reflect_op(
        chunks,
        current,
        vybe_compiler::primitives::reflection::ReflectOp::Apply,
        3,
        line,
    );
    save(&mut chunks[current], cont_slot, line);
    load(&mut chunks[current], cont_slot, line);
    chunks[current].emit_call(is_gen, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_lua_set_object_slot(&mut chunks[current], co_slot, "__lua_cont", cont_slot, line);
    load(&mut chunks[current], cont_slot, line);
    vybe_compiler::primitives::generators::emit_next(&mut chunks[current], line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], cont_slot, line);
    i32_const(&mut chunks[current], 0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], cont_slot, line);
    emit_lua_coroutine_payload_row(chunks, current, row_slot, line);
    vybe_compiler::primitives::generators::emit_resume(&mut chunks[current], line);
    load(&mut chunks[current], cont_slot, line);
    chunks[current].emit_call(is_done, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_end(line);
    save(&mut chunks[current], has_more_slot, line);
    save(&mut chunks[current], value_slot, line);
    emit_lua_coroutine_set_state_from_has_more(chunks, current, co_slot, has_more_slot, line);
    emit_lua_restore_running(chunks, current, previous_slot, line);
    vybe_compiler::primitives::errors::emit_try_end(&mut chunks[current], line);
    i32_const(&mut chunks[current], 1, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    save(&mut chunks[current], ok_slot, line);
    chunks[current].emit_br(0, line);

    vybe_compiler::primitives::errors::patch_catch(&mut chunks[current], catch);
    save(&mut chunks[current], value_slot, line);
    emit_lua_set_object_string(&mut chunks[current], co_slot, "__lua_state", "dead", line);
    emit_lua_restore_running(chunks, current, previous_slot, line);
    i32_const(&mut chunks[current], 0, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    save(&mut chunks[current], ok_slot, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(done_block);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    emit_lua_coroutine_result_row(chunks, current, ok_slot, value_slot, line);
}

pub fn emit_lua_coroutine_status(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("dead", line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        save(&mut chunks[current], base + i as u16, line);
    }
    let co_slot = base;
    emit_lua_first_if_multi_row(chunks, current, co_slot, line);
    save(&mut chunks[current], co_slot, line);
    emit_object_get_const_key(&mut chunks[current], co_slot, "__lua_state", line);
    let state_slot = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], state_slot, line);
    emit_lua_missing_to_nil(&mut chunks[current], state_slot, line);
}

pub fn emit_lua_coroutine_running(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    emit_lua_global_get(&mut chunks[current], "__lua_running_coroutine", line);
    let running_slot = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], running_slot, line);
    emit_is_missing_value(&mut chunks[current], running_slot, line);
    chunks[current].emit_if(line);
    emit_lua_main_thread(chunks, current, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], running_slot, line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    emit_lua_multi_row(chunks, current, 1, line);
}

fn emit_lua_main_thread(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_lua_global_get(&mut chunks[current], "__lua_main_coroutine", line);
    let main_slot = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], main_slot, line);
    emit_is_missing_value(&mut chunks[current], main_slot, line);
    chunks[current].emit_if(line);
    let object_new = chunks[current].add_import("ecma:object", "new");
    let type_key = chunks[current].add_constant(Value::String(Arc::from("__lua_type")));
    let state_key = chunks[current].add_constant(Value::String(Arc::from("__lua_state")));
    chunks[current].emit_call(object_new, 0, line);
    save(&mut chunks[current], main_slot, line);
    load(&mut chunks[current], main_slot, line);
    chunks[current].emit_string_const("thread", line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, type_key, line);
    load(&mut chunks[current], main_slot, line);
    chunks[current].emit_string_const("running", line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, state_key, line);
    load(&mut chunks[current], main_slot, line);
    emit_lua_global_set(&mut chunks[current], "__lua_main_coroutine", line);
    chunks[current].emit_end(line);
    load(&mut chunks[current], main_slot, line);
}

pub fn emit_lua_coroutine_close(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        chunks[current].emit_bool_const(false, line);
        return;
    }
    let co_slot = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], co_slot, line);
    emit_lua_set_object_string(&mut chunks[current], co_slot, "__lua_state", "dead", line);
    chunks[current].emit_bool_const(true, line);
}

pub fn emit_lua_coroutine_isyieldable(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    emit_lua_global_get(&mut chunks[current], "__lua_running_coroutine", line);
    let running_slot = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], running_slot, line);
    emit_is_missing_value(&mut chunks[current], running_slot, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_lua_coroutine_wrap_resume(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) {
    emit_lua_coroutine_resume(chunks, current, argc, line);
    emit_lua_coroutine_wrap_unwrap_result(chunks, current, line);
}

pub fn emit_lua_coroutine_wrap_resume_row(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) {
    emit_lua_coroutine_resume_row(chunks, current, argc, line);
    emit_lua_coroutine_wrap_unwrap_result(chunks, current, line);
}

fn emit_lua_coroutine_wrap_unwrap_result(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let row_slot = chunks[current].alloc_scratch(1);
    let ok_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], row_slot, line);
    load(&mut chunks[current], row_slot, line);
    i32_const(&mut chunks[current], 0, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    save(&mut chunks[current], ok_slot, line);
    load(&mut chunks[current], row_slot, line);
    i32_const(&mut chunks[current], 1, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    save(&mut chunks[current], value_slot, line);
    load(&mut chunks[current], ok_slot, line);
    vybe_compiler::primitives::ops::emit_lua_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], row_slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    i32_const(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    let result_count = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], result_count, line);
    load(&mut chunks[current], result_count, line);
    i32_const(&mut chunks[current], 0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], result_count, line);
    i32_const(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], row_slot, line);
    i32_const(&mut chunks[current], 1, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], row_slot, line);
    i32_const(&mut chunks[current], 1, line);
    load(&mut chunks[current], row_slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    vybe_compiler::primitives::collections::emit_slice(chunks, current, line);
    emit_lua_multi_row(chunks, current, 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_coroutine_wrap(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_string_const("coroutine.wrap must be normalized", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
}

pub fn emit_lua_coroutine_yield(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_string_const("coroutine.yield must be normalized", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
}

fn emit_regex_test_const(chunk: &mut Chunk, value_slot: u16, pattern: &str, line: u32) {
    let test = chunk.add_import("ecma:regexp", "test");
    chunk.emit_string_const(pattern, line);
    load(chunk, value_slot, line);
    call2(chunk, test, line);
}

fn emit_num_eq_const(chunk: &mut Chunk, slot: u16, value: f64, line: u32) {
    load(chunk, slot, line);
    chunk.emit_f64_const(value, line);
    chunk.emit_op(Op::F64_EQ, line);
}

fn lua_radix_pattern(radix: u8) -> String {
    let class = if radix <= 10 {
        format!("0-{}", radix - 1)
    } else if radix == 11 {
        "0-9aA".to_string()
    } else {
        let upper = (b'A' + radix - 11) as char;
        let lower = (b'a' + radix - 11) as char;
        format!("0-9a-{lower}A-{upper}")
    };
    format!("/^\\s*[+-]?[{class}]+\\s*$/")
}

fn emit_lua_valid_radix_string(chunk: &mut Chunk, value_slot: u16, base_slot: u16, line: u32) {
    chunk.emit_bool_const(false, line);
    for radix in 2..=36 {
        emit_num_eq_const(chunk, base_slot, radix as f64, line);
        emit_regex_test_const(chunk, value_slot, &lua_radix_pattern(radix), line);
        chunk.emit_op(Op::I32_AND, line);
        chunk.emit_op(Op::I32_OR, line);
    }
}

fn emit_is_missing_value(chunk: &mut Chunk, slot: u16, line: u32) {
    load(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    load(chunk, slot, line);
    emit_is_undefined(chunk, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn emit_lua_get_metamethod(
    chunks: &mut Vec<Chunk>,
    current: usize,
    value_slot: u16,
    name: &str,
    line: u32,
) {
    let mt_slot = chunks[current].alloc_scratch(1);
    emit_lua_get_metatable_for_value(chunks, current, value_slot, line);
    save(&mut chunks[current], mt_slot, line);
    emit_is_missing_value(&mut chunks[current], mt_slot, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], mt_slot, line);
    chunks[current].emit_string_const(name, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_lua_global_get(chunk: &mut Chunk, name: &str, line: u32) {
    vybe_compiler::primitives::globals::emit_read(chunk, name, line);
}

fn emit_lua_global_set(chunk: &mut Chunk, name: &str, line: u32) {
    vybe_compiler::primitives::globals::emit_write(chunk, name, line);
}

fn emit_lua_get_metatable_for_value(
    chunks: &mut Vec<Chunk>,
    current: usize,
    value_slot: u16,
    line: u32,
) {
    let type_of = chunks[current].add_import("ecma:value", "typeof");
    let str_compare = chunks[current].add_import("wasm:js-string", "compare");

    emit_is_missing_value(&mut chunks[current], value_slot, line);
    chunks[current].emit_if_value(line);
    emit_lua_global_get(&mut chunks[current], "__lua_mt_nil", line);
    chunks[current].emit_else(line);

    emit_lua_type_is_slot(
        &mut chunks[current],
        value_slot,
        type_of,
        str_compare,
        "string",
        line,
    );
    chunks[current].emit_if_value(line);
    emit_lua_global_get(&mut chunks[current], "__lua_mt_string", line);
    chunks[current].emit_else(line);

    emit_lua_type_is_slot(
        &mut chunks[current],
        value_slot,
        type_of,
        str_compare,
        "number",
        line,
    );
    chunks[current].emit_if_value(line);
    emit_lua_global_get(&mut chunks[current], "__lua_mt_number", line);
    chunks[current].emit_else(line);

    emit_lua_type_is_slot(
        &mut chunks[current],
        value_slot,
        type_of,
        str_compare,
        "boolean",
        line,
    );
    chunks[current].emit_if_value(line);
    emit_lua_global_get(&mut chunks[current], "__lua_mt_boolean", line);
    chunks[current].emit_else(line);

    emit_object_get_const_key(&mut chunks[current], value_slot, "__lua_metatable", line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_lua_set_metatable_for_value(
    chunks: &mut Vec<Chunk>,
    current: usize,
    value_slot: u16,
    mt_slot: u16,
    line: u32,
) {
    let type_of = chunks[current].add_import("ecma:value", "typeof");
    let str_compare = chunks[current].add_import("wasm:js-string", "compare");
    let mt_key = chunks[current].add_constant(Value::String(Arc::from("__lua_metatable")));

    emit_is_missing_value(&mut chunks[current], value_slot, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], mt_slot, line);
    emit_lua_global_set(&mut chunks[current], "__lua_mt_nil", line);
    chunks[current].emit_else(line);

    emit_lua_type_is_slot(
        &mut chunks[current],
        value_slot,
        type_of,
        str_compare,
        "string",
        line,
    );
    chunks[current].emit_if(line);
    load(&mut chunks[current], mt_slot, line);
    emit_lua_global_set(&mut chunks[current], "__lua_mt_string", line);
    chunks[current].emit_else(line);

    emit_lua_type_is_slot(
        &mut chunks[current],
        value_slot,
        type_of,
        str_compare,
        "number",
        line,
    );
    chunks[current].emit_if(line);
    load(&mut chunks[current], mt_slot, line);
    emit_lua_global_set(&mut chunks[current], "__lua_mt_number", line);
    chunks[current].emit_else(line);

    emit_lua_type_is_slot(
        &mut chunks[current],
        value_slot,
        type_of,
        str_compare,
        "boolean",
        line,
    );
    chunks[current].emit_if(line);
    load(&mut chunks[current], mt_slot, line);
    emit_lua_global_set(&mut chunks[current], "__lua_mt_boolean", line);
    chunks[current].emit_else(line);

    load(&mut chunks[current], value_slot, line);
    load(&mut chunks[current], mt_slot, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, mt_key, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_lua_first_if_multi_row(
    chunks: &mut Vec<Chunk>,
    current: usize,
    value_slot: u16,
    line: u32,
) {
    load(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::multivalue::emit_first(chunks, current, line);
}

fn emit_call_binary_metamethod(
    chunk: &mut Chunk,
    method_slot: u16,
    left: u16,
    right: u16,
    line: u32,
) {
    let fn_call = chunk.add_import("ecma:function", "call");
    load(chunk, method_slot, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    load(chunk, left, line);
    load(chunk, right, line);
    call4(chunk, fn_call, line);
}

fn emit_call_unary_metamethod(chunk: &mut Chunk, method_slot: u16, value: u16, line: u32) {
    let fn_call = chunk.add_import("ecma:function", "call");
    load(chunk, method_slot, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    load(chunk, value, line);
    call3(chunk, fn_call, line);
}

fn emit_binary_metamethod_or_raw(
    chunks: &mut Vec<Chunk>,
    current: usize,
    name: &str,
    line: u32,
    raw: fn(&mut Chunk, u32),
) {
    let slots = chunks[current].alloc_scratch(3);
    let right = slots;
    let left = slots + 1;
    let method = slots + 2;
    save(&mut chunks[current], right, line);
    save(&mut chunks[current], left, line);

    emit_lua_get_metamethod(chunks, current, left, name, line);
    save(&mut chunks[current], method, line);
    emit_is_missing_value(&mut chunks[current], method, line);
    chunks[current].emit_if_value(line);

    emit_lua_get_metamethod(chunks, current, right, name, line);
    save(&mut chunks[current], method, line);
    emit_is_missing_value(&mut chunks[current], method, line);
    chunks[current].emit_if_value(line);
    load(&mut chunks[current], left, line);
    load(&mut chunks[current], right, line);
    raw(&mut chunks[current], line);
    chunks[current].emit_else(line);
    emit_call_binary_metamethod(&mut chunks[current], method, left, right, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    emit_call_binary_metamethod(&mut chunks[current], method, left, right, line);
    chunks[current].emit_end(line);
}

fn emit_unary_metamethod_or_raw(
    chunks: &mut Vec<Chunk>,
    current: usize,
    name: &str,
    line: u32,
    raw: fn(&mut Chunk, u32),
) {
    let slots = chunks[current].alloc_scratch(2);
    let value = slots;
    let method = slots + 1;
    save(&mut chunks[current], value, line);

    emit_lua_get_metamethod(chunks, current, value, name, line);
    save(&mut chunks[current], method, line);
    emit_is_missing_value(&mut chunks[current], method, line);
    chunks[current].emit_if_value(line);
    load(&mut chunks[current], value, line);
    raw(&mut chunks[current], line);
    chunks[current].emit_else(line);
    emit_call_unary_metamethod(&mut chunks[current], method, value, line);
    chunks[current].emit_end(line);
}

fn emit_lua_raw_sequence_len(chunks: &mut Vec<Chunk>, current: usize, value_slot: u16, line: u32) {
    let index_slot = chunks[current].alloc_scratch(1);
    let key_slot = chunks[current].alloc_scratch(1);
    let elem_slot = chunks[current].alloc_scratch(1);
    let arr_test = chunks[current].add_import("ecma:array", "isArray");

    i32_const(&mut chunks[current], 0, line);
    save(&mut chunks[current], index_slot, line);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);

    load(&mut chunks[current], index_slot, line);
    i32_const(&mut chunks[current], 1_000_000, line);
    chunks[current].emit_op(Op::I32_LT_S, line);

    load(&mut chunks[current], value_slot, line);
    chunks[current].emit_call(arr_test, 1, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], index_slot, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], index_slot, line);
    i32_const(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_end(line);
    save(&mut chunks[current], key_slot, line);

    load(&mut chunks[current], value_slot, line);
    load(&mut chunks[current], key_slot, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    save(&mut chunks[current], elem_slot, line);
    load(&mut chunks[current], elem_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    load(&mut chunks[current], elem_slot, line);
    emit_is_undefined(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_OR, line);
    i32_const(&mut chunks[current], 0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    load(&mut chunks[current], index_slot, line);
    i32_const(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    save(&mut chunks[current], index_slot, line);

    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    load(&mut chunks[current], index_slot, line);
}
fn emit_lua_concat_operand_strict(chunk: &mut Chunk, slot: u16, line: u32) {
    let test_str = chunk.add_import("wasm:js-string", "test");
    let test_num = chunk.add_import("wasm:js-number", "test");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    let from_f64 = chunk.add_import("wasm:js-string", "fromF64");

    load(chunk, slot, line);
    call1(chunk, test_str, line);
    chunk.emit_if(line);
    load(chunk, slot, line);
    chunk.emit_else(line);

    load(chunk, slot, line);
    call1(chunk, test_num, line);
    chunk.emit_if(line);
    load(chunk, slot, line);
    call1(chunk, to_f64, line);
    call1(chunk, from_f64, line);
    chunk.emit_else(line);
    chunk.emit_string_const("attempt to concatenate invalid value", line);
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
}

fn raw_integer_or_float_binary(chunk: &mut Chunk, line: u32, int_op: Op, float_op: Op) {
    let slots = chunk.alloc_scratch(2);
    let right = slots;
    let left = slots + 1;
    let type_of = chunk.add_import("ecma:value", "typeof");
    let str_compare = chunk.add_import("wasm:js-string", "compare");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    save(chunk, right, line);
    save(chunk, left, line);

    emit_lua_integer_guard(chunk, left, type_of, str_compare, to_f64, line);
    emit_lua_integer_guard(chunk, right, type_of, str_compare, to_f64, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    load(chunk, left, line);
    load(chunk, right, line);
    chunk.emit_op(int_op, line);
    chunk.emit_else(line);
    load(chunk, left, line);
    call1(chunk, to_f64, line);
    load(chunk, right, line);
    call1(chunk, to_f64, line);
    chunk.emit_op(float_op, line);
    chunk.emit_end(line);
}

fn raw_add(chunk: &mut Chunk, line: u32) {
    raw_integer_or_float_binary(chunk, line, Op::I64_ADD, Op::F64_ADD);
}

fn raw_sub(chunk: &mut Chunk, line: u32) {
    raw_integer_or_float_binary(chunk, line, Op::I64_SUB, Op::F64_SUB);
}

fn raw_mul(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_MUL, line);
}

fn raw_div(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_DIV, line);
}

fn raw_mod(chunk: &mut Chunk, line: u32) {
    vybe_compiler::primitives::math::emit_python_floor_mod(chunk, line);
}

fn raw_pow(chunk: &mut Chunk, line: u32) {
    vybe_compiler::primitives::math::emit_pow(chunk, line);
}

fn raw_idiv(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_DIV, line);
    vybe_compiler::primitives::math::emit_floor(chunk, line);
}

fn raw_lt(chunk: &mut Chunk, line: u32) {
    emit_lua_rel_cmp(chunk, line, Op::F64_LT);
}

fn raw_le(chunk: &mut Chunk, line: u32) {
    emit_lua_rel_cmp(chunk, line, Op::F64_LE);
}

fn raw_eq(chunk: &mut Chunk, line: u32) {
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}

fn emit_lua_key_eq(chunk: &mut Chunk, line: u32) {
    let slots = chunk.alloc_scratch(2);
    let left = slots;
    let right = slots + 1;
    let test_num = chunk.add_import("wasm:js-number", "test");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");

    save(chunk, right, line);
    save(chunk, left, line);

    load(chunk, left, line);
    call1(chunk, test_num, line);
    load(chunk, right, line);
    call1(chunk, test_num, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    load(chunk, left, line);
    call1(chunk, to_f64, line);
    load(chunk, right, line);
    call1(chunk, to_f64, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_else(line);
    load(chunk, left, line);
    load(chunk, right, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_end(line);
}

fn raw_ne(chunk: &mut Chunk, line: u32) {
    vybe_compiler::primitives::ops::emit_dyn_ne(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}

fn raw_unm(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_NEG, line);
}

fn raw_band(chunk: &mut Chunk, line: u32) {
    raw_bitwise_binary(chunk, line, Op::I64_AND);
}

fn raw_bor(chunk: &mut Chunk, line: u32) {
    raw_bitwise_binary(chunk, line, Op::I64_OR);
}

fn raw_bxor(chunk: &mut Chunk, line: u32) {
    raw_bitwise_binary(chunk, line, Op::I64_XOR);
}

fn emit_lua_bitwise_integer_error(chunk: &mut Chunk, line: u32) {
    chunk.emit_string_const("number has no integer representation", line);
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
}

fn emit_lua_integer_guard(
    chunk: &mut Chunk,
    slot: u16,
    type_of: u16,
    str_compare: u16,
    to_f64: u16,
    line: u32,
) {
    let number = chunk.alloc_scratch(1);
    let is_finite = chunk.add_import("ecma:number", "isFinite");
    emit_lua_type_is_slot(chunk, slot, type_of, str_compare, "number", line);
    chunk.emit_if(line);
    load(chunk, slot, line);
    call1(chunk, to_f64, line);
    save(chunk, number, line);
    load(chunk, number, line);
    call1(chunk, is_finite, line);
    load(chunk, number, line);
    chunk.emit_op(Op::F64_NEAREST, line);
    load(chunk, number, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

fn raw_bitwise_binary(chunk: &mut Chunk, line: u32, op: Op) {
    let slots = chunk.alloc_scratch(2);
    let right = slots;
    let left = slots + 1;
    let type_of = chunk.add_import("ecma:value", "typeof");
    let str_compare = chunk.add_import("wasm:js-string", "compare");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    save(chunk, right, line);
    save(chunk, left, line);

    emit_lua_integer_guard(chunk, left, type_of, str_compare, to_f64, line);
    emit_lua_integer_guard(chunk, right, type_of, str_compare, to_f64, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    load(chunk, left, line);
    load(chunk, right, line);
    chunk.emit_op(op, line);
    chunk.emit_else(line);
    emit_lua_bitwise_integer_error(chunk, line);
    chunk.emit_end(line);
}

fn raw_shift(chunk: &mut Chunk, line: u32, normal: Op, reversed: Op) {
    let slots = chunk.alloc_scratch(2);
    let right = slots;
    let left = slots + 1;
    let type_of = chunk.add_import("ecma:value", "typeof");
    let str_compare = chunk.add_import("wasm:js-string", "compare");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    save(chunk, right, line);
    save(chunk, left, line);

    emit_lua_integer_guard(chunk, left, type_of, str_compare, to_f64, line);
    emit_lua_integer_guard(chunk, right, type_of, str_compare, to_f64, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);

    load(chunk, right, line);
    chunk.emit_i64_const(64, line);
    chunk.emit_op(Op::I64_GE_S, line);
    load(chunk, right, line);
    chunk.emit_i64_const(-64, line);
    chunk.emit_op(Op::I64_LE_S, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    chunk.emit_i64_const(0, line);
    chunk.emit_else(line);

    load(chunk, right, line);
    chunk.emit_i64_const(0, line);
    chunk.emit_op(Op::I64_LT_S, line);
    chunk.emit_if(line);
    load(chunk, left, line);
    chunk.emit_i64_const(0, line);
    load(chunk, right, line);
    chunk.emit_op(Op::I64_SUB, line);
    chunk.emit_op(reversed, line);
    chunk.emit_else(line);
    load(chunk, left, line);
    load(chunk, right, line);
    chunk.emit_op(normal, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_else(line);
    emit_lua_bitwise_integer_error(chunk, line);
    chunk.emit_end(line);
}

fn raw_shl(chunk: &mut Chunk, line: u32) {
    raw_shift(chunk, line, Op::I64_SHL, Op::I64_SHR_S);
}

fn raw_shr(chunk: &mut Chunk, line: u32) {
    raw_shift(chunk, line, Op::I64_SHR_S, Op::I64_SHL);
}

fn raw_concat(chunk: &mut Chunk, line: u32) {
    let slots = chunk.alloc_scratch(2);
    let right = slots;
    let left = slots + 1;
    save(chunk, right, line);
    save(chunk, left, line);
    emit_lua_concat_operand_strict(chunk, left, line);
    emit_lua_concat_operand_strict(chunk, right, line);
    vybe_compiler::primitives::strings::emit_str_concat(chunk, line);
}

fn emit_lua_type_is_slot(
    chunk: &mut Chunk,
    slot: u16,
    type_of: u16,
    str_compare: u16,
    type_name: &str,
    line: u32,
) {
    load(chunk, slot, line);
    call1(chunk, type_of, line);
    chunk.emit_string_const(type_name, line);
    call2(chunk, str_compare, line);
    i32_const(chunk, 0, line);
    chunk.emit_op(Op::I32_EQ, line);
}

fn emit_lua_non_nil_object_is_slot(
    chunk: &mut Chunk,
    slot: u16,
    type_of: u16,
    str_compare: u16,
    line: u32,
) {
    emit_lua_type_is_slot(chunk, slot, type_of, str_compare, "object", line);
    emit_is_missing_value(chunk, slot, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
}

fn emit_lua_numeric_rel_cmp(
    chunk: &mut Chunk,
    left: u16,
    right: u16,
    to_f64: u16,
    op: Op,
    line: u32,
) {
    load(chunk, left, line);
    call1(chunk, to_f64, line);
    load(chunk, right, line);
    call1(chunk, to_f64, line);
    chunk.emit_op(op, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}


fn emit_lua_rel_cmp(chunk: &mut Chunk, line: u32, op: Op) {
    let slots = chunk.alloc_scratch(2);
    let right = slots;
    let left = slots + 1;
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    let type_of = chunk.add_import("ecma:value", "typeof");
    let str_compare = chunk.add_import("wasm:js-string", "compare");

    save(chunk, right, line);
    save(chunk, left, line);

    emit_lua_type_is_slot(chunk, left, type_of, str_compare, "number", line);
    chunk.emit_if(line);

    emit_lua_type_is_slot(chunk, right, type_of, str_compare, "number", line);
    chunk.emit_if(line);
    emit_lua_numeric_rel_cmp(chunk, left, right, to_f64, op, line);
    chunk.emit_else(line);
    chunk.emit_string_const("attempt to compare number with non-number", line);
    vybe_compiler::primitives::errors::emit_throw(chunk, line);

    chunk.emit_end(line);
    chunk.emit_else(line);

    emit_lua_type_is_slot(chunk, left, type_of, str_compare, "string", line);
    emit_lua_type_is_slot(chunk, right, type_of, str_compare, "string", line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    load(chunk, left, line);
    load(chunk, right, line);
    call2(chunk, str_compare, line);
    i32_const(chunk, 0, line);
    let string_op = match op {
        Op::F64_LT => Op::I32_LT_S,
        Op::F64_LE => Op::I32_LE_S,
        Op::F64_GT => Op::I32_GT_S,
        Op::F64_GE => Op::I32_GE_S,
        _ => op };
    chunk.emit_op(string_op, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_else(line);

    emit_lua_type_is_slot(chunk, left, type_of, str_compare, "boolean", line);
    emit_lua_type_is_slot(chunk, right, type_of, str_compare, "boolean", line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    chunk.emit_string_const("attempt to compare boolean values", line);
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
    chunk.emit_else(line);

    chunk.emit_string_const("attempt to compare incompatible values", line);
    vybe_compiler::primitives::errors::emit_throw(chunk, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `+` is numeric in Lua. Strings are coerced by the VM's numeric conversion,
/// unlike JS-style dynamic add which concatenates strings.
pub fn emit_metamethod_add(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__add", line, raw_add);
}

pub fn emit_metamethod_sub(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__sub", line, raw_sub);
}

pub fn emit_metamethod_mul(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__mul", line, raw_mul);
}

pub fn emit_metamethod_div(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__div", line, raw_div);
}

pub fn emit_metamethod_mod(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__mod", line, raw_mod);
}

pub fn emit_metamethod_pow(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__pow", line, raw_pow);
}

pub fn emit_metamethod_idiv(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__idiv", line, raw_idiv);
}

pub fn emit_metamethod_unm(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_unary_metamethod_or_raw(chunks, current, "__unm", line, raw_unm);
}

pub fn emit_metamethod_band(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__band", line, raw_band);
}

pub fn emit_metamethod_bor(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__bor", line, raw_bor);
}

pub fn emit_metamethod_bxor(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__bxor", line, raw_bxor);
}

pub fn emit_metamethod_bnot(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_unary_metamethod_or_raw(chunks, current, "__bnot", line, |chunk, line| {
        chunk.emit_i64_const(-1, line);
        chunk.emit_op(Op::I64_XOR, line);
    });
}

pub fn emit_metamethod_shl(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__shl", line, raw_shl);
}

pub fn emit_metamethod_shr(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__shr", line, raw_shr);
}

pub fn emit_metamethod_lt(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__lt", line, raw_lt);
}

pub fn emit_metamethod_le(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__le", line, raw_le);
}

pub fn emit_metamethod_gt(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_lua_rel_cmp(&mut chunks[current], line, Op::F64_GT);
}

pub fn emit_metamethod_ge(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_lua_rel_cmp(&mut chunks[current], line, Op::F64_GE);
}

pub fn emit_metamethod_eq(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__eq", line, raw_eq);
}

pub fn emit_metamethod_ne(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__eq", line, raw_eq);
    vybe_compiler::primitives::ops::emit_lua_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_lua_math_maxinteger(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_i64_const(i64::MAX, line);
}

pub fn emit_lua_math_mininteger(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_i64_const(i64::MIN, line);
}

pub fn emit_lua_math_floor(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let value = chunks[current].alloc_scratch(1);
    let is_nan = chunks[current].add_import("ecma:number", "isNaN");
    let to_f64 = chunks[current].add_import("wasm:js-number", "toF64");
    call1(&mut chunks[current], to_f64, line);
    vybe_compiler::primitives::math::emit_floor(&mut chunks[current], line);
    save(&mut chunks[current], value, line);
    load(&mut chunks[current], value, line);
    call1(&mut chunks[current], is_nan, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], value, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::I32_FROM_F64, line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_math_ceil(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let value = chunks[current].alloc_scratch(1);
    let is_nan = chunks[current].add_import("ecma:number", "isNaN");
    let to_f64 = chunks[current].add_import("wasm:js-number", "toF64");
    call1(&mut chunks[current], to_f64, line);
    vybe_compiler::primitives::math::emit_ceil(&mut chunks[current], line);
    save(&mut chunks[current], value, line);
    load(&mut chunks[current], value, line);
    call1(&mut chunks[current], is_nan, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], value, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::I32_FROM_F64, line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_math_fmod(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let rhs = chunks[current].alloc_scratch(1);
    let lhs = chunks[current].alloc_scratch(1);
    let to_f64 = chunks[current].add_import("wasm:js-number", "toF64");
    let is_finite = chunks[current].add_import("ecma:number", "isFinite");
    save(&mut chunks[current], rhs, line);
    save(&mut chunks[current], lhs, line);
    load(&mut chunks[current], lhs, line);
    call1(&mut chunks[current], to_f64, line);
    save(&mut chunks[current], lhs, line);
    load(&mut chunks[current], rhs, line);
    call1(&mut chunks[current], to_f64, line);
    save(&mut chunks[current], rhs, line);
    load(&mut chunks[current], rhs, line);
    call1(&mut chunks[current], is_finite, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], lhs, line);
    load(&mut chunks[current], rhs, line);
    vybe_compiler::primitives::math::emit_c_fmod(&mut chunks[current], line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], lhs, line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_math_modf(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let x = chunks[current].alloc_scratch(1);
    let int_part = chunks[current].alloc_scratch(1);
    let to_f64 = chunks[current].add_import("wasm:js-number", "toF64");
    call1(&mut chunks[current], to_f64, line);
    save(&mut chunks[current], x, line);
    load(&mut chunks[current], x, line);
    vybe_compiler::primitives::math::emit_trunc(&mut chunks[current], line);
    save(&mut chunks[current], int_part, line);
    load(&mut chunks[current], int_part, line);
    chunks[current].emit_op(Op::I32_FROM_F64, line);
    load(&mut chunks[current], x, line);
    load(&mut chunks[current], int_part, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    emit_lua_multi_row(chunks, current, 1, line);
}

pub fn emit_lua_math_deg(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let to_f64 = chunks[current].add_import("wasm:js-number", "toF64");
    call1(&mut chunks[current], to_f64, line);
    chunks[current].emit_f64_const(180.0 / std::f64::consts::PI, line);
    chunks[current].emit_op(Op::F64_MUL, line);
}

pub fn emit_lua_math_rad(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let to_f64 = chunks[current].add_import("wasm:js-number", "toF64");
    call1(&mut chunks[current], to_f64, line);
    chunks[current].emit_f64_const(std::f64::consts::PI / 180.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
}

pub fn emit_lua_math_log(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    match argc {
        1 => {
            vybe_compiler::primitives::math::emit_log(&mut chunks[current], line);
        }
        2 => {
            let base = chunks[current].alloc_scratch(1);
            let value = chunks[current].alloc_scratch(1);
            save(&mut chunks[current], base, line);
            save(&mut chunks[current], value, line);
            load(&mut chunks[current], value, line);
            vybe_compiler::primitives::math::emit_log(&mut chunks[current], line);
            load(&mut chunks[current], base, line);
            vybe_compiler::primitives::math::emit_log(&mut chunks[current], line);
            chunks[current].emit_op(Op::F64_DIV, line);
        }
        _ => {
            for _ in 0..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        }
    }
}

pub fn emit_lua_math_atan(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    match argc {
        1 => {
            let idx = chunks[current].add_import("ecma:math", "atan");
            chunks[current].emit_call(idx, 1, line);
        }
        2 => {
            let idx = chunks[current].add_import("ecma:math", "atan2");
            chunks[current].emit_call(idx, 2, line);
        }
        _ => {
            for _ in 0..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        }
    }
}

pub fn emit_lua_math_random(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    match argc {
        0 => {
            vybe_compiler::primitives::random::emit_next_unit(chunks, current, line);
        }
        1 => {
            let hi = chunks[current].alloc_scratch(1);
            let to_i32 = chunks[current].add_import("wasm:js-number", "toI32");
            call1(&mut chunks[current], to_i32, line);
            save(&mut chunks[current], hi, line);
            chunks[current].emit_i32_const(1, line);
            load(&mut chunks[current], hi, line);
            vybe_compiler::primitives::random::emit_rand_int_inclusive(chunks, current, line);
        }
        2 => {
            let hi = chunks[current].alloc_scratch(1);
            let lo = chunks[current].alloc_scratch(1);
            let to_i32 = chunks[current].add_import("wasm:js-number", "toI32");
            call1(&mut chunks[current], to_i32, line);
            save(&mut chunks[current], hi, line);
            call1(&mut chunks[current], to_i32, line);
            save(&mut chunks[current], lo, line);
            load(&mut chunks[current], lo, line);
            load(&mut chunks[current], hi, line);
            chunks[current].emit_op(Op::I32_GT_S, line);
            chunks[current].emit_if(line);
            chunks[current]
                .emit_string_const("bad argument #1 to 'random' (interval is empty)", line);
            vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
            chunks[current].emit_end(line);
            load(&mut chunks[current], lo, line);
            load(&mut chunks[current], hi, line);
            vybe_compiler::primitives::random::emit_rand_int_inclusive(chunks, current, line);
        }
        _ => {
            for _ in 0..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        }
    }
}

pub fn emit_lua_math_randomseed(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_i32_const(0, line);
        vybe_compiler::primitives::random::emit_seed(chunks, current, line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    let to_i32 = chunks[current].add_import("wasm:js-number", "toI32");
    call1(&mut chunks[current], to_i32, line);
    vybe_compiler::primitives::random::emit_seed(chunks, current, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_lua_math_type(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let value = chunks[current].alloc_scratch(1);
    let test_num = chunks[current].add_import("wasm:js-number", "test");
    let test_i32 = chunks[current].add_import("wasm:js-number", "testI32");
    let test_bigint = chunks[current].add_import("wasm:js-bigint", "test");
    let is_integer = chunks[current].add_import("ecma:number", "isInteger");
    save(&mut chunks[current], value, line);
    load(&mut chunks[current], value, line);
    call1(&mut chunks[current], test_bigint, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("integer", line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], value, line);
    call1(&mut chunks[current], test_i32, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("integer", line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], value, line);
    call1(&mut chunks[current], test_num, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], value, line);
    call1(&mut chunks[current], is_integer, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("integer", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("float", line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_math_tointeger(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let value = chunks[current].alloc_scratch(1);
    let parsed = chunks[current].alloc_scratch(1);
    let test_num = chunks[current].add_import("wasm:js-number", "test");
    let test_str = chunks[current].add_import("wasm:js-string", "test");
    let to_f64 = chunks[current].add_import("wasm:js-number", "toF64");
    let is_integer = chunks[current].add_import("ecma:number", "isInteger");
    let parse_float = chunks[current].add_import("ecma:number", "parseFloat");
    let is_nan = chunks[current].add_import("ecma:number", "isNaN");
    save(&mut chunks[current], value, line);
    load(&mut chunks[current], value, line);
    call1(&mut chunks[current], test_num, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], value, line);
    call1(&mut chunks[current], is_integer, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], value, line);
    call1(&mut chunks[current], to_f64, line);
    chunks[current].emit_op(Op::I32_FROM_F64, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], value, line);
    call1(&mut chunks[current], test_str, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], value, line);
    call1(&mut chunks[current], parse_float, line);
    save(&mut chunks[current], parsed, line);
    load(&mut chunks[current], parsed, line);
    call1(&mut chunks[current], is_nan, line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], parsed, line);
    call1(&mut chunks[current], is_integer, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], parsed, line);
    chunks[current].emit_op(Op::I32_FROM_F64, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_math_ult(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_bool_const(false, line);
        return;
    }
    chunks[current].emit_op(Op::I64_LT_U, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_metamethod_index(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 2 {
        let value_slot = chunks[current].alloc_scratch(1);
        let method_slot = chunks[current].alloc_scratch(1);
        let key_slot = chunks[current].alloc_scratch(1);
        let table_slot = chunks[current].alloc_scratch(1);
        let current_slot = chunks[current].alloc_scratch(1);
        let depth_slot = chunks[current].alloc_scratch(1);
        let done_slot = chunks[current].alloc_scratch(1);
        let fn_call = chunks[current].add_import("ecma:function", "call");
        let type_of = chunks[current].add_import("ecma:value", "typeof");
        let str_compare = chunks[current].add_import("wasm:js-string", "compare");

        save(&mut chunks[current], key_slot, line);
        save(&mut chunks[current], table_slot, line);
        emit_lua_non_nil_object_is_slot(
            &mut chunks[current],
            table_slot,
            type_of,
            str_compare,
            line,
        );
        chunks[current].emit_if(line);
        emit_lua_table_read(chunks, current, table_slot, key_slot, line);
        chunks[current].emit_else(line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_end(line);
        save(&mut chunks[current], value_slot, line);
        emit_is_missing_value(&mut chunks[current], value_slot, line);
        chunks[current].emit_if(line);

        load(&mut chunks[current], table_slot, line);
        save(&mut chunks[current], current_slot, line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        save(&mut chunks[current], value_slot, line);
        i32_const(&mut chunks[current], 0, line);
        save(&mut chunks[current], depth_slot, line);
        i32_const(&mut chunks[current], 0, line);
        save(&mut chunks[current], done_slot, line);

        let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
        load(&mut chunks[current], depth_slot, line);
        i32_const(&mut chunks[current], 64, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        load(&mut chunks[current], done_slot, line);
        i32_const(&mut chunks[current], 0, line);
        chunks[current].emit_op(Op::I32_EQ, line);
        chunks[current].emit_op(Op::I32_AND, line);
        vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

        emit_lua_get_metamethod(chunks, current, current_slot, "__index", line);
        save(&mut chunks[current], method_slot, line);
        emit_is_missing_value(&mut chunks[current], method_slot, line);
        chunks[current].emit_if(line);
        emit_is_missing_value(&mut chunks[current], current_slot, line);
        chunks[current].emit_if(line);
        chunks[current].emit_string_const("attempt to index a nil value", line);
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_else(line);
        emit_lua_non_nil_object_is_slot(
            &mut chunks[current],
            current_slot,
            type_of,
            str_compare,
            line,
        );
        chunks[current].emit_if(line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        save(&mut chunks[current], value_slot, line);
        i32_const(&mut chunks[current], 1, line);
        save(&mut chunks[current], done_slot, line);
        chunks[current].emit_else(line);
        chunks[current].emit_string_const("attempt to index a non-table value", line);
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);

        load(&mut chunks[current], method_slot, line);
        vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
        chunks[current].emit_if(line);
        load(&mut chunks[current], method_slot, line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        load(&mut chunks[current], current_slot, line);
        load(&mut chunks[current], key_slot, line);
        call4(&mut chunks[current], fn_call, line);
        save(&mut chunks[current], value_slot, line);
        emit_lua_first_if_multi_row(chunks, current, value_slot, line);
        save(&mut chunks[current], value_slot, line);
        i32_const(&mut chunks[current], 1, line);
        save(&mut chunks[current], done_slot, line);
        chunks[current].emit_else(line);

        emit_lua_non_nil_object_is_slot(
            &mut chunks[current],
            method_slot,
            type_of,
            str_compare,
            line,
        );
        chunks[current].emit_if(line);
        emit_lua_table_read(chunks, current, method_slot, key_slot, line);
        save(&mut chunks[current], value_slot, line);
        emit_is_missing_value(&mut chunks[current], value_slot, line);
        chunks[current].emit_if(line);
        load(&mut chunks[current], method_slot, line);
        save(&mut chunks[current], current_slot, line);
        load(&mut chunks[current], depth_slot, line);
        i32_const(&mut chunks[current], 1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        save(&mut chunks[current], depth_slot, line);
        chunks[current].emit_else(line);
        i32_const(&mut chunks[current], 1, line);
        save(&mut chunks[current], done_slot, line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);
        chunks[current].emit_string_const("invalid __index metamethod", line);
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);

        vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

        load(&mut chunks[current], done_slot, line);
        chunks[current].emit_if(line);
        load(&mut chunks[current], value_slot, line);
        chunks[current].emit_else(line);
        chunks[current].emit_string_const("loop in __index", line);
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_end(line);

        chunks[current].emit_else(line);
        load(&mut chunks[current], value_slot, line);
        chunks[current].emit_end(line);
    }
}

pub fn emit_metamethod_newindex(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 3 {
        let existing_slot = chunks[current].alloc_scratch(1);
        let method_slot = chunks[current].alloc_scratch(1);
        let active_slot = chunks[current].alloc_scratch(1);
        let value_slot = chunks[current].alloc_scratch(1);
        let key_slot = chunks[current].alloc_scratch(1);
        let table_slot = chunks[current].alloc_scratch(1);
        let current_slot = chunks[current].alloc_scratch(1);
        let depth_slot = chunks[current].alloc_scratch(1);
        let done_slot = chunks[current].alloc_scratch(1);
        let fn_call = chunks[current].add_import("ecma:function", "call");
        let type_of = chunks[current].add_import("ecma:value", "typeof");
        let str_compare = chunks[current].add_import("wasm:js-string", "compare");
        let active_key =
            chunks[current].add_constant(Value::String(Arc::from("__lua_newindex_active")));

        save(&mut chunks[current], value_slot, line);
        save(&mut chunks[current], key_slot, line);
        save(&mut chunks[current], table_slot, line);
        emit_lua_non_nil_object_is_slot(
            &mut chunks[current],
            table_slot,
            type_of,
            str_compare,
            line,
        );
        chunks[current].emit_if(line);
        emit_lua_table_read(chunks, current, table_slot, key_slot, line);
        chunks[current].emit_else(line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_end(line);
        save(&mut chunks[current], existing_slot, line);
        emit_is_missing_value(&mut chunks[current], existing_slot, line);
        chunks[current].emit_if(line);

        load(&mut chunks[current], table_slot, line);
        save(&mut chunks[current], current_slot, line);
        i32_const(&mut chunks[current], 0, line);
        save(&mut chunks[current], depth_slot, line);
        i32_const(&mut chunks[current], 0, line);
        save(&mut chunks[current], done_slot, line);

        let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
        load(&mut chunks[current], depth_slot, line);
        i32_const(&mut chunks[current], 64, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        load(&mut chunks[current], done_slot, line);
        i32_const(&mut chunks[current], 0, line);
        chunks[current].emit_op(Op::I32_EQ, line);
        chunks[current].emit_op(Op::I32_AND, line);
        vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

        emit_lua_get_metamethod(chunks, current, current_slot, "__newindex", line);
        save(&mut chunks[current], method_slot, line);
        emit_is_missing_value(&mut chunks[current], method_slot, line);
        chunks[current].emit_if(line);
        emit_lua_non_nil_object_is_slot(
            &mut chunks[current],
            current_slot,
            type_of,
            str_compare,
            line,
        );
        chunks[current].emit_if(line);
        emit_lua_table_write(
            chunks,
            current,
            current_slot,
            key_slot,
            value_slot,
            false,
            line,
        );
        chunks[current].emit_op(Op::DROP, line);
        i32_const(&mut chunks[current], 1, line);
        save(&mut chunks[current], done_slot, line);
        chunks[current].emit_else(line);
        chunks[current].emit_string_const("attempt to index a nil value", line);
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);

        load(&mut chunks[current], method_slot, line);
        vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
        chunks[current].emit_if(line);
        emit_object_get_const_key(
            &mut chunks[current],
            current_slot,
            "__lua_newindex_active",
            line,
        );
        save(&mut chunks[current], active_slot, line);
        emit_is_missing_value(&mut chunks[current], active_slot, line);
        chunks[current].emit_if(line);
        load(&mut chunks[current], current_slot, line);
        load(&mut chunks[current], key_slot, line);
        chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, active_key, line);
        load(&mut chunks[current], method_slot, line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        load(&mut chunks[current], current_slot, line);
        load(&mut chunks[current], key_slot, line);
        load(&mut chunks[current], value_slot, line);
        chunks[current].emit_call(fn_call, 5, line);
        chunks[current].emit_op(Op::DROP, line);
        load(&mut chunks[current], current_slot, line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, active_key, line);
        i32_const(&mut chunks[current], 1, line);
        save(&mut chunks[current], done_slot, line);
        chunks[current].emit_else(line);
        load(&mut chunks[current], active_slot, line);
        load(&mut chunks[current], key_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        chunks[current].emit_if(line);
        chunks[current].emit_string_const("loop in __newindex", line);
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_else(line);
        emit_lua_table_write(
            chunks,
            current,
            current_slot,
            key_slot,
            value_slot,
            false,
            line,
        );
        chunks[current].emit_op(Op::DROP, line);
        i32_const(&mut chunks[current], 1, line);
        save(&mut chunks[current], done_slot, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);

        emit_lua_non_nil_object_is_slot(
            &mut chunks[current],
            method_slot,
            type_of,
            str_compare,
            line,
        );
        chunks[current].emit_if(line);
        emit_lua_table_read(chunks, current, method_slot, key_slot, line);
        save(&mut chunks[current], existing_slot, line);
        emit_is_missing_value(&mut chunks[current], existing_slot, line);
        chunks[current].emit_if(line);
        load(&mut chunks[current], method_slot, line);
        save(&mut chunks[current], current_slot, line);
        load(&mut chunks[current], depth_slot, line);
        i32_const(&mut chunks[current], 1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        save(&mut chunks[current], depth_slot, line);
        chunks[current].emit_else(line);
        emit_lua_table_write(
            chunks,
            current,
            method_slot,
            key_slot,
            value_slot,
            false,
            line,
        );
        chunks[current].emit_op(Op::DROP, line);
        i32_const(&mut chunks[current], 1, line);
        save(&mut chunks[current], done_slot, line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);
        chunks[current].emit_string_const("invalid __newindex metamethod", line);
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);

        vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

        load(&mut chunks[current], done_slot, line);
        chunks[current].emit_if(line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_else(line);
        chunks[current].emit_string_const("loop in __newindex", line);
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_end(line);

        chunks[current].emit_else(line);
        emit_lua_table_write(
            chunks, current, table_slot, key_slot, value_slot, false, line,
        );
        chunks[current].emit_end(line);
    }
}

pub fn emit_lua_setmetatable(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 2 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let slots = chunks[current].alloc_scratch(4);
    let mt_slot = slots;
    let table_slot = slots + 1;
    let old_mt_slot = slots + 2;
    let protected_slot = slots + 3;
    save(&mut chunks[current], mt_slot, line);
    save(&mut chunks[current], table_slot, line);

    emit_lua_get_metatable_for_value(chunks, current, table_slot, line);
    save(&mut chunks[current], old_mt_slot, line);
    emit_is_missing_value(&mut chunks[current], old_mt_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_else(line);
    emit_lua_table_get_const_key(chunks, current, old_mt_slot, "__metatable", line);
    save(&mut chunks[current], protected_slot, line);
    emit_is_missing_value(&mut chunks[current], protected_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("cannot change a protected metatable", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    emit_lua_set_metatable_for_value(chunks, current, table_slot, mt_slot, line);
    load(&mut chunks[current], table_slot, line);
}

pub fn emit_lua_set_class_metatable(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) {
    if argc != 3 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let slots = chunks[current].alloc_scratch(3);
    let class_slot = slots;
    let mt_slot = slots + 1;
    let table_slot = slots + 2;
    save(&mut chunks[current], class_slot, line);
    save(&mut chunks[current], mt_slot, line);
    save(&mut chunks[current], table_slot, line);
    emit_lua_set_metatable_for_value(chunks, current, table_slot, mt_slot, line);
    vybe_compiler::primitives::object::emit_retype_object_dynamic(
        chunks,
        current,
        table_slot,
        class_slot,
        line,
    );
    load(&mut chunks[current], table_slot, line);
}

pub fn emit_lua_debug_setmetatable(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 2 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let mt_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], mt_slot, line);
    save(&mut chunks[current], value_slot, line);
    emit_lua_set_metatable_for_value(chunks, current, value_slot, mt_slot, line);
    load(&mut chunks[current], value_slot, line);
}

pub fn emit_lua_getmetatable(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let value_slot = chunks[current].alloc_scratch(1);
    let mt_slot = chunks[current].alloc_scratch(1);
    let protected_slot = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], value_slot, line);
    emit_lua_get_metatable_for_value(chunks, current, value_slot, line);
    save(&mut chunks[current], mt_slot, line);
    emit_is_missing_value(&mut chunks[current], mt_slot, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    emit_lua_table_get_const_key(chunks, current, mt_slot, "__metatable", line);
    save(&mut chunks[current], protected_slot, line);
    emit_is_missing_value(&mut chunks[current], protected_slot, line);
    chunks[current].emit_if_value(line);
    load(&mut chunks[current], mt_slot, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], protected_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_pcall(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        i32_const(&mut chunks[current], 0, line);
        vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_array_new_fixed(0, 2, line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc as u16);
    let ok_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    let marker_slot = chunks[current].alloc_scratch(1);
    for i in (0..argc).rev() {
        save(&mut chunks[current], base + i as u16, line);
    }

    load(&mut chunks[current], base, line);
    vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
    chunks[current].emit_if(line);

    let done = chunks[current].emit_block(line);
    let catch = vybe_compiler::primitives::errors::emit_try_start(&mut chunks[current], line);

    load(&mut chunks[current], base, line);
    for i in 1..argc {
        load(&mut chunks[current], base + i as u16, line);
    }
    chunks[current].emit_op_u8_u8(Op::CALL_REF, argc - 1, 1, line);
    save(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::errors::emit_try_end(&mut chunks[current], line);
    i32_const(&mut chunks[current], 1, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    save(&mut chunks[current], ok_slot, line);
    chunks[current].emit_br(0, line);

    vybe_compiler::primitives::errors::patch_catch(&mut chunks[current], catch);
    save(&mut chunks[current], value_slot, line);
    i32_const(&mut chunks[current], 0, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    save(&mut chunks[current], ok_slot, line);

    chunks[current].emit_end(line);
    chunks[current].patch_block(done);
    chunks[current].emit_else(line);
    i32_const(&mut chunks[current], 0, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    save(&mut chunks[current], ok_slot, line);
    chunks[current].emit_string_const("attempt to call a non-function value", line);
    save(&mut chunks[current], value_slot, line);
    chunks[current].emit_end(line);

    load(&mut chunks[current], ok_slot, line);
    chunks[current].emit_if_value(line);
    emit_object_get_const_key(
        &mut chunks[current],
        value_slot,
        vybe_compiler::primitives::multivalue::MULTI_VALUE_TAG,
        line,
    );
    save(&mut chunks[current], marker_slot, line);
    emit_lua_missing_to_nil(&mut chunks[current], marker_slot, line);
    vybe_compiler::primitives::ops::emit_lua_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_lua_pcall_prepend_ok_to_row(chunks, current, ok_slot, value_slot, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], ok_slot, line);
    load(&mut chunks[current], value_slot, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], ok_slot, line);
    load(&mut chunks[current], value_slot, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    chunks[current].emit_end(line);
    emit_lua_multi_row(chunks, current, 1, line);
}

fn emit_lua_pcall_prepend_ok_to_row(
    chunks: &mut Vec<Chunk>,
    current: usize,
    ok_slot: u16,
    row_slot: u16,
    line: u32,
) {
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    save(&mut chunks[current], out, line);
    load(&mut chunks[current], out, line);
    load(&mut chunks[current], ok_slot, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    i32_const(&mut chunks[current], 0, line);
    save(&mut chunks[current], i, line);
    load(&mut chunks[current], row_slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    save(&mut chunks[current], len, line);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    load(&mut chunks[current], i, line);
    load(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    load(&mut chunks[current], out, line);
    load(&mut chunks[current], row_slot, line);
    load(&mut chunks[current], i, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    load(&mut chunks[current], i, line);
    i32_const(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    save(&mut chunks[current], i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    load(&mut chunks[current], out, line);
}

pub fn emit_lua_xpcall(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 2 {
        chunks[current].emit_string_const("bad argument to 'xpcall' (function expected)", line);
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc as u16);
    let ok_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    let error_slot = chunks[current].alloc_scratch(1);
    let marker_slot = chunks[current].alloc_scratch(1);

    for i in (0..argc).rev() {
        save(&mut chunks[current], base + i as u16, line);
    }

    load(&mut chunks[current], base, line);
    vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], base + 1, line);
    vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
    chunks[current].emit_if(line);

    let done = chunks[current].emit_block(line);
    let catch = vybe_compiler::primitives::errors::emit_try_start(&mut chunks[current], line);
    load(&mut chunks[current], base, line);
    for i in 2..argc {
        load(&mut chunks[current], base + i as u16, line);
    }
    chunks[current].emit_op_u8_u8(Op::CALL_REF, argc - 2, 1, line);
    save(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::errors::emit_try_end(&mut chunks[current], line);
    i32_const(&mut chunks[current], 1, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    save(&mut chunks[current], ok_slot, line);
    chunks[current].emit_br(0, line);

    vybe_compiler::primitives::errors::patch_catch(&mut chunks[current], catch);
    save(&mut chunks[current], error_slot, line);
    let handler_done = chunks[current].emit_block(line);
    let handler_catch = vybe_compiler::primitives::errors::emit_try_start(&mut chunks[current], line);
    load(&mut chunks[current], base + 1, line);
    load(&mut chunks[current], error_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    save(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::errors::emit_try_end(&mut chunks[current], line);
    chunks[current].emit_br(0, line);
    vybe_compiler::primitives::errors::patch_catch(&mut chunks[current], handler_catch);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_string_const("error in error handling", line);
    save(&mut chunks[current], value_slot, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(handler_done);
    i32_const(&mut chunks[current], 0, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    save(&mut chunks[current], ok_slot, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(done);

    chunks[current].emit_else(line);
    chunks[current].emit_string_const("bad argument #2 to 'xpcall' (function expected)", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("attempt to call a non-function value", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    load(&mut chunks[current], ok_slot, line);
    chunks[current].emit_if_value(line);
    emit_object_get_const_key(
        &mut chunks[current],
        value_slot,
        vybe_compiler::primitives::multivalue::MULTI_VALUE_TAG,
        line,
    );
    save(&mut chunks[current], marker_slot, line);
    emit_lua_missing_to_nil(&mut chunks[current], marker_slot, line);
    vybe_compiler::primitives::ops::emit_lua_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_lua_pcall_prepend_ok_to_row(chunks, current, ok_slot, value_slot, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], ok_slot, line);
    load(&mut chunks[current], value_slot, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], ok_slot, line);
    load(&mut chunks[current], value_slot, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    chunks[current].emit_end(line);
    emit_lua_multi_row(chunks, current, 1, line);
}

pub fn emit_lua_assert(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("assertion failed!", line);
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        save(&mut chunks[current], base + i as u16, line);
    }

    load(&mut chunks[current], base, line);
    vybe_compiler::primitives::ops::emit_lua_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    if argc == 1 {
        load(&mut chunks[current], base, line);
    } else {
        for i in 0..argc {
            load(&mut chunks[current], base + i as u16, line);
        }
        chunks[current].emit_array_new_fixed(0, argc as u16, line);
        emit_lua_multi_row(chunks, current, 1, line);
    }
    chunks[current].emit_else(line);
    if argc >= 2 {
        load(&mut chunks[current], base + 1, line);
    } else {
        chunks[current].emit_string_const("assertion failed!", line);
    }
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_collectgarbage(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_f64_const(0.0, line);
}

pub fn emit_lua_error(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("error", line);
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        return;
    }
    if argc == 1 {
        let value = chunks[current].alloc_scratch(1);
        save(&mut chunks[current], value, line);
        emit_lua_first_if_multi_row(chunks, current, value, line);
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        return;
    }
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
}

pub fn emit_lua_multi_row(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    let value_slot = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], value_slot, line);
    load(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::multivalue::emit_tag(chunks, current, line);
}

pub fn emit_lua_multi_row_prefix(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }

    let row = chunks[current].alloc_scratch(1);
    let prefix = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);

    save(&mut chunks[current], row, line);
    save(&mut chunks[current], prefix, line);

    load(&mut chunks[current], row, line);
    vybe_compiler::primitives::multivalue::emit_as_multi_value(chunks, current, line);
    save(&mut chunks[current], row, line);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    save(&mut chunks[current], out, line);

    i32_const(&mut chunks[current], 0, line);
    save(&mut chunks[current], index, line);
    load(&mut chunks[current], prefix, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    save(&mut chunks[current], len, line);
    let prefix_loop = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    load(&mut chunks[current], index, line);
    load(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    load(&mut chunks[current], out, line);
    load(&mut chunks[current], prefix, line);
    load(&mut chunks[current], index, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    load(&mut chunks[current], index, line);
    i32_const(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    save(&mut chunks[current], index, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, prefix_loop, line);

    i32_const(&mut chunks[current], 0, line);
    save(&mut chunks[current], index, line);
    load(&mut chunks[current], row, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    save(&mut chunks[current], len, line);
    let row_loop = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    load(&mut chunks[current], index, line);
    load(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    load(&mut chunks[current], out, line);
    load(&mut chunks[current], row, line);
    load(&mut chunks[current], index, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    load(&mut chunks[current], index, line);
    i32_const(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    save(&mut chunks[current], index, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, row_loop, line);

    load(&mut chunks[current], out, line);
    vybe_compiler::primitives::multivalue::emit_tag(chunks, current, line);
}

pub fn emit_lua_as_multi_row(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }

    let value = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], value, line);
    load(&mut chunks[current], value, line);
    vybe_compiler::primitives::multivalue::emit_as_multi_value(chunks, current, line);
}

pub fn emit_lua_multi_index0(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }

    let index = chunks[current].alloc_scratch(1);
    let source = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], index, line);
    save(&mut chunks[current], source, line);
    emit_is_missing_value(&mut chunks[current], source, line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], source, line);
    load(&mut chunks[current], index, line);
    vybe_compiler::primitives::multivalue::emit_index0(chunks, current, line);
    let value = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], value, line);
    emit_lua_missing_to_nil(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_mark_rest(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let fixed = chunks[current].alloc_scratch(1);
    let func = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], fixed, line);
    save(&mut chunks[current], func, line);
    emit_lua_set_object_slot(
        &mut chunks[current],
        func,
        "__vybe_rest_fixed_arity",
        fixed,
        line,
    );
    load(&mut chunks[current], func, line);
}

pub fn emit_metamethod_len(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let value_slot = chunks[current].alloc_scratch(1);
    let mt_slot = chunks[current].alloc_scratch(1);
    let len_fn_slot = chunks[current].alloc_scratch(1);
    let str_test = chunks[current].add_import("wasm:js-string", "test");
    let arr_test = chunks[current].add_import("ecma:array", "isArray");
    let type_of = chunks[current].add_import("ecma:value", "typeof");
    let str_compare = chunks[current].add_import("wasm:js-string", "compare");
    let fn_call = chunks[current].add_import("ecma:function", "call");

    save(&mut chunks[current], value_slot, line);

    emit_lua_get_metatable_for_value(chunks, current, value_slot, line);
    save(&mut chunks[current], mt_slot, line);
    emit_is_missing_value(&mut chunks[current], mt_slot, line);
    chunks[current].emit_if(line);
    emit_is_missing_value(&mut chunks[current], value_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("attempt to get length of a nil value", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_else(line);

    load(&mut chunks[current], value_slot, line);
    call1(&mut chunks[current], str_test, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    chunks[current].emit_else(line);

    load(&mut chunks[current], value_slot, line);
    call1(&mut chunks[current], arr_test, line);
    chunks[current].emit_if(line);
    emit_lua_raw_sequence_len(chunks, current, value_slot, line);
    chunks[current].emit_else(line);

    load(&mut chunks[current], value_slot, line);
    chunks[current].emit_call(type_of, 1, line);
    chunks[current].emit_string_const("object", line);
    call2(&mut chunks[current], str_compare, line);
    i32_const(&mut chunks[current], 0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    emit_lua_raw_sequence_len(chunks, current, value_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("attempt to get length of a number value", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);

    load(&mut chunks[current], mt_slot, line);
    chunks[current].emit_string_const("__len", line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    save(&mut chunks[current], len_fn_slot, line);
    emit_is_missing_value(&mut chunks[current], len_fn_slot, line);
    chunks[current].emit_if(line);
    emit_lua_raw_sequence_len(chunks, current, value_slot, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], len_fn_slot, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    load(&mut chunks[current], value_slot, line);
    call3(&mut chunks[current], fn_call, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
}

pub fn emit_metamethod_concat(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_binary_metamethod_or_raw(chunks, current, "__concat", line, raw_concat);
}

pub fn emit_lua_float_repr(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(1);
    let is_integer = chunk.add_import("ecma:number", "isInteger");
    save(chunk, value, line);

    load(chunk, value, line);
    call1(chunk, is_integer, line);
    chunk.emit_if_value(line);
    load(chunk, value, line);
    vybe_compiler::primitives::strings::emit_to_string(chunk, line);
    chunk.emit_string_const(".0", line);
    vybe_compiler::primitives::strings::emit_str_concat(chunk, line);
    chunk.emit_else(line);
    load(chunk, value, line);
    vybe_compiler::primitives::strings::emit_to_string(chunk, line);
    chunk.emit_end(line);
}

pub fn emit_lua_first(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }

    let value_slot = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], value_slot, line);
    emit_is_missing_value(&mut chunks[current], value_slot, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    emit_lua_first_if_multi_row(chunks, current, value_slot, line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_tonumber(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }

    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(1);
    let base = chunk.alloc_scratch(1);
    let parsed = chunk.alloc_scratch(1);
    let test_num = chunk.add_import("wasm:js-number", "test");
    let test_str = chunk.add_import("wasm:js-string", "test");
    let parse_float = chunk.add_import("ecma:number", "parseFloat");
    let parse_int = chunk.add_import("ecma:number", "parseInt");
    let is_nan = chunk.add_import("ecma:number", "isNaN");

    if argc >= 2 {
        save(chunk, base, line);
    } else {
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        save(chunk, base, line);
    }
    save(chunk, value, line);

    if argc < 2 {
        load(chunk, value, line);
        call1(chunk, test_num, line);
        chunk.emit_if(line);
        load(chunk, value, line);
        chunk.emit_else(line);

        load(chunk, value, line);
        call1(chunk, test_str, line);
        chunk.emit_if(line);

        emit_regex_test_const(
            chunk,
            value,
            "/^\\s*[+-]?(?:0[xX][0-9a-fA-F]+|(?:[0-9]+(?:\\.[0-9]*)?|\\.[0-9]+)(?:[eE][+-]?[0-9]+)?)\\s*$/",
            line,
        );
        chunk.emit_if(line);
        emit_regex_test_const(chunk, value, "/^\\s*[+-]?0[xX][0-9a-fA-F]+\\s*$/", line);
        chunk.emit_if(line);
        load(chunk, value, line);
        chunk.emit_f64_const(16.0, line);
        call2(chunk, parse_int, line);
        chunk.emit_else(line);
        load(chunk, value, line);
        emit_regex_test_const(chunk, value, "/[eE]/", line);
        chunk.emit_if(line);
        load(chunk, value, line);
        call1(chunk, parse_float, line);
        chunk.emit_f64_const(0.0, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_else(line);
        load(chunk, value, line);
        call1(chunk, parse_float, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
        chunk.emit_else(line);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunk.emit_end(line);

        chunk.emit_else(line);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunk.emit_end(line);

        chunk.emit_end(line);
        return;
    }

    load(chunk, base, line);
    chunk.emit_f64_const(2.0, line);
    chunk.emit_op(Op::F64_LT, line);
    load(chunk, base, line);
    chunk.emit_f64_const(36.0, line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    chunk.emit_string_const("bad argument #2 to 'tonumber' (base out of range)", line);
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
    chunk.emit_else(line);

    load(chunk, value, line);
    call1(chunk, test_str, line);
    chunk.emit_if(line);
    emit_lua_valid_radix_string(chunk, value, base, line);
    chunk.emit_if(line);
    load(chunk, value, line);
    load(chunk, base, line);
    call2(chunk, parse_int, line);
    save(chunk, parsed, line);
    load(chunk, parsed, line);
    call1(chunk, is_nan, line);
    chunk.emit_if(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_else(line);
    load(chunk, parsed, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
}

pub fn emit_lua_rawlen(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let value = chunks[current].alloc_scratch(1);
    let str_test = chunks[current].add_import("wasm:js-string", "test");
    let arr_test = chunks[current].add_import("ecma:array", "isArray");
    let arr_len = chunks[current].add_import("ecma:array", "length");

    save(&mut chunks[current], value, line);
    load(&mut chunks[current], value, line);
    call1(&mut chunks[current], str_test, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], value, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], value, line);
    call1(&mut chunks[current], arr_test, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], value, line);
    chunks[current].emit_call(arr_len, 1, line);
    chunks[current].emit_else(line);
    i32_const(&mut chunks[current], 0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_rawget(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 2 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let key_slot = chunks[current].alloc_scratch(1);
    let table_slot = chunks[current].alloc_scratch(1);

    save(&mut chunks[current], key_slot, line);
    save(&mut chunks[current], table_slot, line);
    emit_lua_table_read(chunks, current, table_slot, key_slot, line);
}

pub fn emit_lua_rawset(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 3 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let value_slot = chunks[current].alloc_scratch(1);
    let key_slot = chunks[current].alloc_scratch(1);
    let table_slot = chunks[current].alloc_scratch(1);

    save(&mut chunks[current], value_slot, line);
    save(&mut chunks[current], key_slot, line);
    save(&mut chunks[current], table_slot, line);
    emit_lua_table_write(
        chunks, current, table_slot, key_slot, value_slot, true, line,
    );
}

pub fn emit_lua_table_insert(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    match argc {
        2 => {
            let value = chunks[current].alloc_scratch(1);
            let table = chunks[current].alloc_scratch(1);
            let pos = chunks[current].alloc_scratch(1);
            let arr_test = chunks[current].add_import("ecma:array", "isArray");
            save(&mut chunks[current], value, line);
            save(&mut chunks[current], table, line);
            load(&mut chunks[current], table, line);
            call1(&mut chunks[current], arr_test, line);
            chunks[current].emit_if(line);
            load(&mut chunks[current], table, line);
            vybe_compiler::primitives::collections::emit_len(chunks, current, line);
            chunks[current].emit_else(line);
            emit_lua_raw_sequence_len(chunks, current, table, line);
            chunks[current].emit_end(line);
            chunks[current].emit_op(Op::F64_FROM_I32, line);
            chunks[current].emit_f64_const(1.0, line);
            chunks[current].emit_op(Op::F64_ADD, line);
            save(&mut chunks[current], pos, line);
            emit_lua_table_write(chunks, current, table, pos, value, false, line);
        }
        3 => {
            let value = chunks[current].alloc_scratch(1);
            let pos = chunks[current].alloc_scratch(1);
            let table = chunks[current].alloc_scratch(1);
            let insert_at = chunks[current].add_import("ecma:array", "insertAt");
            let len_slot = chunks[current].alloc_scratch(1);
            save(&mut chunks[current], value, line);
            save(&mut chunks[current], pos, line);
            save(&mut chunks[current], table, line);
            load(&mut chunks[current], table, line);
            vybe_compiler::primitives::collections::emit_len(chunks, current, line);
            chunks[current].emit_op(Op::F64_FROM_I32, line);
            save(&mut chunks[current], len_slot, line);
            load(&mut chunks[current], pos, line);
            chunks[current].emit_f64_const(1.0, line);
            chunks[current].emit_op(Op::F64_LT, line);
            load(&mut chunks[current], pos, line);
            load(&mut chunks[current], len_slot, line);
            chunks[current].emit_f64_const(1.0, line);
            chunks[current].emit_op(Op::F64_ADD, line);
            chunks[current].emit_op(Op::F64_GT, line);
            chunks[current].emit_op(Op::I32_OR, line);
            chunks[current].emit_if(line);
            emit_lua_table_bounds_error(&mut chunks[current], line);
            chunks[current].emit_else(line);
            load(&mut chunks[current], table, line);
            load(&mut chunks[current], pos, line);
            chunks[current].emit_f64_const(1.0, line);
            chunks[current].emit_op(Op::F64_SUB, line);
            load(&mut chunks[current], value, line);
            chunks[current].emit_call(insert_at, 3, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_end(line);
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        }
        _ => chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line) }
}

pub fn emit_lua_table_remove(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    match argc {
        1 => {
            let table = chunks[current].alloc_scratch(1);
            let pos = chunks[current].alloc_scratch(1);
            let value = chunks[current].alloc_scratch(1);
            let nil_value = chunks[current].alloc_scratch(1);
            let remove_at = chunks[current].add_import("ecma:array", "removeAt");
            let arr_test = chunks[current].add_import("ecma:array", "isArray");
            save(&mut chunks[current], table, line);
            emit_lua_raw_sequence_len(chunks, current, table, line);
            save(&mut chunks[current], pos, line);
            load(&mut chunks[current], pos, line);
            i32_const(&mut chunks[current], 0, line);
            chunks[current].emit_op(Op::I32_EQ, line);
            chunks[current].emit_if(line);
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            chunks[current].emit_else(line);
            emit_lua_table_read(chunks, current, table, pos, line);
            save(&mut chunks[current], value, line);
            load(&mut chunks[current], table, line);
            call1(&mut chunks[current], arr_test, line);
            chunks[current].emit_if(line);
            load(&mut chunks[current], table, line);
            load(&mut chunks[current], pos, line);
            chunks[current].emit_f64_const(1.0, line);
            chunks[current].emit_op(Op::F64_SUB, line);
            chunks[current].emit_call(remove_at, 2, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_else(line);
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            save(&mut chunks[current], nil_value, line);
            emit_lua_table_write(chunks, current, table, pos, nil_value, false, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_end(line);
            load(&mut chunks[current], value, line);
            chunks[current].emit_end(line);
        }
        2 => {
            let pos = chunks[current].alloc_scratch(1);
            let table = chunks[current].alloc_scratch(1);
            let remove_at = chunks[current].add_import("ecma:array", "removeAt");
            let value = chunks[current].alloc_scratch(1);
            let len_slot = chunks[current].alloc_scratch(1);
            save(&mut chunks[current], pos, line);
            save(&mut chunks[current], table, line);
            load(&mut chunks[current], table, line);
            vybe_compiler::primitives::collections::emit_len(chunks, current, line);
            chunks[current].emit_op(Op::F64_FROM_I32, line);
            save(&mut chunks[current], len_slot, line);
            load(&mut chunks[current], pos, line);
            chunks[current].emit_f64_const(1.0, line);
            chunks[current].emit_op(Op::F64_LT, line);
            load(&mut chunks[current], pos, line);
            load(&mut chunks[current], len_slot, line);
            chunks[current].emit_op(Op::F64_GT, line);
            chunks[current].emit_op(Op::I32_OR, line);
            chunks[current].emit_if(line);
            emit_lua_table_bounds_error(&mut chunks[current], line);
            chunks[current].emit_else(line);
            load(&mut chunks[current], table, line);
            load(&mut chunks[current], pos, line);
            chunks[current].emit_f64_const(1.0, line);
            chunks[current].emit_op(Op::F64_SUB, line);
            chunks[current].emit_call(remove_at, 2, line);
            save(&mut chunks[current], value, line);
            emit_lua_missing_to_nil(&mut chunks[current], value, line);
            chunks[current].emit_end(line);
        }
        _ => chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line) }
}

fn emit_lua_table_concat_validate(
    chunks: &mut Vec<Chunk>,
    current: usize,
    table_slot: u16,
    line: u32,
) {
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let type_of = chunks[current].add_import("ecma:value", "typeof");
    let str_compare = chunks[current].add_import("wasm:js-string", "compare");

    i32_const(&mut chunks[current], 0, line);
    save(&mut chunks[current], i, line);
    load(&mut chunks[current], table_slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    save(&mut chunks[current], len, line);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    load(&mut chunks[current], i, line);
    load(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    load(&mut chunks[current], table_slot, line);
    load(&mut chunks[current], i, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    save(&mut chunks[current], value, line);
    emit_lua_type_is_slot(
        &mut chunks[current],
        value,
        type_of,
        str_compare,
        "string",
        line,
    );
    emit_lua_type_is_slot(
        &mut chunks[current],
        value,
        type_of,
        str_compare,
        "number",
        line,
    );
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("invalid value in table for concat", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    load(&mut chunks[current], i, line);
    i32_const(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    save(&mut chunks[current], i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
}

fn emit_lua_table_sequence_array(
    chunks: &mut Vec<Chunk>,
    current: usize,
    table_slot: u16,
    line: u32,
) {
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    save(&mut chunks[current], out, line);
    emit_lua_raw_sequence_len(chunks, current, table_slot, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    save(&mut chunks[current], len, line);
    chunks[current].emit_f64_const(1.0, line);
    save(&mut chunks[current], i, line);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    load(&mut chunks[current], i, line);
    load(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::F64_LE, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    emit_lua_table_read(chunks, current, table_slot, i, line);
    save(&mut chunks[current], value, line);
    load(&mut chunks[current], out, line);
    load(&mut chunks[current], value, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    load(&mut chunks[current], i, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    save(&mut chunks[current], i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    load(&mut chunks[current], out, line);
}

pub fn emit_lua_table_concat(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let end = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    let sep = chunks[current].alloc_scratch(1);
    let table = chunks[current].alloc_scratch(1);
    let sequence = chunks[current].alloc_scratch(1);

    if argc >= 4 {
        save(&mut chunks[current], end, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        save(&mut chunks[current], end, line);
    }
    if argc >= 3 {
        save(&mut chunks[current], start, line);
    } else {
        chunks[current].emit_f64_const(1.0, line);
        save(&mut chunks[current], start, line);
    }
    if argc >= 2 {
        save(&mut chunks[current], sep, line);
    } else {
        chunks[current].emit_string_const("", line);
        save(&mut chunks[current], sep, line);
    }
    save(&mut chunks[current], table, line);
    emit_lua_table_sequence_array(chunks, current, table, line);
    save(&mut chunks[current], sequence, line);

    if argc >= 3 {
        let slice = chunks[current].add_import("ecma:array", "slice");
        let sliced = chunks[current].alloc_scratch(1);
        let len_slot = chunks[current].alloc_scratch(1);
        load(&mut chunks[current], sequence, line);
        vybe_compiler::primitives::collections::emit_len(chunks, current, line);
        chunks[current].emit_op(Op::F64_FROM_I32, line);
        save(&mut chunks[current], len_slot, line);
        load(&mut chunks[current], start, line);
        chunks[current].emit_f64_const(1.0, line);
        chunks[current].emit_op(Op::F64_LT, line);
        load(&mut chunks[current], start, line);
        load(&mut chunks[current], end, line);
        chunks[current].emit_op(Op::F64_LE, line);
        load(&mut chunks[current], end, line);
        load(&mut chunks[current], len_slot, line);
        chunks[current].emit_op(Op::F64_GT, line);
        chunks[current].emit_op(Op::I32_AND, line);
        chunks[current].emit_op(Op::I32_OR, line);
        chunks[current].emit_if(line);
        emit_lua_table_bounds_error(&mut chunks[current], line);
        chunks[current].emit_else(line);
        load(&mut chunks[current], sequence, line);
        load(&mut chunks[current], start, line);
        chunks[current].emit_f64_const(1.0, line);
        chunks[current].emit_op(Op::F64_SUB, line);
        if argc >= 4 {
            load(&mut chunks[current], end, line);
        } else {
            load(&mut chunks[current], table, line);
            vybe_compiler::primitives::collections::emit_len(chunks, current, line);
        }
        chunks[current].emit_call(slice, 3, line);
        save(&mut chunks[current], sliced, line);
        emit_lua_table_concat_validate(chunks, current, sliced, line);
        load(&mut chunks[current], sliced, line);
        load(&mut chunks[current], sep, line);
        vybe_compiler::primitives::collections::emit_join(chunks, current, line);
        chunks[current].emit_end(line);
    } else {
        emit_lua_table_concat_validate(chunks, current, sequence, line);
        load(&mut chunks[current], sequence, line);
        load(&mut chunks[current], sep, line);
        vybe_compiler::primitives::collections::emit_join(chunks, current, line);
    }
}

pub fn emit_lua_table_sort(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 2 {
        vybe_compiler::primitives::collections::emit_sort_with_comparator(chunks, current, line);
    } else {
        vybe_compiler::primitives::collections::emit_sort(chunks, current, line);
    }
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_lua_table_pack(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        save(&mut chunks[current], base + i as u16, line);
    }
    let table = chunks[current].alloc_scratch(1);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    save(&mut chunks[current], table, line);
    for i in 0..argc {
        load(&mut chunks[current], table, line);
        load(&mut chunks[current], base + i as u16, line);
        vybe_compiler::primitives::collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    load(&mut chunks[current], table, line);
    chunks[current].emit_string_const("n", line);
    chunks[current].emit_f64_const(argc as f64, line);
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    load(&mut chunks[current], table, line);
}

pub fn emit_lua_table_pack_row(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let table = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], table, line);
    load(&mut chunks[current], table, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    save(&mut chunks[current], len, line);
    emit_lua_set_object_slot(&mut chunks[current], table, "n", len, line);
    load(&mut chunks[current], table, line);
}

pub fn emit_lua_table_unpack(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let end = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    let table = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);

    if argc >= 3 {
        save(&mut chunks[current], end, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        save(&mut chunks[current], end, line);
    }
    if argc >= 2 {
        save(&mut chunks[current], start, line);
    } else {
        chunks[current].emit_f64_const(1.0, line);
        save(&mut chunks[current], start, line);
    }
    save(&mut chunks[current], table, line);

    load(&mut chunks[current], end, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_lua_raw_sequence_len(chunks, current, table, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], end, line);
    chunks[current].emit_end(line);
    save(&mut chunks[current], end, line);

    load(&mut chunks[current], end, line);
    load(&mut chunks[current], start, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_f64_const(1000000.0, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("too many results to unpack", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    save(&mut chunks[current], out, line);
    load(&mut chunks[current], start, line);
    save(&mut chunks[current], i, line);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    load(&mut chunks[current], i, line);
    load(&mut chunks[current], end, line);
    chunks[current].emit_op(Op::F64_LE, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    emit_lua_table_read(chunks, current, table, i, line);
    save(&mut chunks[current], value, line);
    emit_lua_missing_to_nil(&mut chunks[current], value, line);
    save(&mut chunks[current], value, line);
    load(&mut chunks[current], out, line);
    load(&mut chunks[current], value, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    load(&mut chunks[current], i, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    save(&mut chunks[current], i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    load(&mut chunks[current], out, line);
    emit_lua_multi_row(chunks, current, 1, line);
}

pub fn emit_lua_table_move(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 4 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let dest = chunks[current].alloc_scratch(1);
    let target = chunks[current].alloc_scratch(1);
    let end = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    let source = chunks[current].alloc_scratch(1);
    let temp = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let dest_key = chunks[current].alloc_scratch(1);
    let type_of = chunks[current].add_import("ecma:value", "typeof");
    let str_compare = chunks[current].add_import("wasm:js-string", "compare");

    if argc >= 5 {
        save(&mut chunks[current], dest, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        save(&mut chunks[current], dest, line);
    }
    save(&mut chunks[current], target, line);
    save(&mut chunks[current], end, line);
    save(&mut chunks[current], start, line);
    save(&mut chunks[current], source, line);

    load(&mut chunks[current], dest, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], source, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], dest, line);
    chunks[current].emit_end(line);
    save(&mut chunks[current], dest, line);

    emit_lua_non_nil_object_is_slot(
        &mut chunks[current],
        source,
        type_of,
        str_compare,
        line,
    );
    chunks[current].emit_if(line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("bad argument #1 to 'move' (table expected)", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    emit_lua_non_nil_object_is_slot(
        &mut chunks[current],
        dest,
        type_of,
        str_compare,
        line,
    );
    chunks[current].emit_if(line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("bad argument #5 to 'move' (table expected)", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    save(&mut chunks[current], temp, line);
    load(&mut chunks[current], start, line);
    save(&mut chunks[current], i, line);

    let read_loop = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    load(&mut chunks[current], i, line);
    load(&mut chunks[current], end, line);
    chunks[current].emit_op(Op::F64_LE, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    emit_lua_table_read(chunks, current, source, i, line);
    save(&mut chunks[current], value, line);
    load(&mut chunks[current], temp, line);
    load(&mut chunks[current], value, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    load(&mut chunks[current], i, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    save(&mut chunks[current], i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, read_loop, line);

    chunks[current].emit_f64_const(0.0, line);
    save(&mut chunks[current], i, line);
    let write_loop = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    load(&mut chunks[current], i, line);
    load(&mut chunks[current], temp, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    chunks[current].emit_op(Op::F64_LT, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    load(&mut chunks[current], temp, line);
    load(&mut chunks[current], i, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    save(&mut chunks[current], value, line);
    load(&mut chunks[current], target, line);
    load(&mut chunks[current], i, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    save(&mut chunks[current], dest_key, line);
    emit_lua_table_write(chunks, current, dest, dest_key, value, false, line);
    chunks[current].emit_op(Op::DROP, line);
    load(&mut chunks[current], i, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    save(&mut chunks[current], i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, write_loop, line);
    load(&mut chunks[current], dest, line);
}

pub fn emit_lua_pairs(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    vybe_compiler::primitives::collections::emit_iter_entries(chunks, current, line);
}

pub fn emit_lua_iter_end(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_lua_ipairs(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    let table = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let row = chunks[current].alloc_scratch(1);

    save(&mut chunks[current], table, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    save(&mut chunks[current], out, line);
    chunks[current].emit_f64_const(1.0, line);
    save(&mut chunks[current], i, line);
    emit_lua_raw_sequence_len(chunks, current, table, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    save(&mut chunks[current], len, line);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    load(&mut chunks[current], i, line);
    load(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::F64_LE, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    emit_lua_table_read(chunks, current, table, i, line);
    save(&mut chunks[current], value, line);
    load(&mut chunks[current], value, line);
    emit_is_missing_value(&mut chunks[current], value, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    load(&mut chunks[current], i, line);
    load(&mut chunks[current], value, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    save(&mut chunks[current], row, line);
    load(&mut chunks[current], out, line);
    load(&mut chunks[current], row, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    load(&mut chunks[current], i, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    save(&mut chunks[current], i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    load(&mut chunks[current], out, line);
}

pub fn emit_lua_next(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let table = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let entries = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let row = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    let after_key = chunks[current].alloc_scratch(1);

    if argc >= 2 {
        save(&mut chunks[current], key, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        save(&mut chunks[current], key, line);
    }
    save(&mut chunks[current], table, line);

    load(&mut chunks[current], table, line);
    vybe_compiler::primitives::collections::emit_iter_entries(chunks, current, line);
    save(&mut chunks[current], entries, line);
    load(&mut chunks[current], entries, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    save(&mut chunks[current], len, line);

    i32_const(&mut chunks[current], 0, line);
    save(&mut chunks[current], i, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    save(&mut chunks[current], result, line);
    emit_is_missing_value(&mut chunks[current], key, line);
    save(&mut chunks[current], after_key, line);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    load(&mut chunks[current], i, line);
    load(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    load(&mut chunks[current], result, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_AND, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    load(&mut chunks[current], entries, line);
    load(&mut chunks[current], i, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    save(&mut chunks[current], row, line);

    load(&mut chunks[current], after_key, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], row, line);
    save(&mut chunks[current], result, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], row, line);
    i32_const(&mut chunks[current], 0, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    load(&mut chunks[current], key, line);
    emit_lua_key_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(true, line);
    save(&mut chunks[current], after_key, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    load(&mut chunks[current], i, line);
    i32_const(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    save(&mut chunks[current], i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    load(&mut chunks[current], result, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], after_key, line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    emit_lua_multi_row(chunks, current, 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("invalid key to next", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], result, line);
    emit_lua_multi_row(chunks, current, 1, line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_tostring(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let tostring_fn = chunks[current].alloc_scratch(1);
    let test_num = chunks[current].add_import("wasm:js-number", "test");
    let test_i32 = chunks[current].add_import("wasm:js-number", "testI32");
    let fn_call = chunks[current].add_import("ecma:function", "call");

    save(&mut chunks[current], value, line);
    load(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("nil", line);
    chunks[current].emit_else(line);

    emit_lua_get_metamethod(chunks, current, value, "__tostring", line);
    save(&mut chunks[current], tostring_fn, line);
    emit_is_missing_value(&mut chunks[current], tostring_fn, line);
    chunks[current].emit_if_value(line);

    load(&mut chunks[current], value, line);
    call1(&mut chunks[current], test_num, line);
    chunks[current].emit_if(line);

    load(&mut chunks[current], value, line);
    call1(&mut chunks[current], test_i32, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], value, line);
    vybe_compiler::primitives::strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], value, line);
    vybe_compiler::primitives::strings::emit_to_string(&mut chunks[current], line);

    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], value, line);
    vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("function:", line);
    load(&mut chunks[current], value, line);
    vybe_compiler::primitives::strings::emit_to_string(&mut chunks[current], line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], value, line);
    vybe_compiler::primitives::strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    load(&mut chunks[current], tostring_fn, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    load(&mut chunks[current], value, line);
    call3(&mut chunks[current], fn_call, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
}

pub fn emit_lua_type(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        chunks[current].emit_string_const("nil", line);
        return;
    }

    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(1);
    let type_slot = chunk.alloc_scratch(1);
    let tag_slot = chunk.alloc_scratch(1);
    let type_of = chunk.add_import("ecma:value", "typeof");
    let str_compare = chunk.add_import("wasm:js-string", "compare");

    save(chunk, value, line);
    emit_is_missing_value(chunk, value, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("nil", line);
    chunk.emit_else(line);

    load(chunk, value, line);
    call1(chunk, type_of, line);
    save(chunk, type_slot, line);
    load(chunk, type_slot, line);
    chunk.emit_string_const("object", line);
    call2(chunk, str_compare, line);
    i32_const(chunk, 0, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    emit_object_get_const_key(chunk, value, "__lua_type", line);
    save(chunk, tag_slot, line);
    emit_is_missing_value(chunk, tag_slot, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("table", line);
    chunk.emit_else(line);
    load(chunk, tag_slot, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    load(chunk, type_slot, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
}

pub fn emit_lua_print(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        save(&mut chunks[current], base + i as u16, line);
    }
    if argc == 0 {
        let log_idx = chunks[current].add_import("web:console", "log");
        vybe_compiler::primitives::io::emit_print_with_import(&mut chunks[current], log_idx, 0, line);
        return;
    }
    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_string_const("", line);
    save(&mut chunks[current], out, line);
    for i in 0..argc {
        if i > 0 {
            load(&mut chunks[current], out, line);
            chunks[current].emit_string_const("\t", line);
            vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
            save(&mut chunks[current], out, line);
        }
        load(&mut chunks[current], out, line);
        load(&mut chunks[current], base + i as u16, line);
        emit_lua_tostring(chunks, current, 1, line);
        vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
        save(&mut chunks[current], out, line);
    }
    load(&mut chunks[current], out, line);
    let log_idx = chunks[current].add_import("web:console", "log");
    vybe_compiler::primitives::io::emit_print_with_import(&mut chunks[current], log_idx, 1, line);
}

pub fn emit_lua_print_row(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 1 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_string_const("", line);
        let log_idx = chunks[current].add_import("web:console", "log");
        vybe_compiler::primitives::io::emit_print_with_import(&mut chunks[current], log_idx, 1, line);
        return;
    }

    let row = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);

    save(&mut chunks[current], row, line);
    i32_const(&mut chunks[current], 0, line);
    save(&mut chunks[current], index, line);
    chunks[current].emit_string_const("", line);
    save(&mut chunks[current], out, line);
    load(&mut chunks[current], row, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    save(&mut chunks[current], len, line);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    load(&mut chunks[current], index, line);
    load(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    load(&mut chunks[current], index, line);
    i32_const(&mut chunks[current], 0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], out, line);
    chunks[current].emit_string_const("\t", line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    save(&mut chunks[current], out, line);
    chunks[current].emit_end(line);

    load(&mut chunks[current], row, line);
    load(&mut chunks[current], index, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    save(&mut chunks[current], value, line);
    load(&mut chunks[current], out, line);
    load(&mut chunks[current], value, line);
    emit_lua_tostring(chunks, current, 1, line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    save(&mut chunks[current], out, line);

    load(&mut chunks[current], index, line);
    i32_const(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    save(&mut chunks[current], index, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    load(&mut chunks[current], out, line);
    let log_idx = chunks[current].add_import("web:console", "log");
    vybe_compiler::primitives::io::emit_print_with_import(&mut chunks[current], log_idx, 1, line);
}

pub fn emit_lua_apply_row(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }

    let row = chunks[current].alloc_scratch(1);
    let callee = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], row, line);
    save(&mut chunks[current], callee, line);

    load(&mut chunks[current], callee, line);
    vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], callee, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    load(&mut chunks[current], row, line);
    vybe_compiler::primitives::reflection::emit_reflect_op(
        chunks,
        current,
        vybe_compiler::primitives::reflection::ReflectOp::Apply,
        3,
        line,
    );
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("attempt to call a non-function value", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_truthy(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    vybe_compiler::primitives::ops::emit_lua_to_bool(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_metamethod_call(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("attempt to call a non-function value", line);
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc as u16);
    let method = chunks[current].alloc_scratch(1);
    let nested_method = chunks[current].alloc_scratch(1);
    let rest_arity = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    for i in (0..argc).rev() {
        save(&mut chunks[current], base + i as u16, line);
    }
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    save(&mut chunks[current], result, line);

    load(&mut chunks[current], base, line);
    vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
    chunks[current].emit_if(line);
    emit_object_get_const_key(&mut chunks[current], base, "__vybe_rest_fixed_arity", line);
    save(&mut chunks[current], rest_arity, line);
    emit_is_missing_value(&mut chunks[current], rest_arity, line);
    chunks[current].emit_if(line);
    emit_lua_call_fixed(chunks, current, base, argc, line);
    save(&mut chunks[current], result, line);
    chunks[current].emit_else(line);
    emit_lua_call_rest_dispatch(chunks, current, base, argc, rest_arity, line);
    save(&mut chunks[current], result, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);

    emit_lua_get_metamethod(chunks, current, base, "__call", line);
    save(&mut chunks[current], method, line);
    emit_is_missing_value(&mut chunks[current], method, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("attempt to call a non-function value", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], method, line);
    vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], method, line);
    load(&mut chunks[current], base, line);
    for i in 1..argc {
        load(&mut chunks[current], base + i as u16, line);
    }
    chunks[current].emit_op_u8_u8(Op::CALL_REF, argc, 1, line);
    save(&mut chunks[current], result, line);
    chunks[current].emit_else(line);
    emit_lua_get_metamethod(chunks, current, method, "__call", line);
    save(&mut chunks[current], nested_method, line);
    load(&mut chunks[current], nested_method, line);
    vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], nested_method, line);
    load(&mut chunks[current], base, line);
    for i in 1..argc {
        load(&mut chunks[current], base + i as u16, line);
    }
    chunks[current].emit_op_u8_u8(Op::CALL_REF, argc, 1, line);
    save(&mut chunks[current], result, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("attempt to call a non-function value", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    load(&mut chunks[current], result, line);
}

pub fn emit_lua_method_call(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    vybe_compiler::primitives::invoke::emit_protocol_method_call(
        chunks,
        current,
        argc,
        line,
        |chunks, current, receiver, method_key, line| {
            load(&mut chunks[current], receiver, line);
            load(&mut chunks[current], method_key, line);
            emit_metamethod_index(chunks, current, 2, line);
        },
        emit_metamethod_call,
    );
}

pub fn emit_lua_apply_row_prefix(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 3 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }

    let row = chunks[current].alloc_scratch(1);
    let prefix = chunks[current].alloc_scratch(1);
    let callee = chunks[current].alloc_scratch(1);
    let args = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);

    save(&mut chunks[current], row, line);
    save(&mut chunks[current], prefix, line);
    save(&mut chunks[current], callee, line);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    save(&mut chunks[current], args, line);

    i32_const(&mut chunks[current], 0, line);
    save(&mut chunks[current], index, line);
    load(&mut chunks[current], prefix, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    save(&mut chunks[current], len, line);
    let prefix_loop = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    load(&mut chunks[current], index, line);
    load(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    load(&mut chunks[current], args, line);
    load(&mut chunks[current], prefix, line);
    load(&mut chunks[current], index, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    load(&mut chunks[current], index, line);
    i32_const(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    save(&mut chunks[current], index, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, prefix_loop, line);

    i32_const(&mut chunks[current], 0, line);
    save(&mut chunks[current], index, line);
    load(&mut chunks[current], row, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    save(&mut chunks[current], len, line);
    let row_loop = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    load(&mut chunks[current], index, line);
    load(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    load(&mut chunks[current], args, line);
    load(&mut chunks[current], row, line);
    load(&mut chunks[current], index, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    load(&mut chunks[current], index, line);
    i32_const(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    save(&mut chunks[current], index, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, row_loop, line);

    load(&mut chunks[current], callee, line);
    vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], callee, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    load(&mut chunks[current], args, line);
    vybe_compiler::primitives::reflection::emit_reflect_op(
        chunks,
        current,
        vybe_compiler::primitives::reflection::ReflectOp::Apply,
        3,
        line,
    );
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("attempt to call a non-function value", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

fn emit_lua_call_fixed(chunks: &mut Vec<Chunk>, current: usize, base: u16, argc: u8, line: u32) {
    load(&mut chunks[current], base, line);
    for i in 1..argc {
        load(&mut chunks[current], base + i as u16, line);
    }
    chunks[current].emit_op_u8_u8(Op::CALL_REF, argc - 1, 1, line);
}

fn emit_lua_call_rest_dispatch(
    chunks: &mut Vec<Chunk>,
    current: usize,
    base: u16,
    argc: u8,
    rest_arity: u16,
    line: u32,
) {
    emit_num_eq_const(&mut chunks[current], rest_arity, 0.0, line);
    chunks[current].emit_if(line);
    emit_lua_call_rest_fixed(chunks, current, base, argc, 0, line);
    chunks[current].emit_else(line);
    emit_num_eq_const(&mut chunks[current], rest_arity, 1.0, line);
    chunks[current].emit_if(line);
    emit_lua_call_rest_fixed(chunks, current, base, argc, 1, line);
    chunks[current].emit_else(line);
    emit_num_eq_const(&mut chunks[current], rest_arity, 2.0, line);
    chunks[current].emit_if(line);
    emit_lua_call_rest_fixed(chunks, current, base, argc, 2, line);
    chunks[current].emit_else(line);
    emit_lua_call_fixed(chunks, current, base, argc, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_lua_call_rest_fixed(
    chunks: &mut Vec<Chunk>,
    current: usize,
    base: u16,
    argc: u8,
    fixed: u8,
    line: u32,
) {
    let rest = chunks[current].alloc_scratch(1);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    save(&mut chunks[current], rest, line);

    let actual = argc.saturating_sub(1);
    for i in fixed..actual {
        load(&mut chunks[current], rest, line);
        load(&mut chunks[current], base + 1 + i as u16, line);
        vybe_compiler::primitives::collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    load(&mut chunks[current], rest, line);
    emit_lua_multi_row(chunks, current, 1, line);
    save(&mut chunks[current], rest, line);

    load(&mut chunks[current], base, line);
    for i in 0..fixed {
        if i < actual {
            load(&mut chunks[current], base + 1 + i as u16, line);
        } else {
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        }
    }
    load(&mut chunks[current], rest, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, fixed + 1, 1, line);
}
