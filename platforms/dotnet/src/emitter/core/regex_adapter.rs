use std::sync::Arc;
use vybe_compiler::compiler::instructions::core_wasm;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

const PATTERN_KEY: &str = "__pattern";
const TIMEOUT_KEY: &str = "__timeout";
const VALUE_KEY: &str = "value";
const SUCCESS_KEY: &str = "success";
const COUNT_KEY: &str = "count";
const GROUPS_KEY: &str = "__groups";
const GROUP_VALUES_KEY: &str = "__group_values";
const RAW_GROUPS_KEY: &str = "groups";
const KEYS_KEY: &str = "__keys";
const GROUP_NAMES_KEY: &str = "__group_names";
const GROUP_NUMBERS_KEY: &str = "__group_numbers";

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

fn emit_regex_pattern_arg(
    chunk: &mut Chunk,
    pattern_slot: u16,
    options_slot: Option<u16>,
    line: u32,
) {
    let Some(options_slot) = options_slot else {
        chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
        return;
    };
    let concat = chunk.add_import("wasm:js-string", "concat");
    let flags_slot = reserve_slot(chunk);

    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::LOCAL_SET, flags_slot, line);
    for (bit, flag) in [(1, "i"), (2, "m"), (16, "s")] {
        chunk.emit_op_u16(Op::LOCAL_GET, options_slot, line);
        chunk.emit_i32_const(bit, line);
        chunk.emit_op(Op::I32_AND, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, flags_slot, line);
        chunk.emit_string_const(flag, line);
        chunk.emit_op_u16(Op::CALL_IMPORT, concat, line);
        chunk.emit(2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, flags_slot, line);
        chunk.emit_end(line);
    }

    chunk.emit_string_const("/", line);
    chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, concat, line);
    chunk.emit(2, line);
    chunk.emit_string_const("/", line);
    chunk.emit_op_u16(Op::CALL_IMPORT, concat, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, flags_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, concat, line);
    chunk.emit(2, line);
}

fn emit_dotnet_match_collection_shape(
    chunk: &mut Chunk,
    result_slot: u16,
    count_slot: u16,
    line: u32,
) {
    let count_key = chunk.add_constant(Value::String(Arc::from(COUNT_KEY)));
    let value_key = chunk.add_constant(Value::String(Arc::from(VALUE_KEY)));
    let success_key = chunk.add_constant(Value::String(Arc::from(SUCCESS_KEY)));
    let i_slot = reserve_slot(chunk);
    let match_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, count_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let block_p = chunk.emit_block(line);
    let (loop_p, _) = chunk.emit_loop_s(line);

    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::compiler::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, match_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, match_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, match_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::STRUCT_SET, value_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, match_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_bool_const(true, line);
    chunk.emit_op_u16(Op::STRUCT_SET, success_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_i32_const(1, line);
    vybe_compiler::compiler::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunk.emit_br(0, line);
    chunk.patch_loop(loop_p);
    chunk.emit_end(line);
    chunk.patch_block(block_p);
    chunk.emit_end(line);
}

fn emit_group_object_from_value_slot(chunk: &mut Chunk, value_slot: u16, line: u32) {
    let value_key = chunk.add_constant(Value::String(Arc::from(VALUE_KEY)));
    let dotnet_value_key = chunk.add_constant(Value::String(Arc::from("Value")));
    let success_key = chunk.add_constant(Value::String(Arc::from(SUCCESS_KEY)));
    let dotnet_success_key = chunk.add_constant(Value::String(Arc::from("Success")));
    let length_key = chunk.add_constant(Value::String(Arc::from("Length")));
    let lower_length_key = chunk.add_constant(Value::String(Arc::from("length")));
    let group_slot = reserve_slot(chunk);
    let length_idx = chunk.add_import("wasm:js-string", "length");

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, group_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, value_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, dotnet_value_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    vybe_compiler::compiler::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_SET, success_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    vybe_compiler::compiler::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_SET, dotnet_success_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, length_idx, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::STRUCT_SET, length_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, length_idx, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::STRUCT_SET, lower_length_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_slot, line);
}

fn emit_dotnet_match_properties(chunk: &mut Chunk, result_slot: u16, obj_slot: u16, line: u32) {
    let object_get = chunk.add_import("ecma:object", "get");
    let length_idx = chunk.add_import("wasm:js-string", "length");
    let index_key = chunk.add_constant(Value::String(Arc::from("index")));
    let dotnet_index_key = chunk.add_constant(Value::String(Arc::from("Index")));
    let length_key = chunk.add_constant(Value::String(Arc::from("Length")));
    let lower_length_key = chunk.add_constant(Value::String(Arc::from("length")));
    let dotnet_success_key = chunk.add_constant(Value::String(Arc::from("Success")));
    let value_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::CONST, index_key, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, object_get, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::STRUCT_SET, dotnet_index_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_string_const("", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, length_idx, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::STRUCT_SET, length_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, length_idx, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::STRUCT_SET, lower_length_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    vybe_compiler::compiler::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_SET, dotnet_success_key, line);
    chunk.emit_op(Op::DROP, line);
}

fn emit_dotnet_match_groups_shape(chunk: &mut Chunk, result_slot: u16, obj_slot: u16, line: u32) {
    let object_get = chunk.add_import("ecma:object", "get");
    let object_set = chunk.add_import("ecma:object", "set");
    let groups_key = chunk.add_constant(Value::String(Arc::from(GROUPS_KEY)));
    let group_values_key = chunk.add_constant(Value::String(Arc::from(GROUP_VALUES_KEY)));
    let dotnet_groups_key = chunk.add_constant(Value::String(Arc::from("Groups")));
    let public_groups_key = chunk.add_constant(Value::String(Arc::from("groups")));
    let count_key = chunk.add_constant(Value::String(Arc::from(COUNT_KEY)));
    let dotnet_count_key = chunk.add_constant(Value::String(Arc::from("Count")));
    let raw_groups_key = chunk.add_constant(Value::String(Arc::from(RAW_GROUPS_KEY)));
    let keys_key = chunk.add_constant(Value::String(Arc::from(KEYS_KEY)));
    let groups_slot = reserve_slot(chunk);
    let group_values_slot = reserve_slot(chunk);
    let raw_groups_slot = reserve_slot(chunk);
    let keys_slot = reserve_slot(chunk);
    let count_slot = reserve_slot(chunk);
    let i_slot = reserve_slot(chunk);
    let value_slot = reserve_slot(chunk);
    let key_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::ARRAY_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, groups_slot, line);
    chunk.emit_op_u16(Op::ARRAY_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, group_values_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, groups_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, count_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, groups_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, dotnet_count_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_values_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, count_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_values_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, dotnet_count_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let block_p = chunk.emit_block(line);
    let (loop_p, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::compiler::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, groups_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    emit_group_object_from_value_slot(chunk, value_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_values_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_i32_const(1, line);
    vybe_compiler::compiler::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunk.emit_br(0, line);
    chunk.patch_loop(loop_p);
    chunk.emit_end(line);
    chunk.patch_block(block_p);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::CONST, raw_groups_key, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, object_get, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, raw_groups_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, raw_groups_slot, line);
    chunk.emit_op_u16(Op::CONST, keys_key, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, object_get, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let block_p = chunk.emit_block(line);
    let (loop_p, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::compiler::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, raw_groups_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, object_get, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, groups_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
    emit_group_object_from_value_slot(chunk, value_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, object_set, line);
    chunk.emit(3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_values_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, object_set, line);
    chunk.emit(3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_i32_const(1, line);
    vybe_compiler::compiler::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunk.emit_br(0, line);
    chunk.patch_loop(loop_p);
    chunk.emit_end(line);
    chunk.patch_block(block_p);
    chunk.emit_end(line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, groups_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, groups_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, groups_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, dotnet_groups_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, groups_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, public_groups_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_values_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, group_values_key, line);
    chunk.emit_op(Op::DROP, line);
}

pub fn emit_regex_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let pattern_key = chunk.add_constant(Value::String(Arc::from(PATTERN_KEY)));
    match argc {
        0 => {
            chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
            core_wasm::dup(chunk, line);
            push_const(chunk, Value::String(Arc::from("")), line);
            chunk.emit_op_u16(Op::STRUCT_SET, pattern_key, line);
            chunk.emit_op(Op::DROP, line);
        }
        _ => {
            for _ in 1..argc {
                chunk.emit_op(Op::DROP, line);
            }
            let pattern_slot = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, pattern_slot, line);
            let timeout_key = chunk.add_constant(Value::String(Arc::from(TIMEOUT_KEY)));
            chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
            core_wasm::dup(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
            chunk.emit_op_u16(Op::STRUCT_SET, pattern_key, line);
            chunk.emit_op(Op::DROP, line);
            if argc >= 3 {
                core_wasm::dup(chunk, line);
                chunk.emit_bool_const(true, line);
                chunk.emit_op_u16(Op::STRUCT_SET, timeout_key, line);
                chunk.emit_op(Op::DROP, line);
            }
        }
    }
}

pub fn emit_regex_is_match(chunks: &mut [Chunk], current: usize, line: u32) {
    let test_idx = chunks[current].add_import("ecma:regexp", "test");
    let chunk = &mut chunks[current];
    let input_slot = reserve_slot(chunk);
    let self_slot = reserve_slot(chunk);
    let pattern_key = chunk.add_constant(Value::String(Arc::from(PATTERN_KEY)));
    let timeout_key = chunk.add_constant(Value::String(Arc::from(TIMEOUT_KEY)));

    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, timeout_key, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_string_const("The regex operation timed out.", line);
    vybe_compiler::compiler::errors::emit_exception_new_finalize(chunk, "RegexMatchTimeoutException", line);
    vybe_compiler::compiler::errors::emit_throw(chunk, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, pattern_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, test_idx, line);
    chunk.emit(2, line);
}

pub fn emit_regex_static_is_match(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let test_idx = chunks[current].add_import("ecma:regexp", "test");
    let chunk = &mut chunks[current];
    let options_slot = (argc >= 3).then(|| reserve_slot(chunk));
    let pattern_slot = reserve_slot(chunk);
    let input_slot = reserve_slot(chunk);

    for _ in 3..argc {
        chunk.emit_op(Op::DROP, line);
    }
    if let Some(slot) = options_slot {
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, pattern_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    emit_regex_pattern_arg(chunk, pattern_slot, options_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, test_idx, line);
    chunk.emit(2, line);
}

pub fn emit_regex_static_replace(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let replace_idx = chunks[current].add_import("ecma:regexp", "replaceAll");
    let chunk = &mut chunks[current];
    let options_slot = (argc >= 4).then(|| reserve_slot(chunk));
    let replacement_slot = reserve_slot(chunk);
    let pattern_slot = reserve_slot(chunk);
    let input_slot = reserve_slot(chunk);

    for _ in 4..argc {
        chunk.emit_op(Op::DROP, line);
    }
    if let Some(slot) = options_slot {
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, replacement_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, pattern_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    emit_regex_pattern_arg(chunk, pattern_slot, options_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, replacement_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, replace_idx, line);
    chunk.emit(3, line);
}

pub fn emit_regex_static_split(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let split_idx = chunks[current].add_import("ecma:regexp", "split");
    let chunk = &mut chunks[current];
    let options_slot = (argc >= 3).then(|| reserve_slot(chunk));
    let pattern_slot = reserve_slot(chunk);
    let input_slot = reserve_slot(chunk);

    for _ in 3..argc {
        chunk.emit_op(Op::DROP, line);
    }
    if let Some(slot) = options_slot {
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, pattern_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    emit_regex_pattern_arg(chunk, pattern_slot, options_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, split_idx, line);
    chunk.emit(2, line);
}

pub fn emit_regex_escape(chunks: &mut [Chunk], current: usize, line: u32) {
    let escape_idx = chunks[current].add_import("ecma:regexp", "escape");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, escape_idx, line);
    chunk.emit(1, line);
}

pub fn emit_regex_unescape(chunks: &mut [Chunk], current: usize, line: u32) {
    let replace_idx = chunks[current].add_import("ecma:regexp", "replaceAll");
    let chunk = &mut chunks[current];
    let input_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_string_const(r"/\\([\\.\+\*\?\^\$\[\]\(\)\{\}\=\!\<\>\|\:\-])/g", line);
    chunk.emit_string_const("$1", line);
    chunk.emit_op_u16(Op::CALL_IMPORT, replace_idx, line);
    chunk.emit(3, line);
}

pub fn emit_regex_replace(chunks: &mut [Chunk], current: usize, line: u32) {
    let replace_idx = chunks[current].add_import("ecma:regexp", "replaceAll");
    let chunk = &mut chunks[current];
    let replacement_slot = reserve_slot(chunk);
    let input_slot = reserve_slot(chunk);
    let self_slot = reserve_slot(chunk);
    let pattern_key = chunk.add_constant(Value::String(Arc::from(PATTERN_KEY)));

    chunk.emit_op_u16(Op::LOCAL_SET, replacement_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, pattern_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, replacement_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, replace_idx, line);
    chunk.emit(3, line);
}

pub fn emit_regex_split(chunks: &mut [Chunk], current: usize, line: u32) {
    let split_idx = chunks[current].add_import("ecma:regexp", "split");
    let chunk = &mut chunks[current];
    let input_slot = reserve_slot(chunk);
    let self_slot = reserve_slot(chunk);
    let pattern_key = chunk.add_constant(Value::String(Arc::from(PATTERN_KEY)));

    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, pattern_key, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, split_idx, line);
    chunk.emit(2, line);
}

pub fn emit_regex_match(chunks: &mut [Chunk], current: usize, line: u32) {
    let exec_idx = chunks[current].add_import("ecma:regexp", "exec");
    let chunk = &mut chunks[current];
    let input_slot = reserve_slot(chunk);
    let self_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);
    let obj_slot = reserve_slot(chunk);
    let pattern_key = chunk.add_constant(Value::String(Arc::from(PATTERN_KEY)));
    let value_key = chunk.add_constant(Value::String(Arc::from(VALUE_KEY)));
    let success_key = chunk.add_constant(Value::String(Arc::from(SUCCESS_KEY)));

    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, pattern_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, exec_idx, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::STRUCT_SET, value_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    vybe_compiler::compiler::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_SET, success_key, line);
    chunk.emit_op(Op::DROP, line);

    emit_dotnet_match_properties(chunk, result_slot, obj_slot, line);
    emit_dotnet_match_groups_shape(chunk, result_slot, obj_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

pub fn emit_regex_static_match(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let exec_idx = chunks[current].add_import("ecma:regexp", "exec");
    let chunk = &mut chunks[current];
    let options_slot = (argc >= 3).then(|| reserve_slot(chunk));
    let pattern_slot = reserve_slot(chunk);
    let input_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);
    let obj_slot = reserve_slot(chunk);
    let value_key = chunk.add_constant(Value::String(Arc::from(VALUE_KEY)));
    let success_key = chunk.add_constant(Value::String(Arc::from(SUCCESS_KEY)));

    for _ in 3..argc {
        chunk.emit_op(Op::DROP, line);
    }
    if let Some(slot) = options_slot {
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, pattern_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);

    emit_regex_pattern_arg(chunk, pattern_slot, options_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, exec_idx, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::STRUCT_SET, value_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    vybe_compiler::compiler::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_SET, success_key, line);
    chunk.emit_op(Op::DROP, line);

    emit_dotnet_match_properties(chunk, result_slot, obj_slot, line);
    emit_dotnet_match_groups_shape(chunk, result_slot, obj_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

pub fn emit_regex_matches(chunks: &mut [Chunk], current: usize, line: u32) {
    let match_all_idx = chunks[current].add_import("ecma:regexp", "matchAll");
    let chunk = &mut chunks[current];
    let input_slot = reserve_slot(chunk);
    let self_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);
    let count_slot = reserve_slot(chunk);
    let pattern_key = chunk.add_constant(Value::String(Arc::from(PATTERN_KEY)));

    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, pattern_key, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, match_all_idx, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    emit_dotnet_match_collection_shape(chunk, result_slot, count_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_regex_static_matches(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let match_all_idx = chunks[current].add_import("ecma:regexp", "matchAll");
    let chunk = &mut chunks[current];
    let options_slot = (argc >= 3).then(|| reserve_slot(chunk));
    let pattern_slot = reserve_slot(chunk);
    let input_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);
    let count_slot = reserve_slot(chunk);

    for _ in 3..argc {
        chunk.emit_op(Op::DROP, line);
    }
    if let Some(slot) = options_slot {
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, pattern_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    emit_regex_pattern_arg(chunk, pattern_slot, options_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, match_all_idx, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    emit_dotnet_match_collection_shape(chunk, result_slot, count_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_regex_get_group_names(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let group_names_key = chunk.add_constant(Value::String(Arc::from(GROUP_NAMES_KEY)));

    chunk.emit_op_u16(Op::STRUCT_GET, group_names_key, line);
}

pub fn emit_regex_group_name_from_number(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let number_slot = reserve_slot(chunk);
    let self_slot = reserve_slot(chunk);
    let group_names_key = chunk.add_constant(Value::String(Arc::from(GROUP_NAMES_KEY)));

    chunk.emit_op_u16(Op::LOCAL_SET, number_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, group_names_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, number_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

pub fn emit_regex_group_number_from_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let object_get = chunks[current].add_import("ecma:object", "get");
    let chunk = &mut chunks[current];
    let name_slot = reserve_slot(chunk);
    let self_slot = reserve_slot(chunk);
    let group_numbers_key = chunk.add_constant(Value::String(Arc::from(GROUP_NUMBERS_KEY)));

    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, group_numbers_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, object_get, line);
    chunk.emit(2, line);
}
