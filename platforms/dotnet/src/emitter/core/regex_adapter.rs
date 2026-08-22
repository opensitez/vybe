use std::sync::Arc;
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::loops;

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

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
// `GROUP_NUMBERS_KEY` ("__group_numbers") stood here. It was READ by
// `GroupNumberFromName` and written by nothing, ever — a second dead key beside
// `__group_names`. A name's index in `__group_names` IS its group number, so
// there is one array with one writer and no second structure to keep in sync.

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
        chunk.emit_call(concat, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, flags_slot, line);
        chunk.emit_end(line);
    }

    chunk.emit_string_const("/", line);
    chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
    chunk.emit_call(concat, 2, line);
    chunk.emit_string_const("/", line);
    chunk.emit_call(concat, 2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, flags_slot, line);
    chunk.emit_call(concat, 2, line);
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
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, count_key, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let block_p = chunk.emit_block(line);
    let (loop_p, _) = chunk.emit_loop_s(line);

    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
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
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, value_key, line);

    chunk.emit_op_u16(Op::LOCAL_GET, match_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_bool_const(true, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, success_key, line);

    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
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

    chunk.emit_struct_new(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, group_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, value_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, dotnet_value_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, success_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, dotnet_success_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(length_idx, 1, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, length_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(length_idx, 1, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, lower_length_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_slot, line);
}

fn emit_dotnet_match_properties(chunk: &mut Chunk, result_slot: u16, obj_slot: u16, line: u32) {
    let object_get = chunk.add_import("ecma:object", "get");
    let length_idx = chunk.add_import("wasm:js-string", "length");
    let dotnet_index_key = chunk.add_constant(Value::String(Arc::from("Index")));
    let length_key = chunk.add_constant(Value::String(Arc::from("Length")));
    let lower_length_key = chunk.add_constant(Value::String(Arc::from("length")));
    let dotnet_success_key = chunk.add_constant(Value::String(Arc::from("Success")));
    let value_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_string_const("index", line);
    chunk.emit_call(object_get, 2, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, dotnet_index_key, line);

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
    chunk.emit_call(length_idx, 1, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, length_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(length_idx, 1, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, lower_length_key, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, dotnet_success_key, line);
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
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, count_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, groups_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, dotnet_count_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_values_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, count_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_values_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, dotnet_count_key, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let block_p = chunk.emit_block(line);
    let (loop_p, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, groups_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    emit_group_object_from_value_slot(chunk, value_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_values_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunk.emit_br(0, line);
    chunk.patch_loop(loop_p);
    chunk.emit_end(line);
    chunk.patch_block(block_p);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_string_const(RAW_GROUPS_KEY, line);
    chunk.emit_call(object_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, raw_groups_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, raw_groups_slot, line);
    chunk.emit_string_const(KEYS_KEY, line);
    chunk.emit_call(object_get, 2, line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, raw_groups_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunk.emit_call(object_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, groups_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
    emit_group_object_from_value_slot(chunk, value_slot, line);
    chunk.emit_call(object_set, 3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_values_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(object_set, 3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
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
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, groups_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, groups_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, dotnet_groups_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, groups_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, public_groups_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, group_values_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, group_values_key, line);
}

/// Populate `__group_names` on a freshly built Regex, from its own pattern.
///
/// ⛔ `emit_regex_get_group_names` / `..._group_name_from_number` /
/// `..._group_number_from_name` all READ this key and NOTHING EVER WROTE IT, so
/// all three answered empty. They looked implemented — declared on the component
/// class, dispatched, with a real emitter each — which is exactly why the gap
/// survived: the VB walker carried its own compile-time copy that worked for
/// literal patterns, and that copy was the only thing making the tests pass.
///
/// .NET's ordering is group 0, then the UNNAMED capturing groups by number, then
/// the NAMED ones in declaration order — `(\w+)(?<n>\d)` gives `0,1,n`.
///
/// Built as a delimited string and `split` at the end rather than by appending
/// to an array: `ARRAY_SET` pushes nothing, so a growing array needs an index
/// dance that string concatenation avoids entirely.
///
/// ⚠ Known limit: an ESCAPED `\(` counts as a capturing group here, so a pattern
/// that matches a literal paren over-reports the unnamed count. Fixing that needs
/// a real pattern scan rather than two regex passes.
///
/// Loop scaffolding comes from `primitives::loops` rather than being spelled out
/// here. Hand-rolling it once already cost a HANG: `emit_loop_end` emits TWO
/// `END`s — one for the loop, one for the enclosing block — and dropping the
/// second leaves the block open, so the `br_if` meant to exit targets the wrong
/// depth. The primitive cannot get that wrong.
fn emit_store_group_names(
    chunks: &mut [Chunk],
    current: usize,
    obj_slot: u16,
    pattern_slot: u16,
    line: u32,
) {
    let match_all_idx = chunks[current].add_import("ecma:regexp", "matchAll");
    let split_idx = chunks[current].add_import("ecma:regexp", "split");
    let chunk = &mut chunks[current];
    let group_names_key = chunk.add_constant(Value::String(Arc::from(GROUP_NAMES_KEY)));

    let named_slot = reserve_slot(chunk);
    let unnamed_slot = reserve_slot(chunk);
    let out_slot = reserve_slot(chunk);
    let i_slot = reserve_slot(chunk);
    let n_slot = reserve_slot(chunk);

    // named = matchAll(pattern, "\(\?<([A-Za-z_][A-Za-z0-9_]*)>")
    chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
    push_const(
        chunk,
        Value::String(Arc::from(r"\(\?<([A-Za-z_][A-Za-z0-9_]*)>")),
        line,
    );
    chunk.emit_call(match_all_idx, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, named_slot, line);

    // unnamed = matchAll(pattern, "\((?!\?)")
    chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
    push_const(chunk, Value::String(Arc::from(r"\((?!\?)")), line);
    chunk.emit_call(match_all_idx, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, unnamed_slot, line);

    push_const(chunk, Value::String(Arc::from("0")), line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);

    // `,1`, `,2`, … for each unnamed capturing group.
    emit_count_into(chunk, unnamed_slot, n_slot, i_slot, line);
    let state = loops::emit_loop_start(chunks, current, line);
    emit_index_below_count(&mut chunks[current], i_slot, n_slot, line);
    loops::emit_loop_cond(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        emit_append_comma(chunk, out_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
        emit_increment(chunk, i_slot, line);
    }
    loops::emit_loop_end(chunks, current, state, line);

    // `,<name>` for each named group — capture 1 of each match.
    emit_count_into(&mut chunks[current], named_slot, n_slot, i_slot, line);
    let state = loops::emit_loop_start(chunks, current, line);
    emit_index_below_count(&mut chunks[current], i_slot, n_slot, line);
    loops::emit_loop_cond(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        emit_append_comma(chunk, out_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, named_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::ARRAY_GET, line);
        vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
        emit_increment(chunk, i_slot, line);
    }
    loops::emit_loop_end(chunks, current, state, line);

    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
    push_const(chunk, Value::String(Arc::from(",")), line);
    chunk.emit_call(split_idx, 2, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, group_names_key, line);
}

/// `n = array.length; i = 0` — the counted-loop preamble.
fn emit_count_into(chunk: &mut Chunk, array_slot: u16, n_slot: u16, i_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op_u16(Op::LOCAL_SET, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
}

/// Leaves `i < n` on the stack for [`loops::emit_loop_cond`], which applies the
/// ToBoolean/negate itself — so this must NOT pre-negate.
fn emit_index_below_count(chunk: &mut Chunk, i_slot: u16, n_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
}

/// Leaves `out & ","` on the stack, ready for the piece being appended.
fn emit_append_comma(chunk: &mut Chunk, out_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
    push_const(chunk, Value::String(Arc::from(",")), line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
}

fn emit_increment(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    push_const(chunk, Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

pub fn emit_regex_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let pattern_key = chunk.add_constant(Value::String(Arc::from(PATTERN_KEY)));
    match argc {
        0 => {
            chunk.emit_struct_new(0, 0, line);
            core_wasm::dup(chunk, line);
            push_const(chunk, Value::String(Arc::from("")), line);
            chunk.emit_struct_field_op(Op::STRUCT_SET, 0, pattern_key, line);
        }
        _ => {
            for _ in 1..argc {
                chunk.emit_op(Op::DROP, line);
            }
            let pattern_slot = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, pattern_slot, line);
            let timeout_key = chunk.add_constant(Value::String(Arc::from(TIMEOUT_KEY)));
            chunk.emit_struct_new(0, 0, line);
            core_wasm::dup(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
            chunk.emit_struct_field_op(Op::STRUCT_SET, 0, pattern_key, line);
            if argc >= 3 {
                core_wasm::dup(chunk, line);
                chunk.emit_bool_const(true, line);
                chunk.emit_struct_field_op(Op::STRUCT_SET, 0, timeout_key, line);
            }
            // The object is still the only thing on the stack here; park it in a
            // slot so the group-name scan can write to it, then hand it back.
            let obj_slot = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
            emit_store_group_names(chunks, current, obj_slot, pattern_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
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
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, timeout_key, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_struct_new(0, 0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_string_const("The regex operation timed out.", line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        chunk,
        "RegexMatchTimeoutException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, pattern_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_call(test_idx, 2, line);
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
    chunk.emit_call(test_idx, 2, line);
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
    chunk.emit_call(replace_idx, 3, line);
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
    chunk.emit_call(split_idx, 2, line);
}

pub fn emit_regex_escape(chunks: &mut [Chunk], current: usize, line: u32) {
    let escape_idx = chunks[current].add_import("ecma:regexp", "escape");
    let chunk = &mut chunks[current];
    chunk.emit_call(escape_idx, 1, line);
}

pub fn emit_regex_unescape(chunks: &mut [Chunk], current: usize, line: u32) {
    let replace_idx = chunks[current].add_import("ecma:regexp", "replaceAll");
    let chunk = &mut chunks[current];
    let input_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_string_const(r"/\\([\\.\+\*\?\^\$\[\]\(\)\{\}\=\!\<\>\|\:\-])/g", line);
    chunk.emit_string_const("$1", line);
    chunk.emit_call(replace_idx, 3, line);
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
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, pattern_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, replacement_slot, line);
    chunk.emit_call(replace_idx, 3, line);
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
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, pattern_key, line);
    chunk.emit_call(split_idx, 2, line);
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
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, pattern_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_call(exec_idx, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunk.emit_struct_new(0, 0, line);
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
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, value_key, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, success_key, line);

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
    chunk.emit_call(exec_idx, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunk.emit_struct_new(0, 0, line);
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
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, value_key, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, success_key, line);

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
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, pattern_key, line);
    chunk.emit_call(match_all_idx, 2, line);
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
    chunk.emit_call(match_all_idx, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    emit_dotnet_match_collection_shape(chunk, result_slot, count_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_regex_get_group_names(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let group_names_key = chunk.add_constant(Value::String(Arc::from(GROUP_NAMES_KEY)));

    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, group_names_key, line);
}

pub fn emit_regex_group_name_from_number(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let number_slot = reserve_slot(chunk);
    let self_slot = reserve_slot(chunk);
    let group_names_key = chunk.add_constant(Value::String(Arc::from(GROUP_NAMES_KEY)));

    chunk.emit_op_u16(Op::LOCAL_SET, number_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, group_names_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, number_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

/// `regex.GroupNumberFromName(name)` — the number of a named group, or -1.
///
/// ⛔ This used to read a `__group_numbers` key which, like `__group_names`
/// before it, **nothing ever wrote** — a second dead key on the same object, so
/// the call answered `undefined` for every name.
///
/// No separate map is needed: `__group_names` is already built in .NET's own
/// order — `0`, then the unnamed groups by number, then the named ones — so a
/// name's INDEX in that array IS its group number. One array, one writer, and
/// `GroupNameFromNumber` is its exact inverse by construction.
pub fn emit_regex_group_number_from_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let name_slot = reserve_slot(chunk);
    let self_slot = reserve_slot(chunk);
    let names_slot = reserve_slot(chunk);
    let i_slot = reserve_slot(chunk);
    let n_slot = reserve_slot(chunk);
    let res_slot = reserve_slot(chunk);
    let group_names_key = chunk.add_constant(Value::String(Arc::from(GROUP_NAMES_KEY)));

    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, group_names_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, names_slot, line);

    push_const(chunk, Value::F64(-1.0), line);
    chunk.emit_op_u16(Op::LOCAL_SET, res_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, names_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op_u16(Op::LOCAL_SET, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);

    // Scans the whole array rather than breaking on the first hit: group names
    // are unique, so the answer is the same, and it keeps every branch inside
    // the `if` — nothing has to `br` out through two levels to get wrong.
    let state = loops::emit_loop_start(chunks, current, line);
    emit_index_below_count(&mut chunks[current], i_slot, n_slot, line);
    loops::emit_loop_cond(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, names_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, res_slot, line);
        chunk.emit_end(line);
        emit_increment(chunk, i_slot, line);
    }
    loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, res_slot, line);
}
