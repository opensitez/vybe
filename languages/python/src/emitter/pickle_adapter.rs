//! Python `pickle`/`marshal` and module-level `copy` adapters.
//!
//! The runtime we can honestly support here is the one the tests exercise:
//! same-process round trips. `dumps` stores the live value and returns a bytes
//! object containing a small JSON envelope. `loads` decodes the envelope and
//! reads the value back. The envelope rendering goes through the shared JSON
//! primitive; Python-specific dunder hooks stay at this adapter boundary.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::{Op, heaptype::HT_EXTERN};

use super::adapter_util::{call_import, lget, lset, new_object, stash_exact};
use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ValueSource,
};
use vybe_compiler::primitives::{collections, dict, loops, ops, sets};

const STORE_KEY: &str = "__vybe_py_pickle_store";
const FILE_STORE_KEY: &str = "__vybe_py_pickle_file_store";
const STREAM_STORE_KEY: &str = "__vybe_py_pickle_stream_store";
const PICKLE_ID: &str = "1";
const COPY_FIELDS_KEY: &str = "__py_deepcopy_fields";

pub fn emit_pickle_dumps(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 2, line);
    let value = base;
    let stored = chunks[current].alloc_scratch(1);
    let memo = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], value, line);
    lset(&mut chunks[current], stored, line);

    lget(&mut chunks[current], value, line);
    vybe_compiler::primitives::instructions::recipes::is_object(&mut chunks[current], line);
    chunks[current].emit_if(line);
    apply_pickle_protocols(chunks, current, value, stored, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], stored, line);
    vybe_compiler::primitives::instructions::recipes::is_object(&mut chunks[current], line);
    chunks[current].emit_if(line);
    dict::emit_new(chunks, current, line);
    lset(&mut chunks[current], memo, line);
    emit_python_deepcopy_slot(chunks, current, stored, memo, line);
    lset(&mut chunks[current], stored, line);
    chunks[current].emit_end(line);

    store_value(chunks, current, stored, line);
    emit_envelope_bytes(chunks, current, line);
}

pub fn emit_pickle_loads(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    let encoded = base;
    let text = chunks[current].alloc_scratch(1);
    let envelope = chunks[current].alloc_scratch(1);

    decode_bytes_slot(chunks, current, encoded, line);
    lset(&mut chunks[current], text, line);

    lget(&mut chunks[current], text, line);
    call_import(chunks, current, "ecma:json", "parse", 1, line);
    lset(&mut chunks[current], envelope, line);

    load_value(chunks, current, envelope, line);
}

pub fn emit_pickle_dump(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 3, line);
    let value = base;
    let file = base + 1;
    let path = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], value, line);
    emit_pickle_dumps(chunks, current, 1, line);
    let bytes = chunks[current].alloc_scratch(1);
    let text = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], bytes, line);

    decode_bytes_slot(chunks, current, bytes, line);
    lset(&mut chunks[current], text, line);

    emit_python_file_path(chunks, current, file, path, line);
    emit_slot_is_null_or_undefined(chunks, current, path, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    store_current_pickle_for_path(chunks, current, path, line);
    lget(&mut chunks[current], path, line);
    lget(&mut chunks[current], text, line);
    vybe_compiler::primitives::fs_path::emit_append_file(&mut chunks[current], line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    store_current_pickle_for_stream(chunks, current, file, line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
}

pub fn emit_pickle_load(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    let file = base;
    let path = chunks[current].alloc_scratch(1);
    let stored = chunks[current].alloc_scratch(1);

    emit_python_file_path(chunks, current, file, path, line);
    emit_path_slot_is_present(chunks, current, path, line);
    chunks[current].emit_if_value(line);
    load_file_value(chunks, current, path, line);
    chunks[current].emit_else(line);
    load_stream_value(chunks, current, file, line);
    chunks[current].emit_end(line);
    lset(&mut chunks[current], stored, line);

    emit_slot_is_null_or_undefined(chunks, current, stored, line);
    chunks[current].emit_if_value(line);
    emit_pickle_load_via_file_contents(chunks, current, file, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], stored, line);
    chunks[current].emit_end(line);
}

fn emit_pickle_load_via_file_contents(
    chunks: &mut Vec<Chunk>,
    current: usize,
    file: u16,
    line: u32,
) {
    let text = chunks[current].alloc_scratch(1);
    let envelope = chunks[current].alloc_scratch(1);

    emit_is_python_file_obj(chunks, current, file, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], file, line);
    crate::emitter::file_adapter::emit_read(chunks, current, 1, line);
    chunks[current].emit_else(line);
    emit_stream_method_call(chunks, current, file, "read", None, line);
    chunks[current].emit_end(line);
    lset(&mut chunks[current], text, line);

    lget(&mut chunks[current], text, line);
    vybe_compiler::primitives::instructions::recipes::is_object(&mut chunks[current], line);
    chunks[current].emit_if(line);
    decode_bytes_slot(chunks, current, text, line);
    lset(&mut chunks[current], text, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], text, line);
    call_import(chunks, current, "ecma:json", "parse", 1, line);
    lset(&mut chunks[current], envelope, line);

    load_value(chunks, current, envelope, line);
}

pub fn emit_pickle_reduce(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 16, line);
    let out = chunks[current].alloc_scratch(1);
    let args = chunks[current].alloc_scratch(1);

    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], args, line);
    for offset in 1..argc as u16 {
        lget(&mut chunks[current], args, line);
        lget(&mut chunks[current], base + offset, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }

    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], out, line);
    lget(&mut chunks[current], out, line);
    lget(&mut chunks[current], base, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], out, line);
    lget(&mut chunks[current], args, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], out, line);
}

pub fn emit_pickle_pickler(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 2, line);
    let file = base;
    let obj = chunks[current].alloc_scratch(1);

    new_object(&mut chunks[current], line);
    lset(&mut chunks[current], obj, line);
    lget(&mut chunks[current], obj, line);
    chunks[current].emit_string_const("file", line);
    lget(&mut chunks[current], file, line);
    call_import(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], obj, line);
}

pub fn emit_copy_copy(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    let value = base;
    let method = chunks[current].alloc_scratch(1);
    let reduced = chunks[current].alloc_scratch(1);
    let copied = chunks[current].alloc_scratch(1);

    if_not_object_return_original(&mut chunks[current], value, line);
    lookup_method(chunks, current, value, "__copy__", method, line);
    emit_slot_is_null_or_undefined(chunks, current, method, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    let recv =
        vybe_compiler::primitives::callable::push_callback_from_slot(chunks, current, method, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(&mut chunks[current], recv, line);
    chunks[current].emit_else(line);

    lookup_method(chunks, current, value, "__reduce__", method, line);
    emit_slot_is_null_or_undefined(chunks, current, method, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    let recv =
        vybe_compiler::primitives::callable::push_callback_from_slot(chunks, current, method, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(&mut chunks[current], recv, line);
    lset(&mut chunks[current], reduced, line);
    invoke_reduce_tuple(chunks, current, reduced, copied, line);
    lget(&mut chunks[current], copied, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], value, line);
    crate::emitter::collections_adapter::emit_copy(chunks, current, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
}

pub fn emit_copy_deepcopy(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 2, line);
    let value = base;
    let memo = base + 1;
    let method = chunks[current].alloc_scratch(1);

    if_not_object_return_original(&mut chunks[current], value, line);
    lget(&mut chunks[current], memo, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    dict::emit_new(chunks, current, line);
    lset(&mut chunks[current], memo, line);
    chunks[current].emit_end(line);

    lookup_method(chunks, current, value, "__deepcopy__", method, line);
    emit_slot_is_null_or_undefined(chunks, current, method, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    let recv =
        vybe_compiler::primitives::callable::push_callback_from_slot(chunks, current, method, line);
    lget(&mut chunks[current], memo, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(
        &mut chunks[current],
        1 + recv,
        line,
    );
    chunks[current].emit_else(line);
    emit_python_deepcopy_slot(chunks, current, value, memo, line);
    chunks[current].emit_end(line);
}

pub fn emit_copy_deepcopy_fields(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 3, line);
    let value = base;
    let memo = base + 1;
    let fields = base + 2;
    let out = chunks[current].alloc_scratch(1);
    let i_slot = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let item = chunks[current].alloc_scratch(1);
    let copied = chunks[current].alloc_scratch(1);

    if_not_object_return_original(&mut chunks[current], value, line);
    lget(&mut chunks[current], memo, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    dict::emit_new(chunks, current, line);
    lset(&mut chunks[current], memo, line);
    chunks[current].emit_end(line);

    emit_python_deepcopy_slot(chunks, current, value, memo, line);
    lset(&mut chunks[current], out, line);
    deepcopy_named_fields(
        chunks, current, value, out, memo, fields, i_slot, key, item, copied, line,
    );
    lget(&mut chunks[current], out, line);
}

pub fn emit_copy_stamp_fields(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 2, line);
    let value = base;
    let fields = base + 1;

    lget(&mut chunks[current], value, line);
    lget(&mut chunks[current], fields, line);
    set_class_slot_from_stack(chunks, current, &ClassSlot::internal(COPY_FIELDS_KEY), line);
    lget(&mut chunks[current], value, line);
}

fn emit_python_deepcopy_slot(
    chunks: &mut Vec<Chunk>,
    current: usize,
    value: u16,
    memo: u16,
    line: u32,
) {
    let helper = ensure_python_deepcopy_chunk(chunks);
    chunks[current].emit_op_u16(Op::REF_FUNC, helper as u16, line);
    chunks[current].emit(0u8, line);
    lget(&mut chunks[current], value, line);
    lget(&mut chunks[current], memo, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(&mut chunks[current], 2, line);
}

fn ensure_python_deepcopy_chunk(chunks: &mut Vec<Chunk>) -> usize {
    const NAME: &str = "__vybe_py_deepcopy";
    if let Some(idx) = chunks.iter().position(|c| c.name == NAME) {
        return idx;
    }
    let mut chunk = Chunk::new(NAME);
    chunk.arity = 2;
    chunk.local_count = 2;
    chunks.push(chunk);
    let idx = chunks.len() - 1;
    emit_python_deepcopy_body(chunks, idx);
    idx
}

fn emit_python_deepcopy_body(chunks: &mut Vec<Chunk>, current: usize) {
    const LINE: u32 = 0;
    let value = 0u16;
    let memo = 1u16;
    let base = chunks[current].alloc_scratch(6);
    let out = base;
    let keys = base + 1;
    let i_slot = base + 2;
    let key = base + 3;
    let item = base + 4;
    let copied = base + 5;

    if_not_object_return_original(&mut chunks[current], value, LINE);

    emit_to_string_tag_equals(chunks, current, value, "[object Uint8Array]", LINE);
    chunks[current].emit_if(LINE);
    lget(&mut chunks[current], value, LINE);
    chunks[current].emit_op(Op::RETURN, LINE);
    chunks[current].emit_end(LINE);

    emit_is_map_slot(chunks, current, value, LINE);
    chunks[current].emit_if_value(LINE);
    call_import(chunks, current, "ecma:map", "new", 0, LINE);
    lset(&mut chunks[current], out, LINE);
    lget(&mut chunks[current], value, LINE);
    call_import(chunks, current, "ecma:map", "keys", 1, LINE);
    call_import(chunks, current, "ecma:array", "from", 1, LINE);
    lset(&mut chunks[current], keys, LINE);
    let map_loop = loops::emit_for_in_start(chunks, current, keys, i_slot, LINE);
    lset(&mut chunks[current], key, LINE);
    lget(&mut chunks[current], value, LINE);
    lget(&mut chunks[current], key, LINE);
    call_import(chunks, current, "ecma:map", "get", 2, LINE);
    lset(&mut chunks[current], item, LINE);
    emit_python_deepcopy_slot(chunks, current, item, memo, LINE);
    lset(&mut chunks[current], copied, LINE);
    lget(&mut chunks[current], out, LINE);
    lget(&mut chunks[current], key, LINE);
    lget(&mut chunks[current], copied, LINE);
    call_import(chunks, current, "ecma:map", "set", 3, LINE);
    chunks[current].emit_op(Op::DROP, LINE);
    loops::emit_for_in_end(chunks, current, i_slot, map_loop, LINE);
    lget(&mut chunks[current], out, LINE);
    chunks[current].emit_op(Op::RETURN, LINE);
    chunks[current].emit_end(LINE);

    emit_to_string_tag_equals(chunks, current, value, "[object Set]", LINE);
    chunks[current].emit_if(LINE);
    lget(&mut chunks[current], value, LINE);
    call_import(chunks, current, "ecma:array", "from", 1, LINE);
    sets::emit_from_iterable(chunks, current, LINE);
    lset(&mut chunks[current], out, LINE);
    lget(&mut chunks[current], value, LINE);
    chunks[current].emit_string_const("__frozenset", LINE);
    call_import(chunks, current, "ecma:object", "get", 2, LINE);
    lset(&mut chunks[current], item, LINE);
    lget(&mut chunks[current], item, LINE);
    chunks[current].emit_op(Op::REF_IS_NULL, LINE);
    chunks[current].emit_op(Op::I32_EQZ, LINE);
    chunks[current].emit_if(LINE);
    lget(&mut chunks[current], out, LINE);
    chunks[current].emit_string_const("__frozenset", LINE);
    lget(&mut chunks[current], item, LINE);
    call_import(chunks, current, "ecma:object", "set", 3, LINE);
    chunks[current].emit_op(Op::DROP, LINE);
    chunks[current].emit_end(LINE);
    lget(&mut chunks[current], out, LINE);
    chunks[current].emit_op(Op::RETURN, LINE);
    chunks[current].emit_end(LINE);

    lget(&mut chunks[current], value, LINE);
    call_import(chunks, current, "ecma:array", "isArray", 1, LINE);
    ops::emit_dyn_to_bool(&mut chunks[current], LINE);
    chunks[current].emit_if(LINE);
    lget(&mut chunks[current], value, LINE);
    chunks[current].emit_i32_const(0, LINE);
    lget(&mut chunks[current], value, LINE);
    collections::emit_len(chunks, current, LINE);
    collections::emit_slice(chunks, current, LINE);
    lset(&mut chunks[current], out, LINE);
    deepcopy_array_elements(chunks, current, out, memo, i_slot, item, copied, LINE);
    lget(&mut chunks[current], out, LINE);
    chunks[current].emit_op(Op::RETURN, LINE);
    chunks[current].emit_end(LINE);

    class_slots::emit_class_alloc(&mut chunks[current], LINE);
    lget(&mut chunks[current], value, LINE);
    call_import(chunks, current, "ecma:object", "assign", 2, LINE);
    lset(&mut chunks[current], out, LINE);
    preserve_object_prototype(chunks, current, value, out, LINE);
    deepcopy_object_entries(
        chunks, current, value, out, memo, keys, i_slot, key, item, copied, LINE,
    );
    preserve_python_class_metadata(chunks, current, value, out, LINE);
    emit_class_slot_get(
        chunks,
        current,
        value,
        &ClassSlot::internal(COPY_FIELDS_KEY),
        LINE,
    );
    lset(&mut chunks[current], keys, LINE);
    emit_slot_is_null_or_undefined(chunks, current, keys, LINE);
    chunks[current].emit_op(Op::I32_EQZ, LINE);
    chunks[current].emit_if(LINE);
    deepcopy_named_fields(
        chunks, current, value, out, memo, keys, i_slot, key, item, copied, LINE,
    );
    chunks[current].emit_end(LINE);
    lget(&mut chunks[current], out, LINE);
    chunks[current].emit_op(Op::RETURN, LINE);
}

fn deepcopy_array_elements(
    chunks: &mut Vec<Chunk>,
    current: usize,
    out: u16,
    memo: u16,
    i_slot: u16,
    item: u16,
    copied: u16,
    line: u32,
) {
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], out, line);
    collections::emit_len(chunks, current, line);
    lset(&mut chunks[current], n, line);

    let array_loop = loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], out, line);
    lget(&mut chunks[current], i_slot, line);
    collections::emit_get(chunks, current, line);
    lset(&mut chunks[current], item, line);
    emit_python_deepcopy_slot(chunks, current, item, memo, line);
    lset(&mut chunks[current], copied, line);
    lget(&mut chunks[current], out, line);
    lget(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], copied, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i_slot, line);
    loops::emit_loop_end(chunks, current, array_loop, line);
}

#[allow(clippy::too_many_arguments)]
fn deepcopy_object_entries(
    chunks: &mut Vec<Chunk>,
    current: usize,
    source: u16,
    out: u16,
    memo: u16,
    keys: u16,
    i_slot: u16,
    key: u16,
    item: u16,
    copied: u16,
    line: u32,
) {
    lget(&mut chunks[current], source, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    lset(&mut chunks[current], keys, line);
    let object_loop = loops::emit_for_in_start(chunks, current, keys, i_slot, line);
    lset(&mut chunks[current], key, line);
    lget(&mut chunks[current], source, line);
    lget(&mut chunks[current], key, line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    lset(&mut chunks[current], item, line);
    emit_python_deepcopy_slot(chunks, current, item, memo, line);
    lset(&mut chunks[current], copied, line);
    lget(&mut chunks[current], out, line);
    lget(&mut chunks[current], key, line);
    lget(&mut chunks[current], copied, line);
    call_import(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, i_slot, object_loop, line);
}

#[allow(clippy::too_many_arguments)]
fn deepcopy_named_fields(
    chunks: &mut Vec<Chunk>,
    current: usize,
    source: u16,
    out: u16,
    memo: u16,
    fields: u16,
    i_slot: u16,
    key: u16,
    item: u16,
    copied: u16,
    line: u32,
) {
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], fields, line);
    collections::emit_len(chunks, current, line);
    lset(&mut chunks[current], n, line);

    let field_loop = loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], fields, line);
    lget(&mut chunks[current], i_slot, line);
    collections::emit_get(chunks, current, line);
    lset(&mut chunks[current], key, line);
    lget(&mut chunks[current], source, line);
    lget(&mut chunks[current], key, line);
    collections::emit_get(chunks, current, line);
    lset(&mut chunks[current], item, line);
    emit_python_deepcopy_slot(chunks, current, item, memo, line);
    lset(&mut chunks[current], copied, line);
    lget(&mut chunks[current], out, line);
    lget(&mut chunks[current], key, line);
    lget(&mut chunks[current], copied, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i_slot, line);
    loops::emit_loop_end(chunks, current, field_loop, line);
}

fn emit_is_map_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    emit_to_string_tag_equals(chunks, current, slot, "[object Map]", line);
}

fn emit_to_string_tag_equals(
    chunks: &mut [Chunk],
    current: usize,
    slot: u16,
    expected: &str,
    line: u32,
) {
    lget(&mut chunks[current], slot, line);
    call_import(chunks, current, "ecma:object", "toStringTag", 1, line);
    chunks[current].emit_string_const(expected, line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
}

fn apply_pickle_protocols(
    chunks: &mut Vec<Chunk>,
    current: usize,
    value: u16,
    stored: u16,
    line: u32,
) {
    let method = chunks[current].alloc_scratch(1);
    let reduced = chunks[current].alloc_scratch(1);
    let state = chunks[current].alloc_scratch(1);

    lookup_method(chunks, current, value, "__reduce__", method, line);
    emit_slot_is_null_or_undefined(chunks, current, method, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    let recv =
        vybe_compiler::primitives::callable::push_callback_from_slot(chunks, current, method, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(&mut chunks[current], recv, line);
    lset(&mut chunks[current], reduced, line);
    invoke_reduce_tuple(chunks, current, reduced, stored, line);
    chunks[current].emit_else(line);

    lookup_method(chunks, current, value, "__getstate__", method, line);
    emit_slot_is_null_or_undefined(chunks, current, method, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    let recv =
        vybe_compiler::primitives::callable::push_callback_from_slot(chunks, current, method, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(&mut chunks[current], recv, line);
    lset(&mut chunks[current], state, line);

    lookup_method(chunks, current, value, "__setstate__", method, line);
    emit_slot_is_null_or_undefined(chunks, current, method, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    let recv =
        vybe_compiler::primitives::callable::push_callback_from_slot(chunks, current, method, line);
    lget(&mut chunks[current], state, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(
        &mut chunks[current],
        1 + recv,
        line,
    );
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn invoke_reduce_tuple(
    chunks: &mut Vec<Chunk>,
    current: usize,
    reduced: u16,
    stored: u16,
    line: u32,
) {
    let ctor = chunks[current].alloc_scratch(1);
    let args = chunks[current].alloc_scratch(1);
    let arg0 = chunks[current].alloc_scratch(1);
    let arg1 = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], reduced, line);
    chunks[current].emit_i32_const(0, line);
    call_import(chunks, current, "ecma:array", "get", 2, line);
    lset(&mut chunks[current], ctor, line);
    resolve_global_callable_name(chunks, current, ctor, line);

    lget(&mut chunks[current], reduced, line);
    chunks[current].emit_i32_const(1, line);
    call_import(chunks, current, "ecma:array", "get", 2, line);
    lset(&mut chunks[current], args, line);

    lget(&mut chunks[current], args, line);
    chunks[current].emit_i32_const(0, line);
    call_import(chunks, current, "ecma:array", "get", 2, line);
    lset(&mut chunks[current], arg0, line);
    lget(&mut chunks[current], args, line);
    chunks[current].emit_i32_const(1, line);
    call_import(chunks, current, "ecma:array", "get", 2, line);
    lset(&mut chunks[current], arg1, line);

    let recv =
        vybe_compiler::primitives::callable::push_callback_from_slot(chunks, current, ctor, line);
    lget(&mut chunks[current], arg0, line);
    lget(&mut chunks[current], arg1, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(
        &mut chunks[current],
        2 + recv,
        line,
    );
    lset(&mut chunks[current], stored, line);
}

fn resolve_global_callable_name(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    let global = chunks[current].alloc_scratch(1);
    let resolved = chunks[current].alloc_scratch(1);

    call_import(chunks, current, "ecma:globalThis", "get", 0, line);
    lset(&mut chunks[current], global, line);
    lget(&mut chunks[current], global, line);
    lget(&mut chunks[current], slot, line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    lset(&mut chunks[current], resolved, line);

    lget(&mut chunks[current], resolved, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], resolved, line);
    lset(&mut chunks[current], slot, line);
    chunks[current].emit_end(line);
}

fn preserve_object_prototype(
    chunks: &mut [Chunk],
    current: usize,
    source: u16,
    target: u16,
    line: u32,
) {
    lget(&mut chunks[current], target, line);
    lget(&mut chunks[current], source, line);
    call_import(chunks, current, "ecma:object", "getPrototypeOf", 1, line);
    call_import(chunks, current, "ecma:object", "setPrototypeOf", 2, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn preserve_python_class_metadata(
    chunks: &mut [Chunk],
    current: usize,
    source: u16,
    target: u16,
    line: u32,
) {
    copy_class_slot_if_present(chunks, current, source, target, &ClassSlot::ProtoLink, line);
    for key in ["__class__", "__type", "__types"] {
        copy_property_if_present(chunks, current, source, target, key, line);
    }
    for method in [
        "__str__",
        "__repr__",
        "__copy__",
        "__deepcopy__",
        "__reduce__",
        "__getstate__",
        "__setstate__",
    ] {
        copy_method_if_present(chunks, current, source, target, method, line);
    }
    delete_property_if_present(chunks, current, target, "constructor", line);
    preserve_prototype_from_constructor(chunks, current, source, target, line);
}

fn copy_class_slot_if_present(
    chunks: &mut [Chunk],
    current: usize,
    source: u16,
    target: u16,
    key: &ClassSlot,
    line: u32,
) {
    let value = chunks[current].alloc_scratch(1);
    emit_class_slot_get(chunks, current, source, key, line);
    lset(&mut chunks[current], value, line);
    emit_slot_is_null_or_undefined(chunks, current, value, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], target, line);
    lget(&mut chunks[current], value, line);
    let resolved = class_slots::resolve(key, &PlainNames);
    class_slots::emit_class_set(
        &mut chunks[current],
        ObjSource::Stack,
        &resolved,
        vybe_compiler::primitives::class_slots::ValueSource::Stack,
        line,
    );
    chunks[current].emit_end(line);
}

fn set_class_slot_from_stack(chunks: &mut [Chunk], current: usize, key: &ClassSlot, line: u32) {
    let resolved = class_slots::resolve(key, &PlainNames);
    class_slots::emit_class_set(
        &mut chunks[current],
        ObjSource::Stack,
        &resolved,
        ValueSource::Stack,
        line,
    );
}

fn copy_method_if_present(
    chunks: &mut [Chunk],
    current: usize,
    source: u16,
    target: u16,
    key: &str,
    line: u32,
) {
    let value = chunks[current].alloc_scratch(1);
    lookup_method(chunks, current, source, key, value, line);
    emit_slot_is_null_or_undefined(chunks, current, value, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], target, line);
    chunks[current].emit_string_const(key, line);
    lget(&mut chunks[current], value, line);
    call_import(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

fn delete_property_if_present(
    chunks: &mut [Chunk],
    current: usize,
    target: u16,
    key: &str,
    line: u32,
) {
    lget(&mut chunks[current], target, line);
    chunks[current].emit_string_const(key, line);
    call_import(chunks, current, "ecma:object", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn copy_property_if_present(
    chunks: &mut [Chunk],
    current: usize,
    source: u16,
    target: u16,
    key: &str,
    line: u32,
) {
    let value = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], source, line);
    chunks[current].emit_string_const(key, line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    lset(&mut chunks[current], value, line);
    emit_slot_is_null_or_undefined(chunks, current, value, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], target, line);
    chunks[current].emit_string_const(key, line);
    lget(&mut chunks[current], value, line);
    call_import(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

fn preserve_prototype_from_constructor(
    chunks: &mut [Chunk],
    current: usize,
    source: u16,
    target: u16,
    line: u32,
) {
    let ctor = chunks[current].alloc_scratch(1);
    let proto = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], source, line);
    chunks[current].emit_string_const("constructor", line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    lset(&mut chunks[current], ctor, line);
    lget(&mut chunks[current], ctor, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], ctor, line);
    chunks[current].emit_string_const("prototype", line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    lset(&mut chunks[current], proto, line);
    lget(&mut chunks[current], proto, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], target, line);
    lget(&mut chunks[current], proto, line);
    call_import(chunks, current, "ecma:object", "setPrototypeOf", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn store_value(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    let global = chunks[current].alloc_scratch(1);
    let store = chunks[current].alloc_scratch(1);
    emit_store_object(chunks, current, global, store, line);

    lget(&mut chunks[current], store, line);
    chunks[current].emit_string_const(PICKLE_ID, line);
    lget(&mut chunks[current], value_slot, line);
    call_import(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn load_value(chunks: &mut [Chunk], current: usize, envelope_slot: u16, line: u32) {
    let global = chunks[current].alloc_scratch(1);
    let store = chunks[current].alloc_scratch(1);
    let id = chunks[current].alloc_scratch(1);
    emit_store_object(chunks, current, global, store, line);

    lget(&mut chunks[current], envelope_slot, line);
    chunks[current].emit_string_const("id", line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    lset(&mut chunks[current], id, line);

    lget(&mut chunks[current], store, line);
    lget(&mut chunks[current], id, line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
}

fn emit_store_object(chunks: &mut [Chunk], current: usize, global: u16, store: u16, line: u32) {
    emit_named_store_object(chunks, current, STORE_KEY, global, store, line);
}

fn emit_file_store_object(chunks: &mut [Chunk], current: usize, global: u16, store: u16, line: u32) {
    emit_named_store_object(chunks, current, FILE_STORE_KEY, global, store, line);
}

fn emit_stream_store_map(chunks: &mut [Chunk], current: usize, global: u16, store: u16, line: u32) {
    call_import(chunks, current, "ecma:globalThis", "get", 0, line);
    lset(&mut chunks[current], global, line);

    lget(&mut chunks[current], global, line);
    chunks[current].emit_string_const(STREAM_STORE_KEY, line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    lset(&mut chunks[current], store, line);

    emit_slot_is_null_or_undefined(chunks, current, store, line);
    chunks[current].emit_if(line);
    call_import(chunks, current, "ecma:map", "new", 0, line);
    lset(&mut chunks[current], store, line);
    lget(&mut chunks[current], global, line);
    chunks[current].emit_string_const(STREAM_STORE_KEY, line);
    lget(&mut chunks[current], store, line);
    call_import(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

fn emit_named_store_object(
    chunks: &mut [Chunk],
    current: usize,
    key: &str,
    global: u16,
    store: u16,
    line: u32,
) {
    call_import(chunks, current, "ecma:globalThis", "get", 0, line);
    lset(&mut chunks[current], global, line);

    lget(&mut chunks[current], global, line);
    chunks[current].emit_string_const(key, line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    lset(&mut chunks[current], store, line);

    lget(&mut chunks[current], store, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    new_object(&mut chunks[current], line);
    lset(&mut chunks[current], store, line);
    lget(&mut chunks[current], global, line);
    chunks[current].emit_string_const(key, line);
    lget(&mut chunks[current], store, line);
    call_import(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

fn store_current_pickle_for_path(chunks: &mut [Chunk], current: usize, path_slot: u16, line: u32) {
    let global = chunks[current].alloc_scratch(1);
    let store = chunks[current].alloc_scratch(1);
    let file_store = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);

    emit_store_object(chunks, current, global, store, line);
    lget(&mut chunks[current], store, line);
    chunks[current].emit_string_const(PICKLE_ID, line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    lset(&mut chunks[current], value, line);

    emit_file_store_object(chunks, current, global, file_store, line);
    lget(&mut chunks[current], file_store, line);
    lget(&mut chunks[current], path_slot, line);
    lget(&mut chunks[current], value, line);
    call_import(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn load_file_value(chunks: &mut [Chunk], current: usize, path_slot: u16, line: u32) {
    let global = chunks[current].alloc_scratch(1);
    let file_store = chunks[current].alloc_scratch(1);

    emit_file_store_object(chunks, current, global, file_store, line);
    lget(&mut chunks[current], file_store, line);
    lget(&mut chunks[current], path_slot, line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
}

fn store_current_pickle_for_stream(chunks: &mut [Chunk], current: usize, stream_slot: u16, line: u32) {
    let global = chunks[current].alloc_scratch(1);
    let store = chunks[current].alloc_scratch(1);
    let stream_store = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);

    emit_store_object(chunks, current, global, store, line);
    lget(&mut chunks[current], store, line);
    chunks[current].emit_string_const(PICKLE_ID, line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    lset(&mut chunks[current], value, line);

    emit_stream_store_map(chunks, current, global, stream_store, line);
    lget(&mut chunks[current], stream_store, line);
    lget(&mut chunks[current], stream_slot, line);
    lget(&mut chunks[current], value, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn load_stream_value(chunks: &mut [Chunk], current: usize, stream_slot: u16, line: u32) {
    let global = chunks[current].alloc_scratch(1);
    let stream_store = chunks[current].alloc_scratch(1);

    emit_stream_store_map(chunks, current, global, stream_store, line);
    lget(&mut chunks[current], stream_store, line);
    lget(&mut chunks[current], stream_slot, line);
    call_import(chunks, current, "ecma:map", "get", 2, line);
}

fn emit_envelope_bytes(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let payload = chunks[current].alloc_scratch(1);
    let default = chunks[current].alloc_scratch(1);
    let sort = chunks[current].alloc_scratch(1);
    let props = chunks[current].alloc_scratch(1);
    let norm = chunks[current].alloc_scratch(1);
    let item_sep = chunks[current].alloc_scratch(1);
    let kv_sep = chunks[current].alloc_scratch(1);
    let text = chunks[current].alloc_scratch(1);

    new_object(&mut chunks[current], line);
    lset(&mut chunks[current], payload, line);
    lget(&mut chunks[current], payload, line);
    chunks[current].emit_string_const("id", line);
    chunks[current].emit_string_const(PICKLE_ID, line);
    call_import(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_ref_null(HT_EXTERN, line);
    lset(&mut chunks[current], default, line);
    chunks[current].emit_bool_const(false, line);
    lset(&mut chunks[current], sort, line);
    chunks[current].emit_bool_const(true, line);
    lset(&mut chunks[current], props, line);
    chunks[current].emit_string_const(",", line);
    lset(&mut chunks[current], item_sep, line);
    chunks[current].emit_string_const(":", line);
    lset(&mut chunks[current], kv_sep, line);

    vybe_compiler::primitives::json::emit_normalize(
        chunks, current, payload, default, sort, props, line,
    );
    lset(&mut chunks[current], norm, line);
    vybe_compiler::primitives::json::emit_render_separated(
        chunks, current, norm, item_sep, kv_sep, line,
    );
    lset(&mut chunks[current], text, line);

    encode_string_slot(chunks, current, text, line);
}

fn encode_string_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    call_import(chunks, current, "web:encoding", "encoderNew", 0, line);
    lget(&mut chunks[current], slot, line);
    call_import(chunks, current, "web:encoding", "encode", 2, line);
}

fn decode_bytes_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    call_import(chunks, current, "web:encoding", "decoderNew", 0, line);
    lget(&mut chunks[current], slot, line);
    call_import(chunks, current, "web:encoding", "decode", 2, line);
}

fn emit_class_slot_get(
    chunks: &mut [Chunk],
    current: usize,
    object_slot: u16,
    key: &ClassSlot,
    line: u32,
) {
    lget(&mut chunks[current], object_slot, line);
    let resolved = class_slots::resolve(key, &PlainNames);
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Stack,
        &resolved,
        Dest::Stack,
        line,
    );
}

fn emit_is_python_file_obj(chunks: &mut [Chunk], current: usize, file_slot: u16, line: u32) {
    emit_class_slot_get(
        chunks,
        current,
        file_slot,
        &ClassSlot::internal("__fpath"),
        line,
    );
    let path = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], path, line);
    emit_slot_is_null_or_undefined(chunks, current, path, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
}

fn emit_python_file_path(
    chunks: &mut [Chunk],
    current: usize,
    file_slot: u16,
    out_slot: u16,
    line: u32,
) {
    emit_class_slot_get(
        chunks,
        current,
        file_slot,
        &ClassSlot::internal("__fpath"),
        line,
    );
    lset(&mut chunks[current], out_slot, line);
    emit_slot_is_null_or_undefined(chunks, current, out_slot, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], file_slot, line);
    chunks[current].emit_string_const("__fpath", line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    lset(&mut chunks[current], out_slot, line);
    emit_slot_is_null_or_undefined(chunks, current, out_slot, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], file_slot, line);
    chunks[current].emit_string_const("name", line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    lset(&mut chunks[current], out_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], out_slot, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
    lset(&mut chunks[current], out_slot, line);
    chunks[current].emit_end(line);
}

fn emit_slot_is_null_or_undefined(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    lget(&mut chunks[current], slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    lget(&mut chunks[current], slot, line);
    call_import(chunks, current, "wasm:js-undefined", "test", 1, line);
    chunks[current].emit_op(Op::I32_OR, line);
}

fn emit_path_slot_is_present(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    lget(&mut chunks[current], slot, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
}

fn emit_stream_method_call(
    chunks: &mut [Chunk],
    current: usize,
    receiver_slot: u16,
    method_name: &str,
    arg_slot: Option<u16>,
    line: u32,
) {
    let method = chunks[current].alloc_scratch(1);
    lookup_method(chunks, current, receiver_slot, method_name, method, line);
    let recv =
        vybe_compiler::primitives::callable::push_callback_from_slot(chunks, current, method, line);
    if let Some(arg) = arg_slot {
        lget(&mut chunks[current], arg, line);
    }
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(
        &mut chunks[current],
        recv + u8::from(arg_slot.is_some()),
        line,
    );
}

fn lookup_method(
    chunks: &mut [Chunk],
    current: usize,
    object_slot: u16,
    method_name: &str,
    out_slot: u16,
    line: u32,
) {
    vybe_compiler::primitives::globals::emit_read(
        &mut chunks[current],
        "__vybe_js_get_method",
        line,
    );
    lget(&mut chunks[current], object_slot, line);
    chunks[current].emit_string_const(method_name, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(&mut chunks[current], 2, line);
    lset(&mut chunks[current], out_slot, line);
}

fn if_not_object_return_original(chunk: &mut Chunk, value: u16, line: u32) {
    lget(chunk, value, line);
    vybe_compiler::primitives::instructions::recipes::is_object(chunk, line);
    lget(chunk, value, line);
    chunk.emit_ref_type_op(
        Op::REF_TEST,
        vybe_runtime::opcode::heaptype::HeapType::Abstract(
            vybe_runtime::opcode::heaptype::HT_STRUCT,
        ),
        line,
    );
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    lget(chunk, value, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.emit_end(line);
}
