//! Python `json.dumps` — Python semantics over the shared `vybe_compiler::primitives::json`
//! core.
//!
//! The walker reshapes `json.dumps(obj, cls=…, default=…, sort_keys=…,
//! indent=…, separators=…)` into the fixed positional form
//! `__py_json_dumps(value, default, sort_keys, indent, item_sep, kv_sep)`
//! (see `walker::rewrite_json_dumps`). This adapter normalizes the value tree
//! (applying the `default`/`cls` encoder hook to non-serializable objects like
//! `datetime`), then renders it: indented output goes through
//! `ecma:json.stringify` (byte-identical to Python's), compact output through
//! the shared separator-aware renderer so Python's `", "` / `": "` defaults
//! come out right.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use vybe_compiler::primitives::class_slots::{self, ClassSlot, ObjSource, PlainNames, ValueSource};
use vybe_compiler::primitives::{callable, collections, functions, ops};

use super::adapter_util::{call_import, lget, lset, stash_exact, struct_set};

const JSON_TO_PY_CHUNK: &str = "__py_json_to_python";

/// `emit = "common:python.json_dumps"`.
/// Stack in (bottom→top): value, default, sort_keys, indent, item_sep, kv_sep.
/// Leaves the JSON string on the stack.
pub fn emit_json_dumps(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    // Self-sufficient target: the walker's rewrite supplies all six, but a
    // BINDING (`from json import dumps`) calls it with only the positional
    // args the user wrote. Fill the missing trailing slots with Python's own
    // defaults so the adapter behaves the same either way — that is what lets
    // a name bound to this target match `json.dumps(...)` exactly.
    if argc < 6 {
        let c = &mut chunks[current];
        // default=None, sort_keys=False, indent=None, separators=(", ", ": ")
        for i in argc..6 {
            match i {
                1 | 3 => {
                    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
                }
                2 => c.emit_i32_const(0, line),
                4 => c.emit_string_const(", ", line),
                _ => c.emit_string_const(": ", line),
            }
        }
    }

    let (
        value_slot,
        default_slot,
        sort_slot,
        indent_slot,
        item_slot,
        kv_slot,
        props_slot,
        norm_slot,
    ) = {
        let c = &mut chunks[current];
        (
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
        )
    };

    {
        let c = &mut chunks[current];
        // Pop the six args (top → bottom).
        c.emit_op_u16(Op::LOCAL_SET, kv_slot, line);
        c.emit_op_u16(Op::LOCAL_SET, item_slot, line);
        c.emit_op_u16(Op::LOCAL_SET, indent_slot, line);
        c.emit_op_u16(Op::LOCAL_SET, sort_slot, line);
        c.emit_op_u16(Op::LOCAL_SET, default_slot, line);
        c.emit_op_u16(Op::LOCAL_SET, value_slot, line);
        // props = false — Python routes class instances to the encoder hook.
        c.emit_bool_const(false, line);
        c.emit_op_u16(Op::LOCAL_SET, props_slot, line);
    }

    // normalized = normalize(value, default, sort_keys, props=false)
    vybe_compiler::primitives::json::emit_normalize(
        chunks,
        current,
        value_slot,
        default_slot,
        sort_slot,
        props_slot,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, norm_slot, line);

    // indent is not None → ecma:json.stringify(normalized, null, indent), whose
    // indented output is byte-identical to Python's. Otherwise render with the
    // (Python-default or caller-supplied) compact separators.
    {
        let c = &mut chunks[current];
        c.emit_op_u16(Op::LOCAL_GET, indent_slot, line);
        c.emit_op(Op::REF_IS_NULL, line);
        c.emit_op(Op::I32_EQZ, line);
        c.emit_if_value(line);
        c.emit_op_u16(Op::LOCAL_GET, norm_slot, line);
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        c.emit_op_u16(Op::LOCAL_GET, indent_slot, line);
        let idx = c.add_import("ecma:json", "stringify");
        c.emit_call(idx, 3, line);
        c.emit_else(line);
    }
    vybe_compiler::primitives::json::emit_render_separated(
        chunks, current, norm_slot, item_slot, kv_slot, line,
    );
    chunks[current].emit_end(line);
}

/// `emit = "common:python.json_loads"`.
/// Stack: `[text] -> [value|null]`. Literal syntax errors are normalized by
/// the walker into `JSONDecodeError`; dynamic invalid JSON falls back to the
/// shared parse-or-null primitive rather than raw host trapping.
pub fn emit_json_loads(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let _base = stash_exact(chunks, current, argc, 1, line);
    lget(&mut chunks[current], _base, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    vybe_compiler::primitives::json::emit_parse_or_null(chunks, current, line);
    let parsed = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], parsed, line);
    let helper = ensure_json_to_python_chunk(chunks, line);
    chunks[current].emit_op_u16(Op::REF_FUNC, helper as u16, line);
    chunks[current].emit(0, line);
    lget(&mut chunks[current], parsed, line);
    callable::emit_direct_invoke_chunk(&mut chunks[current], 1, line);
}

/// `loads(dumps(value))` without relying on the intermediate string object's
/// runtime representation. The value still goes through the shared JSON
/// normalizer, then through Python's JSON value adapter.
pub fn emit_json_roundtrip(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    let value_slot = base;
    emit_is_json_map(chunks, current, value_slot, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], value_slot, line);
    chunks[current].emit_op(Op::RETURN, line);
    chunks[current].emit_end(line);

    let default_slot = chunks[current].alloc_scratch(1);
    let sort_slot = chunks[current].alloc_scratch(1);
    let props_slot = chunks[current].alloc_scratch(1);
    let norm_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    lset(&mut chunks[current], default_slot, line);
    chunks[current].emit_bool_const(false, line);
    lset(&mut chunks[current], sort_slot, line);
    chunks[current].emit_bool_const(true, line);
    lset(&mut chunks[current], props_slot, line);
    vybe_compiler::primitives::json::emit_normalize(
        chunks,
        current,
        value_slot,
        default_slot,
        sort_slot,
        props_slot,
        line,
    );
    lset(&mut chunks[current], norm_slot, line);
    emit_json_to_python_call(chunks, current, norm_slot, line);
}

fn ensure_json_to_python_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    if let Some(idx) = chunks.iter().position(|c| c.name == JSON_TO_PY_CHUNK) {
        return idx;
    }

    let idx = chunks.len();
    chunks.push(functions::create_function_chunk(JSON_TO_PY_CHUNK, 1));
    // `create_function_chunk` records arity but does not advance the scratch
    // allocator. Reserve the parameter slot so helper temporaries start at 1
    // and do not overwrite the JSON value being converted.
    chunks[idx].alloc_scratch(1);
    emit_json_to_python_body(chunks, idx, line);
    idx
}

fn emit_json_to_python_body(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let value = 0u16;
    let out = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let item = chunks[current].alloc_scratch(1);
    let converted = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::RETURN, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], value, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::RETURN, line);
    chunks[current].emit_end(line);

    emit_is_json_map(chunks, current, value, line);
    chunks[current].emit_if(line);
    emit_convert_json_object(
        chunks, current, value, out, keys, n, i, key, item, converted, line,
    );
    lget(&mut chunks[current], out, line);
    chunks[current].emit_op(Op::RETURN, line);
    chunks[current].emit_end(line);

    emit_is_plain_json_object(chunks, current, value, line);
    chunks[current].emit_if(line);
    emit_convert_json_object(
        chunks, current, value, out, keys, n, i, key, item, converted, line,
    );
    lget(&mut chunks[current], out, line);
    chunks[current].emit_op(Op::RETURN, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::RETURN, line);
}

#[allow(clippy::too_many_arguments)]
fn emit_convert_json_object(
    chunks: &mut Vec<Chunk>,
    current: usize,
    value: u16,
    out: u16,
    keys: u16,
    n: u16,
    i: u16,
    key: u16,
    item: u16,
    converted: u16,
    line: u32,
) {
    collections::emit_map_new(chunks, current, line);
    lset(&mut chunks[current], out, line);
    lget(&mut chunks[current], value, line);
    chunks[current].emit_string_const("__keys", line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    lset(&mut chunks[current], keys, line);
    lget(&mut chunks[current], keys, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    lget(&mut chunks[current], keys, line);
    call_import(chunks, current, "wasm:js-undefined", "test", 1, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], value, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    lset(&mut chunks[current], keys, line);
    chunks[current].emit_end(line);
    lget(&mut chunks[current], keys, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    lset(&mut chunks[current], n, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], i, line);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    lget(&mut chunks[current], i, line);
    lget(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    lget(&mut chunks[current], keys, line);
    lget(&mut chunks[current], i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(&mut chunks[current], key, line);
    emit_is_json_map(chunks, current, value, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], value, line);
    lget(&mut chunks[current], key, line);
    call_import(chunks, current, "ecma:map", "get", 2, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], value, line);
    lget(&mut chunks[current], key, line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    chunks[current].emit_end(line);
    lset(&mut chunks[current], item, line);
    emit_json_to_python_call(chunks, current, item, line);
    lset(&mut chunks[current], converted, line);
    lget(&mut chunks[current], out, line);
    lget(&mut chunks[current], key, line);
    lget(&mut chunks[current], converted, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
}

fn emit_json_to_python_call(chunks: &mut Vec<Chunk>, current: usize, slot: u16, line: u32) {
    let helper = ensure_json_to_python_chunk(chunks, line);
    chunks[current].emit_op_u16(Op::REF_FUNC, helper as u16, line);
    chunks[current].emit(0, line);
    lget(&mut chunks[current], slot, line);
    callable::emit_direct_invoke_chunk(&mut chunks[current], 1, line);
}

fn emit_is_plain_json_object(chunks: &mut Vec<Chunk>, current: usize, value: u16, line: u32) {
    lget(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);

    lget(&mut chunks[current], value, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("object", line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
    chunks[current].emit_op(Op::I32_AND, line);

    lget(&mut chunks[current], value, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_AND, line);
}

fn emit_is_json_map(chunks: &mut Vec<Chunk>, current: usize, value: u16, line: u32) {
    lget(&mut chunks[current], value, line);
    call_import(chunks, current, "ecma:object", "toStringTag", 1, line);
    let tag = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], tag, line);
    lget(&mut chunks[current], tag, line);
    chunks[current].emit_string_const("Map", line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
    lget(&mut chunks[current], tag, line);
    chunks[current].emit_string_const("[object Map]", line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
    chunks[current].emit_op(Op::I32_OR, line);
}

/// Python `JSONDecodeError` has ordinary exception behaviour plus parse
/// position fields. Stack in: `msg, doc, pos, lineno, colno`; short direct
/// constructor calls are padded with nulls by the adapter utility.
pub fn emit_json_decode_error(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 5, line);
    lget(&mut chunks[current], base, line);
    crate::emitter::runtime_adapter::emit_py_exception(chunks, current, 1, "JSONDecodeError", line);
    let obj = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], obj, line);

    let field = |chunk: &mut Chunk, obj: u16, key: &str, slot: u16, line: u32| {
        lget(chunk, obj, line);
        lget(chunk, slot, line);
        struct_set(chunk, &ClassSlot::internal(key), line);
    };
    field(&mut chunks[current], obj, "msg", base, line);
    field(&mut chunks[current], obj, "doc", base + 1, line);
    field(&mut chunks[current], obj, "pos", base + 2, line);
    field(&mut chunks[current], obj, "lineno", base + 3, line);
    field(&mut chunks[current], obj, "colno", base + 4, line);

    lget(&mut chunks[current], obj, line);
    let type_slot = class_slots::resolve(&ClassSlot::TypeIdentity, &PlainNames);
    class_slots::emit_class_set(
        &mut chunks[current],
        ObjSource::Stack,
        &type_slot,
        ValueSource::ConstStr("JSONDecodeError".to_string()),
        line,
    );
    lget(&mut chunks[current], obj, line);
}

/// Python `raise JSONDecodeError(...)`.
///
/// Stack in: `msg, doc, pos, lineno, colno`. Emits a throw and leaves null for
/// expression-position continuity.
pub fn emit_json_decode_error_raise(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_json_decode_error(chunks, current, argc, line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}
