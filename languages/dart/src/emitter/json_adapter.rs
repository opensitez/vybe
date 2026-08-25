//! `dart:convert` jsonEncode/jsonDecode — dart semantics over the shared
//! JSON machinery.
//!
//! Rendering and parsing stay shared: `ecma:json.stringify` renders and the
//! shared `ecma:json.parseOrNull` parses. What dart ADDS sits in one
//! recursive CLEAN pass (`__dart_json_clean`) run before stringify:
//!
//! - a non-finite double THROWS `JsonUnsupportedObjectError` (ECMA renders
//!   `null`, dart refuses);
//! - a CYCLE throws it too, detected with an identity `seen` stack (a DAG —
//!   the same list twice — stays legal, so the entry is popped on the way
//!   out);
//! - a class instance (a record stamped `__type`) without `toJson` support
//!   throws it;
//! - a dart MAP re-emerges as a CLEAN object — its keys read from the
//!   `__dart_map_order` insertion record when present, and every `__`-
//!   prefixed bookkeeping field is dropped. Feeding the raw record to
//!   stringify serialized the bookkeeping (or nothing at all).
//!
//! `jsonDecode` maps a parse failure to a TYPED `FormatException`:
//! `parseOrNull` answers null, and null-for-input-that-isn't-`"null"` IS the
//! failure — no try/catch machinery needed.

use vybe_compiler::primitives::{collections, ops};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const MAP_ORDER_KEY: &str = "__dart_map_order";

fn slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn add_call(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module.to_string(), name.to_string());
    chunk.emit_call(idx, argc, line);
}

fn ref_func(chunk: &mut Chunk, func_idx: usize, line: u32) {
    chunk.emit_op_u16(Op::REF_FUNC, func_idx as u16, line);
    chunk.emit(0, line);
}

fn call_ref(chunk: &mut Chunk, argc: u8, line: u32) {
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, argc, line);
}

/// `jsonEncode(value[, toEncodable])` — clean, then render.
/// Stack: `[value]` or `[value, toEncodable]` → `[string]`.
pub fn emit_dart_json_encode(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let helper = build_clean_helper(chunks, line);
    let c = &mut chunks[current];
    let value_slot = slot(c);
    let hook_slot = slot(c);
    if argc >= 2 {
        c.emit_op_u16(Op::LOCAL_SET, hook_slot, line);
    } else {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        c.emit_op_u16(Op::LOCAL_SET, hook_slot, line);
    }
    c.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    ref_func(c, helper, line);
    c.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    c.emit_array_new_fixed(0, 0, line); // seen = []
    c.emit_op_u16(Op::LOCAL_GET, hook_slot, line);
    call_ref(c, 3, line);
    add_call(c, "ecma:json", "stringify", 1, line);
}

/// `jsonDecode(text[, reviver])` — parse, typed `FormatException` on
/// failure, reviver applied bottom-up over the parsed tree.
/// Stack: `[text]` or `[text, reviver]` → `[value]`.
pub fn emit_dart_json_decode(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let text_slot = slot(&mut chunks[current]);
    let out_slot = slot(&mut chunks[current]);
    let reviver_slot = slot(&mut chunks[current]);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, reviver_slot, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, reviver_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    add_call(&mut chunks[current], "ecma:json", "parseOrNull", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    // Only the literal spelling "null" parses TO null legally.
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_string_const("null", line);
    chunks[current].emit_op(Op::STRING_EQ, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    crate::emitter::string_adapter::emit_dart_named_exception_throw(
        chunks,
        current,
        "FormatException",
        "Unexpected character",
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    // reviver: applied to every map member and list element, bottom-up.
    chunks[current].emit_op_u16(Op::LOCAL_GET, reviver_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    {
        let helper = build_revive_helper(chunks, line);
        ref_func(&mut chunks[current], helper, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, reviver_slot, line);
        call_ref(&mut chunks[current], 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    }
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

/// `__dart_json_revive(value, reviver)` — mutate the parsed tree in place:
/// each list element and map member becomes `reviver(key, revive(child))`,
/// children first (dart's contract, same shape as ECMA §25.5.1.1's
/// InternalizeJSONProperty minus the root call, which the corpus does not
/// observe).
fn build_revive_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let hidx = chunks.len();
    let mut h =
        vybe_compiler::primitives::functions::create_function_chunk("__dart_json_revive", 2);
    h.alloc_scratch(2); // params: value = 0, reviver = 1
    chunks.push(h);
    let (value_slot, reviver_slot) = (0u16, 1u16);

    // Only containers recurse; everything else returns as-is.
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[hidx].emit_op(Op::REF_IS_NULL, line);
    chunks[hidx].emit_if(line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[hidx].emit_op(Op::RETURN, line);
    chunks[hidx].emit_end(line);

    let n_slot = slot(&mut chunks[hidx]);
    let i_slot = slot(&mut chunks[hidx]);
    let key_slot = slot(&mut chunks[hidx]);

    // Array: v[i] = reviver(i, revive(v[i])).
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    add_call(&mut chunks[hidx], "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[hidx], line);
    chunks[hidx].emit_if(line);
    {
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        collections::emit_len(chunks, hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_SET, n_slot, line);
        chunks[hidx].emit_i32_const(0, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_SET, i_slot, line);
        let lp = vybe_compiler::primitives::loops::emit_loop_start(chunks, hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, n_slot, line);
        chunks[hidx].emit_op(Op::I32_LT_S, line);
        vybe_compiler::primitives::loops::emit_loop_cond(chunks, hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        // reviver(i, revive(v[i], reviver))
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, reviver_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        ref_func(&mut chunks[hidx], hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        collections::emit_get(chunks, hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, reviver_slot, line);
        call_ref(&mut chunks[hidx], 2, line);
        call_ref(&mut chunks[hidx], 2, line);
        collections::emit_set(chunks, hidx, line);
        chunks[hidx].emit_op(Op::DROP, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunks[hidx].emit_i32_const(1, line);
        chunks[hidx].emit_op(Op::I32_ADD, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_SET, i_slot, line);
        vybe_compiler::primitives::loops::emit_loop_end(chunks, hidx, lp, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[hidx].emit_op(Op::RETURN, line);
    }
    chunks[hidx].emit_end(line);

    // Object: v[k] = reviver(k, revive(v[k])) for own keys.
    kind_test(chunks, hidx, value_slot, "wasm:js-number", line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    add_call(&mut chunks[hidx], "wasm:js-string", "test", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[hidx], line);
    chunks[hidx].emit_op(Op::I32_OR, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    add_call(&mut chunks[hidx], "wasm:js-boolean", "test", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[hidx], line);
    chunks[hidx].emit_op(Op::I32_OR, line);
    chunks[hidx].emit_op(Op::I32_EQZ, line);
    chunks[hidx].emit_if(line);
    {
        let keys_slot = slot(&mut chunks[hidx]);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        add_call(&mut chunks[hidx], "ecma:object", "keys", 1, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
        collections::emit_len(chunks, hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_SET, n_slot, line);
        chunks[hidx].emit_i32_const(0, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_SET, i_slot, line);
        let lp = vybe_compiler::primitives::loops::emit_loop_start(chunks, hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, n_slot, line);
        chunks[hidx].emit_op(Op::I32_LT_S, line);
        vybe_compiler::primitives::loops::emit_loop_cond(chunks, hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        collections::emit_get(chunks, hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_SET, key_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, key_slot, line);
        // reviver(k, revive(v[k], reviver))
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, reviver_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, key_slot, line);
        ref_func(&mut chunks[hidx], hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, key_slot, line);
        add_call(&mut chunks[hidx], "ecma:object", "get", 2, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, reviver_slot, line);
        call_ref(&mut chunks[hidx], 2, line);
        call_ref(&mut chunks[hidx], 2, line);
        add_call(&mut chunks[hidx], "ecma:object", "set", 3, line);
        chunks[hidx].emit_op(Op::DROP, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunks[hidx].emit_i32_const(1, line);
        chunks[hidx].emit_op(Op::I32_ADD, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_SET, i_slot, line);
        vybe_compiler::primitives::loops::emit_loop_end(chunks, hidx, lp, line);
    }
    chunks[hidx].emit_end(line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[hidx].emit_op(Op::RETURN, line);
    hidx
}

/// Throw `JsonUnsupportedObjectError` from inside the helper chunk.
fn emit_unsupported_throw(chunks: &mut [Chunk], current: usize, message: &str, line: u32) {
    crate::emitter::string_adapter::emit_dart_named_exception_throw(
        chunks,
        current,
        "JsonUnsupportedObjectError",
        message,
        line,
    );
}

/// `[cond]` — is the value of the given wasm builtin kind?
fn kind_test(chunks: &mut [Chunk], current: usize, value_slot: u16, module: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    add_call(&mut chunks[current], module, "test", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
}

/// `__dart_json_clean(value, seen)` — see the module header.
fn build_clean_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let hidx = chunks.len();
    let mut h =
        vybe_compiler::primitives::functions::create_function_chunk("__dart_json_clean", 3);
    h.alloc_scratch(3); // params: value = 0, seen = 1, hook = 2
    chunks.push(h);
    let (value_slot, seen_slot, hook_slot) = (0u16, 1u16, 2u16);

    // null / undefined → null.
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[hidx].emit_op(Op::REF_IS_NULL, line);
    chunks[hidx].emit_if(line);
    chunks[hidx].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[hidx].emit_op(Op::RETURN, line);
    chunks[hidx].emit_end(line);
    kind_test(chunks, hidx, value_slot, "wasm:js-undefined", line);
    chunks[hidx].emit_if(line);
    chunks[hidx].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[hidx].emit_op(Op::RETURN, line);
    chunks[hidx].emit_end(line);

    // number → finite passes, NaN/±Infinity throw.
    kind_test(chunks, hidx, value_slot, "wasm:js-number", line);
    chunks[hidx].emit_if(line);
    {
        let f_slot = slot(&mut chunks[hidx]);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        add_call(&mut chunks[hidx], "wasm:js-number", "toF64", 1, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_TEE, f_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, f_slot, line);
        chunks[hidx].emit_op(Op::F64_NE, line); // NaN != NaN
        chunks[hidx].emit_if(line);
        emit_unsupported_throw(chunks, hidx, "Converting object to an encodable object failed: NaN", line);
        chunks[hidx].emit_end(line);
        for inf in [f64::INFINITY, f64::NEG_INFINITY] {
            chunks[hidx].emit_op_u16(Op::LOCAL_GET, f_slot, line);
            chunks[hidx].emit_f64_const(inf, line);
            chunks[hidx].emit_op(Op::F64_EQ, line);
            chunks[hidx].emit_if(line);
            emit_unsupported_throw(
                chunks,
                hidx,
                "Converting object to an encodable object failed: Infinity",
                line,
            );
            chunks[hidx].emit_end(line);
        }
        // Return the F64, not the original: a big int literal is a
        // `Value::I64`, which `ecma:json.stringify` cannot serialize —
        // `{'big': 9007199254740992}` rendered as NOTHING. 2^53 is exact in
        // f64, so the coercion is lossless exactly as far as JSON itself is.
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, f_slot, line);
        chunks[hidx].emit_op(Op::RETURN, line);
    }
    chunks[hidx].emit_end(line);

    // string / bool → pass.
    for module in ["wasm:js-string", "wasm:js-boolean"] {
        kind_test(chunks, hidx, value_slot, module, line);
        chunks[hidx].emit_if(line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[hidx].emit_op(Op::RETURN, line);
        chunks[hidx].emit_end(line);
    }

    // Cycle check — identity membership in `seen` (reference equality is
    // what `contains` answers for two handles to one object).
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, seen_slot, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_contains(chunks, hidx, line);
    ops::emit_dyn_to_bool(&mut chunks[hidx], line);
    chunks[hidx].emit_if(line);
    emit_unsupported_throw(chunks, hidx, "Cyclic object detected", line);
    chunks[hidx].emit_end(line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, seen_slot, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_push(chunks, hidx, line);
    chunks[hidx].emit_op(Op::DROP, line);

    let out_slot = slot(&mut chunks[hidx]);
    let n_slot = slot(&mut chunks[hidx]);
    let i_slot = slot(&mut chunks[hidx]);

    // Array → array of clean(elem).
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    add_call(&mut chunks[hidx], "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[hidx], line);
    chunks[hidx].emit_if(line);
    {
        collections::emit_array_new(chunks, hidx, 0, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_SET, out_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        collections::emit_len(chunks, hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_SET, n_slot, line);
        chunks[hidx].emit_i32_const(0, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_SET, i_slot, line);
        let lp = vybe_compiler::primitives::loops::emit_loop_start(chunks, hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, n_slot, line);
        chunks[hidx].emit_op(Op::I32_LT_S, line);
        vybe_compiler::primitives::loops::emit_loop_cond(chunks, hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, out_slot, line);
        // clean(value[i], seen)
        ref_func(&mut chunks[hidx], hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        collections::emit_get(chunks, hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, seen_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, hook_slot, line);
        call_ref(&mut chunks[hidx], 3, line);
        collections::emit_push(chunks, hidx, line);
        chunks[hidx].emit_op(Op::DROP, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunks[hidx].emit_i32_const(1, line);
        chunks[hidx].emit_op(Op::I32_ADD, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_SET, i_slot, line);
        vybe_compiler::primitives::loops::emit_loop_end(chunks, hidx, lp, line);
        // seen.pop() — a DAG may reuse this array after we return.
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, seen_slot, line);
        collections::emit_pop(chunks, hidx, line);
        chunks[hidx].emit_op(Op::DROP, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, out_slot, line);
        chunks[hidx].emit_op(Op::RETURN, line);
    }
    chunks[hidx].emit_end(line);

    // `toJson()` first — dart's encoder hook, and what ECMA stringify's own
    // `toJSON` step (§25.5.2.3) did before this clean pass ran ahead of it.
    // The method lives on the prototype; the host `get` proto-walks. Invoked
    // with the ambient receiver, exactly as compiled method dispatch does.
    let tojson_slot = slot(&mut chunks[hidx]);
    let saved_this_slot = slot(&mut chunks[hidx]);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[hidx].emit_string_const("toJson", line);
    add_call(&mut chunks[hidx], "ecma:object", "get", 2, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_SET, tojson_slot, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, tojson_slot, line);
    chunks[hidx].emit_op(Op::REF_IS_NULL, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, tojson_slot, line);
    add_call(&mut chunks[hidx], "wasm:js-undefined", "test", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[hidx], line);
    chunks[hidx].emit_op(Op::I32_OR, line);
    chunks[hidx].emit_op(Op::I32_EQZ, line);
    chunks[hidx].emit_if(line);
    {
        vybe_compiler::primitives::globals::emit_read(&mut chunks[hidx], "__js_this", line);
        chunks[hidx].emit_op_u16(Op::LOCAL_SET, saved_this_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        vybe_compiler::primitives::globals::emit_write(&mut chunks[hidx], "__js_this", line);
        // clean(value.toJson(), seen)
        ref_func(&mut chunks[hidx], hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, tojson_slot, line);
        call_ref(&mut chunks[hidx], 0, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, seen_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, hook_slot, line);
        call_ref(&mut chunks[hidx], 3, line);
        let cleaned_slot = slot(&mut chunks[hidx]);
        chunks[hidx].emit_op_u16(Op::LOCAL_SET, cleaned_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, saved_this_slot, line);
        vybe_compiler::primitives::globals::emit_write(&mut chunks[hidx], "__js_this", line);
        // pop the cycle entry for this value before returning.
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, seen_slot, line);
        collections::emit_pop(chunks, hidx, line);
        chunks[hidx].emit_op(Op::DROP, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, cleaned_slot, line);
        chunks[hidx].emit_op(Op::RETURN, line);
    }
    chunks[hidx].emit_end(line);

    // A class instance (record stamped `__type`) has no JSON form here —
    // dart's contract is JsonUnsupportedObjectError.
    let type_slot = slot(&mut chunks[hidx]);
    {
        let key = chunks[hidx].add_constant(vybe_runtime::Value::String(std::sync::Arc::from(
            "__type",
        )));
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[hidx].emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_SET, type_slot, line);
    }
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, type_slot, line);
    chunks[hidx].emit_op(Op::REF_IS_NULL, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, type_slot, line);
    add_call(&mut chunks[hidx], "wasm:js-undefined", "test", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[hidx], line);
    chunks[hidx].emit_op(Op::I32_OR, line);
    chunks[hidx].emit_op(Op::I32_EQZ, line);
    chunks[hidx].emit_if(line);
    {
        // `toEncodable:` — dart's fallback hook for a non-encodable object;
        // only when absent does the typed error throw.
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, hook_slot, line);
        chunks[hidx].emit_op(Op::REF_IS_NULL, line);
        chunks[hidx].emit_op(Op::I32_EQZ, line);
        chunks[hidx].emit_if(line);
        {
            // clean(hook(value), seen, hook)
            ref_func(&mut chunks[hidx], hidx, line);
            chunks[hidx].emit_op_u16(Op::LOCAL_GET, hook_slot, line);
            chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
            call_ref(&mut chunks[hidx], 1, line);
            chunks[hidx].emit_op_u16(Op::LOCAL_GET, seen_slot, line);
            chunks[hidx].emit_op_u16(Op::LOCAL_GET, hook_slot, line);
            call_ref(&mut chunks[hidx], 3, line);
            let hooked_slot = slot(&mut chunks[hidx]);
            chunks[hidx].emit_op_u16(Op::LOCAL_SET, hooked_slot, line);
            chunks[hidx].emit_op_u16(Op::LOCAL_GET, seen_slot, line);
            collections::emit_pop(chunks, hidx, line);
            chunks[hidx].emit_op(Op::DROP, line);
            chunks[hidx].emit_op_u16(Op::LOCAL_GET, hooked_slot, line);
            chunks[hidx].emit_op(Op::RETURN, line);
        }
        chunks[hidx].emit_end(line);
        emit_unsupported_throw(
            chunks,
            hidx,
            "Converting object to an encodable object failed",
            line,
        );
    }
    chunks[hidx].emit_end(line);

    // Map → clean object: keys from the insertion-order record when present
    // (falling back to Object.keys), `__`-prefixed bookkeeping dropped.
    let keys_slot = slot(&mut chunks[hidx]);
    let key_slot = slot(&mut chunks[hidx]);
    {
        let order_key = chunks[hidx].add_constant(vybe_runtime::Value::String(
            std::sync::Arc::from(MAP_ORDER_KEY),
        ));
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[hidx].emit_struct_field_op(Op::STRUCT_GET, 0, order_key, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    }
    kind_test(chunks, hidx, keys_slot, "wasm:js-undefined", line);
    chunks[hidx].emit_if(line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    add_call(&mut chunks[hidx], "ecma:object", "keys", 1, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    chunks[hidx].emit_end(line);

    chunks[hidx].emit_struct_new(0, 0, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    collections::emit_len(chunks, hidx, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_SET, n_slot, line);
    chunks[hidx].emit_i32_const(0, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let lp = vybe_compiler::primitives::loops::emit_loop_start(chunks, hidx, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, n_slot, line);
    chunks[hidx].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, hidx, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    collections::emit_get(chunks, hidx, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    // skip bookkeeping keys
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    add_call(&mut chunks[hidx], "ecma:string", "String", 1, line);
    chunks[hidx].emit_string_const("__", line);
    add_call(&mut chunks[hidx], "ecma:string", "startsWith", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[hidx], line);
    chunks[hidx].emit_op(Op::I32_EQZ, line);
    chunks[hidx].emit_if(line);
    {
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, out_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, key_slot, line);
        add_call(&mut chunks[hidx], "ecma:string", "String", 1, line);
        // clean(value[key], seen)
        ref_func(&mut chunks[hidx], hidx, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, key_slot, line);
        add_call(&mut chunks[hidx], "ecma:object", "get", 2, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, seen_slot, line);
        chunks[hidx].emit_op_u16(Op::LOCAL_GET, hook_slot, line);
        call_ref(&mut chunks[hidx], 3, line);
        add_call(&mut chunks[hidx], "ecma:object", "set", 3, line);
        chunks[hidx].emit_op(Op::DROP, line);
    }
    chunks[hidx].emit_end(line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[hidx].emit_i32_const(1, line);
    chunks[hidx].emit_op(Op::I32_ADD, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, hidx, lp, line);

    chunks[hidx].emit_op_u16(Op::LOCAL_GET, seen_slot, line);
    collections::emit_pop(chunks, hidx, line);
    chunks[hidx].emit_op(Op::DROP, line);
    chunks[hidx].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[hidx].emit_op(Op::RETURN, line);

    hidx
}
