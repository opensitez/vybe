//! Shared cross-language JSON — wider than `ecma:json`.
//!
//! `ecma:json.stringify` is the ECMA-262 §25.5 surface and stays pure ECMA, but
//! the language dialects need more than it can express: Python `json.dumps`
//! uses `", "` / `": "` default separators, `sort_keys`, and `cls=`/`default=`
//! encoder hooks; PHP `json_encode` normalizes associative arrays and honours
//! `JsonSerializable`. This module is the compatibility hinge — the same way
//! `xml.rs` gives Go/​.NET/​Java/​DOM one portable QName shape.
//!
//! The cross-language heart is [`emit_normalize`]: a recursive walk that turns
//! any value tree into a JSON-serializable shape (`Map`→`Object`, optional
//! `sort_keys`, and — for class instances a dialect can't serialize natively —
//! an encoder-hook callback, *not* the ECMA replacer). Rendering then delegates
//! to `ecma:json.stringify` wherever separators already match (all indented
//! output, PHP/JS compact); [`emit_render_separated`] handles only the
//! compact-with-spaces case ECMA can't produce (Python default output).

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use crate::primitives::functions::create_function_chunk;
use crate::primitives::loops::LoopState;

// ── low-level emit helpers (mirror the per-adapter convention) ───────────────

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    chunk.emit_string_const(value, line);
}

/// `REF_FUNC idx` with a zero upvalue count (helper chunks capture nothing).
fn ref_func(chunk: &mut Chunk, func_idx: usize, line: u32) {
    chunk.emit_op_u16(Op::REF_FUNC, func_idx as u16, line);
    chunk.emit(0, line);
}

fn call_ref(chunk: &mut Chunk, argc: u8, line: u32) {
    chunk.emit_op(Op::CALL_REF, line);
    chunk.emit(argc, line);
}

fn add_call(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module.to_string(), name.to_string());
    chunk.emit_call(idx, argc, line);
}

fn dyn_get(chunk: &mut Chunk, obj_slot: u16, key: &str, line: u32) {
    lget(chunk, obj_slot, line);
    push_str(chunk, key, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

fn loop_start(chunk: &mut Chunk, line: u32) -> LoopState {
    let block_patch = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    LoopState {
        block_patch,
        loop_patch,
        body_block_patch: None,
    }
}

/// Emit `if !(i < n) break;` for a numeric counter loop condition.
fn loop_break_unless_lt(chunk: &mut Chunk, i_slot: u16, n_slot: u16, line: u32) {
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    crate::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);
}

fn loop_end(chunk: &mut Chunk, state: LoopState, line: u32) {
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(state.loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(state.block_patch);
}

fn bump(chunk: &mut Chunk, i_slot: u16, line: u32) {
    lget(chunk, i_slot, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
}

/// `if value passes wasm:js-<kind>.test → return value`.
fn return_if_kind(chunk: &mut Chunk, value_slot: u16, kind: &str, line: u32) {
    lget(chunk, value_slot, line);
    let idx = chunk.add_import(format!("wasm:js-{kind}"), "test");
    chunk.emit_call(idx, 1, line);
    chunk.emit_if(line);
    lget(chunk, value_slot, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.emit_end(line);
}

/// Push i32 `1` when `value` is a container-ish object (not null / number /
/// string / boolean); i32 `0` otherwise. Consumes nothing, leaves an i32.
fn push_is_object(chunk: &mut Chunk, value_slot: u16, line: u32) {
    // not null
    lget(chunk, value_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    // AND not number
    lget(chunk, value_slot, line);
    let num = chunk.add_import("wasm:js-number", "test");
    chunk.emit_call(num, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    // AND not string
    lget(chunk, value_slot, line);
    let s = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(s, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    // AND not boolean
    lget(chunk, value_slot, line);
    let b = chunk.add_import("wasm:js-boolean", "test");
    chunk.emit_call(b, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
}

// ── normalize helper ─────────────────────────────────────────────────────────

/// Build the recursive `__json_normalize(value, default, sortKeys, props)`
/// helper chunk and return its index.
///
/// * `default` — encoder-hook funcref (or null). Called for a class instance
///   the dialect can't serialize natively; its result is normalized in turn.
/// * `sortKeys` — sort object keys lexicographically.
/// * `props` — when truthy, every object serializes its own enumerable keys
///   (PHP / plain dict). When falsy, an object carrying an own `__type`
///   (a class instance, e.g. Python `datetime`) routes to the hook instead.
fn build_normalize_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let helper_idx = chunks.len();
    let mut h = create_function_chunk("__json_normalize", 4);
    h.alloc_scratch(4); // reserve the four param slots

    let value_slot = 0u16;
    let default_slot = 1u16;
    let sort_slot = 2u16;
    let props_slot = 3u16;

    let out_slot = alloc_local(&mut h);
    let keys_slot = alloc_local(&mut h);
    let key_slot = alloc_local(&mut h);
    let i_slot = alloc_local(&mut h);
    let n_slot = alloc_local(&mut h);

    // null → pass through.
    lget(&mut h, value_slot, line);
    h.emit_op(Op::REF_IS_NULL, line);
    h.emit_if(line);
    lget(&mut h, value_slot, line);
    h.emit_op(Op::RETURN, line);
    h.emit_end(line);

    // undefined / primitives → pass through.
    return_if_kind(&mut h, value_slot, "undefined", line);
    return_if_kind(&mut h, value_slot, "number", line);
    return_if_kind(&mut h, value_slot, "string", line);
    return_if_kind(&mut h, value_slot, "boolean", line);

    // Array → array of normalize(elem).
    lget(&mut h, value_slot, line);
    add_call(&mut h, "ecma:array", "isArray", 1, line);
    crate::primitives::ops::emit_dyn_to_bool(&mut h, line);
    h.emit_if(line);
    {
        h.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        lset(&mut h, out_slot, line);
        lget(&mut h, value_slot, line);
        h.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut h, n_slot, line);
        h.emit_f64_const(0.0, line);
        lset(&mut h, i_slot, line);
        let lp = loop_start(&mut h, line);
        loop_break_unless_lt(&mut h, i_slot, n_slot, line);
        lget(&mut h, out_slot, line);
        // normalize(value[i], default, sortKeys, props)
        ref_func(&mut h, helper_idx, line);
        lget(&mut h, value_slot, line);
        lget(&mut h, i_slot, line);
        h.emit_op(Op::ARRAY_GET, line);
        lget(&mut h, default_slot, line);
        lget(&mut h, sort_slot, line);
        lget(&mut h, props_slot, line);
        call_ref(&mut h, 4, line);
        add_call(&mut h, "ecma:array", "push", 2, line);
        h.emit_op(Op::DROP, line);
        bump(&mut h, i_slot, line);
        loop_end(&mut h, lp, line);
        lget(&mut h, out_slot, line);
        h.emit_op(Op::RETURN, line);
    }
    h.emit_end(line);

    // Class instance (own `__type`, and not props-mode) → encoder hook / throw.
    lget(&mut h, value_slot, line);
    push_str(&mut h, "__type", line);
    add_call(&mut h, "ecma:object", "hasOwn", 2, line);
    crate::primitives::ops::emit_dyn_to_bool(&mut h, line); // i32: hasType
    lget(&mut h, props_slot, line);
    crate::primitives::ops::emit_dyn_to_bool(&mut h, line);
    h.emit_op(Op::I32_EQZ, line); // i32: !props
    h.emit_op(Op::I32_AND, line);
    h.emit_if(line);
    {
        // default != null ?
        lget(&mut h, default_slot, line);
        h.emit_op(Op::REF_IS_NULL, line);
        h.emit_op(Op::I32_EQZ, line);
        h.emit_if(line);
        {
            // normalize(default(value), default, sortKeys, props)
            ref_func(&mut h, helper_idx, line);
            lget(&mut h, default_slot, line);
            lget(&mut h, value_slot, line);
            call_ref(&mut h, 1, line);
            lget(&mut h, default_slot, line);
            lget(&mut h, sort_slot, line);
            lget(&mut h, props_slot, line);
            call_ref(&mut h, 4, line);
            h.emit_op(Op::RETURN, line);
        }
        h.emit_else(line);
        {
            // throw TypeError("Object of type <T> is not JSON serializable")
            h.emit_op_u16(Op::STRUCT_NEW, 0, line);
            h.emit_dup(line);
            push_str(&mut h, "Object of type ", line);
            dyn_get(&mut h, value_slot, "__type", line);
            crate::primitives::convert::emit_to_string(&mut h, line);
            push_str(&mut h, " is not JSON serializable", line);
            crate::primitives::strings::emit_concat(&mut h, 3, line);
            crate::primitives::errors::emit_exception_new_finalize(&mut h, "TypeError", line);
            crate::primitives::errors::emit_throw(&mut h, line);
        }
        h.emit_end(line);
    }
    h.emit_end(line);

    // Object / Map → fromEntries([k, normalize(v[k])]) in (optionally sorted)
    // key order.
    lget(&mut h, value_slot, line);
    add_call(&mut h, "ecma:object", "keys", 1, line);
    lset(&mut h, keys_slot, line);
    // sort_keys
    lget(&mut h, sort_slot, line);
    crate::primitives::ops::emit_dyn_to_bool(&mut h, line);
    h.emit_if(line);
    lget(&mut h, keys_slot, line);
    add_call(&mut h, "ecma:array", "sort", 1, line);
    lset(&mut h, keys_slot, line);
    h.emit_end(line);

    h.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    lset(&mut h, out_slot, line);
    lget(&mut h, keys_slot, line);
    h.emit_op(Op::ARRAY_LENGTH, line);
    lset(&mut h, n_slot, line);
    h.emit_f64_const(0.0, line);
    lset(&mut h, i_slot, line);
    let lp = loop_start(&mut h, line);
    loop_break_unless_lt(&mut h, i_slot, n_slot, line);
    lget(&mut h, keys_slot, line);
    lget(&mut h, i_slot, line);
    h.emit_op(Op::ARRAY_GET, line);
    lset(&mut h, key_slot, line);
    // pair = [key, normalize(value[key], ...)]
    lget(&mut h, out_slot, line);
    lget(&mut h, key_slot, line);
    ref_func(&mut h, helper_idx, line);
    lget(&mut h, value_slot, line);
    lget(&mut h, key_slot, line);
    h.emit_op(Op::ARRAY_GET, line);
    lget(&mut h, default_slot, line);
    lget(&mut h, sort_slot, line);
    lget(&mut h, props_slot, line);
    call_ref(&mut h, 4, line);
    h.emit_op_u16(Op::ARRAY_NEW_FIXED, 2, line);
    add_call(&mut h, "ecma:array", "push", 2, line);
    h.emit_op(Op::DROP, line);
    bump(&mut h, i_slot, line);
    loop_end(&mut h, lp, line);
    lget(&mut h, out_slot, line);
    add_call(&mut h, "ecma:object", "fromEntries", 1, line);
    h.emit_op(Op::RETURN, line);

    chunks.push(h);
    helper_idx
}

/// Emit `normalize(value_slot, default_slot, sort_slot, props_slot)`, leaving
/// the JSON-serializable value tree on the stack.
pub fn emit_normalize(
    chunks: &mut Vec<Chunk>,
    current: usize,
    value_slot: u16,
    default_slot: u16,
    sort_slot: u16,
    props_slot: u16,
    line: u32,
) {
    let helper = build_normalize_helper(chunks, line);
    let c = &mut chunks[current];
    ref_func(c, helper, line);
    lget(c, value_slot, line);
    lget(c, default_slot, line);
    lget(c, sort_slot, line);
    lget(c, props_slot, line);
    call_ref(c, 4, line);
}

/// Shared JSON parse that returns `null` instead of throwing on invalid text.
/// Stack: `[json_text] -> [value|null]`.
pub fn emit_parse_or_null(chunks: &mut [Chunk], current: usize, line: u32) {
    add_call(&mut chunks[current], "ecma:json", "parseOrNull", 1, line);
}

/// Shared object/array friendly stringify used by language DOM wrappers.
/// Stack: `[value] -> [json_text]`.
pub fn emit_stringify_props(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let value_slot = chunks[current].alloc_scratch(4);
    let default_slot = value_slot + 1;
    let sort_slot = value_slot + 2;
    let props_slot = value_slot + 3;

    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, default_slot, line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sort_slot, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, props_slot, line);
    emit_normalize(
        chunks,
        current,
        value_slot,
        default_slot,
        sort_slot,
        props_slot,
        line,
    );
    add_call(&mut chunks[current], "ecma:json", "stringify", 1, line);
}

// ── separator-aware renderer ─────────────────────────────────────────────────

/// Build the recursive `__json_render(value, itemSep, kvSep)` helper chunk.
///
/// The value must already be normalized (only arrays / plain objects /
/// primitives). Primitives delegate to `ecma:json.stringify` for correct
/// quoting and number formatting; containers are assembled with the given
/// separators so callers get Python-style `", "` / `": "` compact output that
/// `ecma:json.stringify` cannot itself produce.
fn build_render_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let helper_idx = chunks.len();
    let mut h = create_function_chunk("__json_render", 3);
    h.alloc_scratch(3);

    let value_slot = 0u16;
    let item_slot = 1u16;
    let kv_slot = 2u16;

    let parts_slot = alloc_local(&mut h);
    let keys_slot = alloc_local(&mut h);
    let key_slot = alloc_local(&mut h);
    let i_slot = alloc_local(&mut h);
    let n_slot = alloc_local(&mut h);

    // Array → "[" + parts.join(itemSep) + "]".
    lget(&mut h, value_slot, line);
    add_call(&mut h, "ecma:array", "isArray", 1, line);
    crate::primitives::ops::emit_dyn_to_bool(&mut h, line);
    h.emit_if(line);
    {
        h.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        lset(&mut h, parts_slot, line);
        lget(&mut h, value_slot, line);
        h.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut h, n_slot, line);
        h.emit_f64_const(0.0, line);
        lset(&mut h, i_slot, line);
        let lp = loop_start(&mut h, line);
        loop_break_unless_lt(&mut h, i_slot, n_slot, line);
        lget(&mut h, parts_slot, line);
        ref_func(&mut h, helper_idx, line);
        lget(&mut h, value_slot, line);
        lget(&mut h, i_slot, line);
        h.emit_op(Op::ARRAY_GET, line);
        lget(&mut h, item_slot, line);
        lget(&mut h, kv_slot, line);
        call_ref(&mut h, 3, line);
        add_call(&mut h, "ecma:array", "push", 2, line);
        h.emit_op(Op::DROP, line);
        bump(&mut h, i_slot, line);
        loop_end(&mut h, lp, line);
        push_str(&mut h, "[", line);
        lget(&mut h, parts_slot, line);
        lget(&mut h, item_slot, line);
        add_call(&mut h, "ecma:array", "join", 2, line);
        push_str(&mut h, "]", line);
        crate::primitives::strings::emit_concat(&mut h, 3, line);
        h.emit_op(Op::RETURN, line);
    }
    h.emit_end(line);

    // Object → "{" + parts.join(itemSep) + "}" where each part is
    // stringify(key) + kvSep + render(value[key]).
    push_is_object(&mut h, value_slot, line);
    h.emit_if(line);
    {
        lget(&mut h, value_slot, line);
        add_call(&mut h, "ecma:object", "keys", 1, line);
        lset(&mut h, keys_slot, line);
        h.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        lset(&mut h, parts_slot, line);
        lget(&mut h, keys_slot, line);
        h.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut h, n_slot, line);
        h.emit_f64_const(0.0, line);
        lset(&mut h, i_slot, line);
        let lp = loop_start(&mut h, line);
        loop_break_unless_lt(&mut h, i_slot, n_slot, line);
        lget(&mut h, keys_slot, line);
        lget(&mut h, i_slot, line);
        h.emit_op(Op::ARRAY_GET, line);
        lset(&mut h, key_slot, line);
        lget(&mut h, parts_slot, line);
        // stringify(key)
        lget(&mut h, key_slot, line);
        add_call(&mut h, "ecma:json", "stringify", 1, line);
        // kvSep
        lget(&mut h, kv_slot, line);
        // render(value[key])
        ref_func(&mut h, helper_idx, line);
        lget(&mut h, value_slot, line);
        lget(&mut h, key_slot, line);
        h.emit_op(Op::ARRAY_GET, line);
        lget(&mut h, item_slot, line);
        lget(&mut h, kv_slot, line);
        call_ref(&mut h, 3, line);
        crate::primitives::strings::emit_concat(&mut h, 3, line);
        add_call(&mut h, "ecma:array", "push", 2, line);
        h.emit_op(Op::DROP, line);
        bump(&mut h, i_slot, line);
        loop_end(&mut h, lp, line);
        push_str(&mut h, "{", line);
        lget(&mut h, parts_slot, line);
        lget(&mut h, item_slot, line);
        add_call(&mut h, "ecma:array", "join", 2, line);
        push_str(&mut h, "}", line);
        crate::primitives::strings::emit_concat(&mut h, 3, line);
        h.emit_op(Op::RETURN, line);
    }
    h.emit_end(line);

    // Primitive → ecma:json.stringify.
    lget(&mut h, value_slot, line);
    add_call(&mut h, "ecma:json", "stringify", 1, line);
    h.emit_op(Op::RETURN, line);

    chunks.push(h);
    helper_idx
}

/// Emit `render(value_slot, item_slot, kv_slot)`, leaving the assembled JSON
/// string on the stack. `value_slot` must hold an already-normalized tree.
pub fn emit_render_separated(
    chunks: &mut Vec<Chunk>,
    current: usize,
    value_slot: u16,
    item_slot: u16,
    kv_slot: u16,
    line: u32,
) {
    let helper = build_render_helper(chunks, line);
    let c = &mut chunks[current];
    ref_func(c, helper, line);
    lget(c, value_slot, line);
    lget(c, item_slot, line);
    lget(c, kv_slot, line);
    call_ref(c, 3, line);
}
