//! Kotlin structural equality — `[builtin_slots.array] eq`.
//!
//! Kotlin's `==` is `equals()`: structural for every collection.
//! `setOf(1, 2) == setOf(2, 1)` is true (order-independent), maps compare
//! per key, lists in order. The platform fallback is reference/primitive
//! equality, so two equal sets compared false.
//!
//! Same shape as Python's `emit_py_value_eq` (the precedent for this slot),
//! but Kotlin's sets and maps are DICTS carrying `__keys` — not ECMA `Set`/
//! `Map` objects — so the legs probe for `__keys` and walk keys with
//! `ecma:object` calls. A set literal's `__keys` also carries the
//! `__kt_set` marker spelling itself, while adapter-built sets attach the
//! marker as a PROP only; the key walk skips it so the two shapes compare
//! equal.

use vybe_compiler::primitives::callable;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const VALUE_EQ_CHUNK: &str = "__kt_value_eq";
const SET_MARKER: &str = crate::emitter::tostring::SET_MARKER;

/// `common:kotlin.value_eq` — [a, b] → [bool i32].
pub fn emit_value_eq(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let idx = ensure_value_eq_chunk(chunks, line);
    let c = &mut chunks[current];
    let b = c.alloc_scratch(1);
    let a = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, b, line);
    c.emit_op_u16(Op::LOCAL_SET, a, line);
    c.emit_op_u16(Op::REF_FUNC, idx as u16, line);
    c.emit(0, line);
    c.emit_op_u16(Op::LOCAL_GET, a, line);
    c.emit_op_u16(Op::LOCAL_GET, b, line);
    callable::emit_direct_invoke_chunk(c, 2, line);
}

fn ensure_value_eq_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    if let Some(idx) = chunks.iter().position(|c| c.name == VALUE_EQ_CHUNK) {
        return idx;
    }
    build_value_eq_chunk(chunks, line)
}

/// Push i32 `1` when `slot` holds a Kotlin dict (own `__keys`).
/// `has`, not `hasOwn`: `hasOwn` HIDES `__`-prefixed keys, so the probe
/// answered false for every dict.
fn emit_is_dict(c: &mut Chunk, slot: u16, line: u32) {
    let typeof_fn = c.add_import("ecma:value", "typeof");
    let has_own = c.add_import("ecma:object", "has");
    let is_array = c.add_import("ecma:array", "isArray");
    let cast_bool = c.add_import("wasm:js-boolean", "cast");
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
    c.emit_call(typeof_fn, 1, line);
    c.emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(c, line);
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
    c.emit_call(is_array, 1, line);
    c.emit_call(cast_bool, 1, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
    c.emit_string_const("__keys", line);
    c.emit_call(has_own, 2, line);
    c.emit_call(cast_bool, 1, line);
    c.emit_op(Op::I32_AND, line);
    // ...and NOT a class instance: instances carry `__types`, and comparing
    // them structurally made two distinct plain objects `==` (Kotlin's
    // default equals is IDENTITY unless the class overrides it).
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
    c.emit_string_const("__types", line);
    c.emit_call(has_own, 2, line);
    c.emit_call(cast_bool, 1, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_op(Op::I32_AND, line);
}

fn emit_is_ecma_set(c: &mut Chunk, slot: u16, line: u32) {
    let tag = c.add_import("ecma:object", "toStringTag");
    let str_eq = c.add_import("wasm:js-string", "equals");
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
    c.emit_call(tag, 1, line);
    c.emit_string_const("[object Set]", line);
    c.emit_call(str_eq, 2, line);
}

fn build_value_eq_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let self_idx = chunks.len();
    let mut c = vybe_compiler::primitives::functions::create_function_chunk(VALUE_EQ_CHUNK, 2);
    c.alloc_scratch(2); // arg slots 0 (a) and 1 (b)
    let a = 0u16;
    let b = 1u16;

    let is_array = c.add_import("ecma:array", "isArray");
    let cast_bool = c.add_import("wasm:js-boolean", "cast");
    let json_str = c.add_import("ecma:json", "stringify");
    let str_eq = c.add_import("wasm:js-string", "equals");

    // ── leg 0: class-declared equality ─────────────────────────────────
    //
    // `data class` equality is derived by normalize_class.rs and published by
    // the shared class primitive under the Eq protocol slot. Sets must honor
    // the same slot as `a == b`, not fall back to object identity. Probe the
    // slot directly rather than keying on `__value_eq`: older constructor
    // shapes did not stamp every structural class path consistently, while
    // the slot itself is the source of truth for operator dispatch.
    let method_slot = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, a, line);
    let slot_key = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(
        vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Eq),
    )));
    c.emit_struct_field_op(Op::STRUCT_GET, 0, slot_key, line);
    c.emit_op_u16(Op::LOCAL_SET, method_slot, line);
    c.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    c.emit_op_u16(Op::LOCAL_GET, a, line);
    c.emit_op_u16(Op::LOCAL_GET, b, line);
    callable::emit_direct_invoke_chunk(&mut c, 2, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);

    // ── leg 1: array vs array — in order, via JSON ──────────────────────
    c.emit_op_u16(Op::LOCAL_GET, a, line);
    c.emit_call(is_array, 1, line);
    c.emit_call(cast_bool, 1, line);
    c.emit_op_u16(Op::LOCAL_GET, b, line);
    c.emit_call(is_array, 1, line);
    c.emit_call(cast_bool, 1, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, a, line);
    c.emit_call(json_str, 1, line);
    c.emit_op_u16(Op::LOCAL_GET, b, line);
    c.emit_call(json_str, 1, line);
    c.emit_call(str_eq, 2, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);

    // ── leg 2: ECMA Set vs ECMA Set — order-independent ────────────────
    emit_is_ecma_set(&mut c, a, line);
    emit_is_ecma_set(&mut c, b, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_if(line);
    emit_ecma_set_eq_body(&mut c, self_idx, a, b, line);
    c.emit_end(line);

    // ── leg 2: dict vs dict — sets AND maps, order-independent ──────────
    emit_is_dict(&mut c, a, line);
    emit_is_dict(&mut c, b, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_if(line);
    emit_dict_eq_body(&mut c, self_idx, a, b, line);
    c.emit_end(line);

    // ── leg 3: identity / primitive ─────────────────────────────────────
    c.emit_op_u16(Op::LOCAL_GET, a, line);
    c.emit_op_u16(Op::LOCAL_GET, b, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut c, line);
    c.emit_op(Op::RETURN, line);

    chunks.push(c);
    self_idx
}

/// ECMA Set equality for Kotlin: same size and every left element has a
/// recursively equal element on the right. Returns from the enclosing function.
fn emit_ecma_set_eq_body(c: &mut Chunk, self_idx: usize, a: u16, b: u16, line: u32) {
    let iter_for_of = c.add_import("ecma:object", "iterForOf");

    let av = c.alloc_scratch(1);
    let bv = c.alloc_scratch(1);
    let n = c.alloc_scratch(1);
    let m = c.alloc_scratch(1);
    let i = c.alloc_scratch(1);
    let j = c.alloc_scratch(1);
    let left_value = c.alloc_scratch(1);
    let found = c.alloc_scratch(1);

    c.emit_op_u16(Op::LOCAL_GET, a, line);
    c.emit_call(iter_for_of, 1, line);
    c.emit_op_u16(Op::LOCAL_SET, av, line);
    c.emit_op_u16(Op::LOCAL_GET, b, line);
    c.emit_call(iter_for_of, 1, line);
    c.emit_op_u16(Op::LOCAL_SET, bv, line);

    c.emit_op_u16(Op::LOCAL_GET, av, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    c.emit_op_u16(Op::LOCAL_SET, n, line);
    c.emit_op_u16(Op::LOCAL_GET, bv, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    c.emit_op_u16(Op::LOCAL_SET, m, line);
    c.emit_op_u16(Op::LOCAL_GET, n, line);
    c.emit_op_u16(Op::LOCAL_GET, m, line);
    c.emit_op(Op::I32_NE, line);
    c.emit_if(line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);

    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, i, line);
    let outer_block = c.emit_block(line);
    let (outer_loop, _) = c.emit_loop_s(line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_op_u16(Op::LOCAL_GET, n, line);
    c.emit_op(Op::I32_GE_S, line);
    c.emit_br_if(1, line);

    c.emit_op_u16(Op::LOCAL_GET, av, line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_op(Op::ARRAY_GET, line);
    c.emit_op_u16(Op::LOCAL_SET, left_value, line);
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, found, line);
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, j, line);

    let inner_block = c.emit_block(line);
    let (inner_loop, _) = c.emit_loop_s(line);
    c.emit_op_u16(Op::LOCAL_GET, j, line);
    c.emit_op_u16(Op::LOCAL_GET, m, line);
    c.emit_op(Op::I32_GE_S, line);
    c.emit_br_if(1, line);

    c.emit_op_u16(Op::REF_FUNC, self_idx as u16, line);
    c.emit(0, line);
    c.emit_op_u16(Op::LOCAL_GET, left_value, line);
    c.emit_op_u16(Op::LOCAL_GET, bv, line);
    c.emit_op_u16(Op::LOCAL_GET, j, line);
    c.emit_op(Op::ARRAY_GET, line);
    callable::emit_direct_invoke_chunk(c, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
    c.emit_if(line);
    c.emit_i32_const(1, line);
    c.emit_op_u16(Op::LOCAL_SET, found, line);
    c.emit_end(line);

    c.emit_op_u16(Op::LOCAL_GET, j, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, j, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(inner_loop);
    c.emit_end(line);
    c.patch_block(inner_block);

    c.emit_op_u16(Op::LOCAL_GET, found, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);

    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, i, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(outer_loop);
    c.emit_end(line);
    c.patch_block(outer_block);

    c.emit_i32_const(1, line);
    c.emit_op(Op::RETURN, line);
}

/// Every key of `a` (marker skipped) must exist in `b` with a recursively
/// equal value, and the marker-free key counts must match. Returns from the
/// enclosing function on every path.
fn emit_dict_eq_body(c: &mut Chunk, self_idx: usize, a: u16, b: u16, line: u32) {
    let has_own = c.add_import("ecma:object", "hasOwn");
    let obj_get = c.add_import("ecma:object", "get");
    let cast_bool = c.add_import("wasm:js-boolean", "cast");

    let keys = c.alloc_scratch(1);
    let n = c.alloc_scratch(1);
    let i = c.alloc_scratch(1);
    let k = c.alloc_scratch(1);
    let count_a = c.alloc_scratch(1);
    let count_b = c.alloc_scratch(1);

    // count_a = 0; walk a's keys: skip marker, check presence + value in b.
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, count_a, line);
    // STRUCT_GET, not `ecma:object.get`: host accessors hide `__` keys.
    c.emit_op_u16(Op::LOCAL_GET, a, line);
    {
        let key = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("__keys")));
        c.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    }
    c.emit_op_u16(Op::LOCAL_SET, keys, line);
    c.emit_op_u16(Op::LOCAL_GET, keys, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    c.emit_op_u16(Op::LOCAL_SET, n, line);
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, i, line);

    let block = c.emit_block(line);
    let (lp, _) = c.emit_loop_s(line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_op_u16(Op::LOCAL_GET, n, line);
    c.emit_op(Op::I32_GE_S, line);
    c.emit_br_if(1, line);

    c.emit_op_u16(Op::LOCAL_GET, keys, line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_op(Op::ARRAY_GET, line);
    c.emit_op_u16(Op::LOCAL_SET, k, line);
    // i += 1 up front so `continue` paths can just branch.
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, i, line);

    // marker key: not an element, skip. `dyn_eq`, NOT js-string.equals —
    // numeric keys are stored as numbers and the string helper traps.
    c.emit_op_u16(Op::LOCAL_GET, k, line);
    c.emit_string_const(SET_MARKER, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(c, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, count_a, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, count_a, line);

    // key missing from b -> false
    c.emit_op_u16(Op::LOCAL_GET, b, line);
    c.emit_op_u16(Op::LOCAL_GET, k, line);
    c.emit_call(has_own, 2, line);
    c.emit_call(cast_bool, 1, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);

    // values differ -> false. RECURSES, so nested collections compare.
    c.emit_op_u16(Op::REF_FUNC, self_idx as u16, line);
    c.emit(0, line);
    c.emit_op_u16(Op::LOCAL_GET, a, line);
    c.emit_op_u16(Op::LOCAL_GET, k, line);
    c.emit_call(obj_get, 2, line);
    c.emit_op_u16(Op::LOCAL_GET, b, line);
    c.emit_op_u16(Op::LOCAL_GET, k, line);
    c.emit_call(obj_get, 2, line);
    callable::emit_direct_invoke_chunk(c, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    c.emit_end(line);

    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(lp);
    c.emit_end(line);
    c.patch_block(block);

    // count_b = b's marker-free key count; sizes must match.
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, count_b, line);
    c.emit_op_u16(Op::LOCAL_GET, b, line);
    {
        let key = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("__keys")));
        c.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    }
    c.emit_op_u16(Op::LOCAL_SET, keys, line);
    c.emit_op_u16(Op::LOCAL_GET, keys, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    c.emit_op_u16(Op::LOCAL_SET, n, line);
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, i, line);

    let block2 = c.emit_block(line);
    let (lp2, _) = c.emit_loop_s(line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_op_u16(Op::LOCAL_GET, n, line);
    c.emit_op(Op::I32_GE_S, line);
    c.emit_br_if(1, line);
    c.emit_op_u16(Op::LOCAL_GET, keys, line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_op(Op::ARRAY_GET, line);
    c.emit_string_const(SET_MARKER, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(c, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, count_b, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, count_b, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, i, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(lp2);
    c.emit_end(line);
    c.patch_block(block2);

    c.emit_op_u16(Op::LOCAL_GET, count_a, line);
    c.emit_op_u16(Op::LOCAL_GET, count_b, line);
    c.emit_op(Op::I32_EQ, line);
    c.emit_op(Op::RETURN, line);
}

/// `common:kotlin.ref_eq` — reference/primitive equality only (typed-array
/// `==`). [a, b] → [bool].
pub fn emit_ref_eq(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}
