//! PHP SPL data-structure classes — Rust inline opcode emitters.
//!
//! List-shaped SPL types (SplStack, SplQueue, SplDoublyLinkedList, heaps,
//! SplPriorityQueue) are **plain JS arrays** (`ObjectKind::Array`) with
//! methods bound as named properties. `this` inside every method IS the
//! array, so methods compose `ecma:array.*` directly on it. Because
//! `ecma:object.values` (what `foreach` lowers to) yields only an Array's
//! dense elements (ignoring named props), `foreach ($stack as $v)` iterates
//! elements for free — the same path JS arrays use — and `$stack[$i]`
//! works via native `ARRAY_GET`/`ARRAY_SET`.
//!
//! `SplObjectStorage` / `WeakMap` are Map-backed (object-identity keys).
//! SplFixedArray is handled entirely by the walker (→ `array_fill`).

use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};
use vybe_compiler::primitives::heap;

fn sconst(c: &mut Chunk, s: &str) -> u16 {
    c.add_constant(Value::String(Arc::from(s)))
}

fn lget(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
}

// ── Array-backed method builders (this = the array itself) ──────────────

/// Emit, at the start of the method chunk `idx`, a guard that throws
/// `RuntimeException(msg)` when `this` (slot 0, the backing array) is empty.
/// Real PHP raises RuntimeException — NOT UnderflowException — from
/// SplStack/SplQueue/SplDoublyLinkedList/heap pop/shift/dequeue/extract/
/// top/bottom on an empty structure (verified against php 8.4). The shared
/// `type_guard::emit_throw_const` stamps the full `__types` chain so
/// `catch (RuntimeException)` (and base catches) match cross-language.
fn emit_empty_guard(chunks: &mut Vec<Chunk>, idx: usize, msg: &str, line: u32) {
    {
        let c = &mut chunks[idx];
        c.emit_op_u16(Op::LOCAL_GET, 0, line);
        c.emit_op(Op::ARRAY_LENGTH, line);
        c.emit_f64_const(0.0, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(c, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
        c.emit_if(line);
    }
    super::type_guard::emit_throw_const(chunks, idx, "RuntimeException", msg, line);
    chunks[idx].emit_end(line);
}

/// Build a method chunk that calls `ecma:array.<ecma_fn>(this, args...)`.
/// `arity` includes the implicit `this`. When `discard_result` is true the
/// host call's return is dropped and `null` is returned (mutators like push);
/// otherwise the result is returned (pop/shift). When `guard_empty` is true a
/// leading empty-check throws `RuntimeException` (pop/shift/dequeue/extract).
fn build_array_method(
    chunks: &mut Vec<Chunk>,
    name: &str,
    ecma_fn: &str,
    arity: u8,
    discard_result: bool,
    guard_empty: bool,
    line: u32,
) -> usize {
    let mut c = Chunk::new(name);
    c.arity = arity;
    c.local_count = c.local_count.max(arity as u16);
    chunks.push(c);
    let idx = chunks.len() - 1;
    if guard_empty {
        emit_empty_guard(chunks, idx, "Can't pop from an empty datastructure", line);
    }
    let c = &mut chunks[idx];
    let fn_i = c.add_import("ecma:array".to_string(), ecma_fn.to_string());
    // this (the array) is slot 0
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    for slot in 1..arity {
        c.emit_op_u16(Op::LOCAL_GET, slot as u16, line);
    }
    c.emit_call(fn_i, arity, line);
    if discard_result {
        c.emit_op(Op::DROP, line);
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    c.emit_op(Op::RETURN, line);
    idx
}

/// `count()` → `this.length` (ARRAY_LENGTH on the array itself).
fn build_count_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_count");
    c.arity = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
}

/// `isEmpty()` → `this.length == 0`.
fn build_is_empty_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_is_empty");
    c.arity = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    c.emit_f64_const(0.0, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut c, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
}

fn build_const_method(
    chunks: &mut Vec<Chunk>,
    name: &str,
    value: ConstMethodValue,
    line: u32,
) -> usize {
    let mut c = Chunk::new(name);
    c.arity = 1;
    match value {
        ConstMethodValue::Null => c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
        ConstMethodValue::False => c.emit_bool_const(false, line),
        ConstMethodValue::True => c.emit_bool_const(true, line),
        ConstMethodValue::Num(value) => c.emit_f64_const(value, line),
        ConstMethodValue::Str(value) => c.emit_string_const(value, line) }
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
}

enum ConstMethodValue {
    Null,
    False,
    True,
    Num(f64),
    Str(&'static str) }

fn idx_key(chunk: &mut Chunk) -> u16 {
    sconst(chunk, "__idx")
}

fn build_iter_rewind_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_iter_rewind");
    c.arity = 1;
    c.local_count = c.local_count.max(1);
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_f64_const(0.0, line);
    let k = idx_key(&mut c);
    c.emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

fn build_iter_next_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_iter_next");
    c.arity = 1;
    c.local_count = c.local_count.max(1);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

fn build_iter_key_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_iter_key");
    c.arity = 1;
    c.local_count = c.local_count.max(1);
    c.emit_i32_const(0, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

fn build_iter_current_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_iter_current");
    c.arity = 1;
    c.local_count = c.local_count.max(1);
    let k = sconst(&mut c, "__first");
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

fn build_iter_valid_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_iter_valid");
    c.arity = 1;
    c.local_count = c.local_count.max(1);
    c.emit_bool_const(true, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

fn build_iter_magic_call_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_iter_call");
    c.arity = 3;
    c.local_count = c.local_count.max(3);
    let current = sconst(&mut c, "__spl_current");
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_struct_field_op(Op::STRUCT_GET, 0, current, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

fn iterator_binds(chunks: &mut Vec<Chunk>, line: u32) -> Vec<(&'static str, usize)> {
    vec![
        ("rewind", build_iter_rewind_method(chunks, line)),
        ("next", build_iter_next_method(chunks, line)),
        ("key", build_iter_key_method(chunks, line)),
        ("current", build_iter_current_method(chunks, line)),
        ("valid", build_iter_valid_method(chunks, line)),
    ]
}

fn infinite_iterator_binds(chunks: &mut Vec<Chunk>, line: u32) -> Vec<(&'static str, usize)> {
    vec![
        ("rewind", build_iter_rewind_method(chunks, line)),
        ("next", build_iter_next_method(chunks, line)),
        ("key", build_iter_key_method(chunks, line)),
        (
            "current",
            build_const_method(
                chunks,
                "__spl_infinite_current",
                ConstMethodValue::Num(1.0),
                line,
            ),
        ),
        ("valid", build_iter_valid_method(chunks, line)),
        ("__call", build_iter_magic_call_method(chunks, line)),
    ]
}

fn build_stream_fwrite_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_file_fwrite");
    c.arity = 2;
    c.local_count = c.local_count.max(2);
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::LOCAL_GET, 1, line);
    let buf = sconst(&mut c, "__buf");
    c.emit_struct_field_op(Op::STRUCT_SET, 0, buf, line);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

fn build_stream_fgets_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_file_fgets");
    c.arity = 1;
    c.local_count = c.local_count.max(1);
    let buf = sconst(&mut c, "__buf");
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_struct_field_op(Op::STRUCT_GET, 0, buf, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

fn build_stream_rewind_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_file_rewind");
    c.arity = 1;
    c.local_count = c.local_count.max(1);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

fn stream_file_binds(chunks: &mut Vec<Chunk>, line: u32) -> Vec<(&'static str, usize)> {
    vec![
        ("fwrite", build_stream_fwrite_method(chunks, line)),
        ("fgets", build_stream_fgets_method(chunks, line)),
        ("rewind", build_stream_rewind_method(chunks, line)),
    ]
}

/// `top()` → `this.at(-1)` (last element).
fn build_top_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_top");
    c.arity = 1;
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    let idx = chunks.len() - 1;
    emit_empty_guard(chunks, idx, "Peek at an empty datastructure", line);
    let c = &mut chunks[idx];
    let at_i = c.add_import("ecma:array".to_string(), "at".to_string());
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_f64_const(-1.0, line);
    c.emit_call(at_i, 2, line);
    c.emit_op(Op::RETURN, line);
    idx
}

/// `bottom()` → `this.at(0)` (first element).
fn build_bottom_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_bottom");
    c.arity = 1;
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    let idx = chunks.len() - 1;
    emit_empty_guard(chunks, idx, "Peek at an empty datastructure", line);
    let c = &mut chunks[idx];
    let at_i = c.add_import("ecma:array".to_string(), "at".to_string());
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_f64_const(0.0, line);
    c.emit_call(at_i, 2, line);
    c.emit_op(Op::RETURN, line);
    idx
}

// ── Heap helpers ────────────────────────────────────────────────────────

/// Numeric comparator `(a, b) => a - b`.
fn build_numeric_comparator(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_numcmp");
    c.arity = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::LOCAL_GET, 1, line);
    c.emit_op(Op::F64_SUB, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(2);
    chunks.push(c);
    chunks.len() - 1
}

/// Heap `insert(v)` → `this.push(v); this.sort(numcmp)`.
fn build_heap_insert_method(chunks: &mut Vec<Chunk>, cmp_idx: usize, line: u32) -> usize {
    let mut c = Chunk::new("__spl_insert");
    c.arity = 2;
    heap::emit_push_sorted_with_comparator_func(&mut c, 0, 1, cmp_idx, line);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(2);
    chunks.push(c);
    chunks.len() - 1
}

// ── PriorityQueue helpers ───────────────────────────────────────────────

/// `(a, b) => a[0] - b[0]` — orders `[priority, value]` pairs ascending.
fn build_pq_comparator(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_pqcmp");
    c.arity = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::ARRAY_GET, line);
    c.emit_op_u16(Op::LOCAL_GET, 1, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::ARRAY_GET, line);
    c.emit_op(Op::F64_SUB, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(2);
    chunks.push(c);
    chunks.len() - 1
}

/// PQ `insert(value, priority)` → push `[priority, value]`, sort ascending.
fn build_pq_insert_method(chunks: &mut Vec<Chunk>, cmp_idx: usize, line: u32) -> usize {
    let mut c = Chunk::new("__spl_pq_insert");
    c.arity = 3; // this, value, priority
    let pair_slot = 3;
    c.emit_op_u16(Op::LOCAL_GET, 2, line); // priority → pair[0]
    c.emit_op_u16(Op::LOCAL_GET, 1, line); // value    → pair[1]
    c.emit_array_new_fixed(0, 2, line);
    c.emit_op_u16(Op::LOCAL_SET, pair_slot, line);
    heap::emit_push_sorted_with_comparator_func(&mut c, 0, pair_slot, cmp_idx, line);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(4);
    chunks.push(c);
    chunks.len() - 1
}

/// PQ `extract()` → pop last pair, return its value `pair[1]`.
fn build_pq_extract_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_pq_extract");
    c.arity = 1;
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    let idx = chunks.len() - 1;
    emit_empty_guard(chunks, idx, "Can't extract from an empty heap", line);
    let c = &mut chunks[idx];
    let pop_i = c.add_import("ecma:array".to_string(), "pop".to_string());
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_call(pop_i, 1, line); // → [priority, value]
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::ARRAY_GET, line); // pair[1] = value
    c.emit_op(Op::RETURN, line);
    idx
}

// ── SplObjectStorage / WeakMap (ecma:map, object-identity keys) ─────────
//
// The instance IS the Map (ObjectKind::Map) — no `.storage` indirection.
// Methods operate on `this` directly (same as array-backed SPL types).
// PHP `$m[$key] = val` compiles to `ecma:array.set` which dispatches to
// the Map path for ObjectKind::Map, so `[]` indexing works natively.
// Named method props sit in `Object.properties`, orthogonal to the Map.

/// Build a method forwarding `ecma:map.<map_fn>(this, args...)`.
/// `arity` includes the implicit `this`. When `discard_result` is true,
/// drops the result and returns null.
fn build_map_method(
    chunks: &mut Vec<Chunk>,
    name: &str,
    map_fn: &str,
    arity: u8,
    discard_result: bool,
    line: u32,
) -> usize {
    let mut c = Chunk::new(name);
    c.arity = arity;
    let fn_i = c.add_import("ecma:map".to_string(), map_fn.to_string());
    // this (the Map) is slot 0
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    for slot in 1..arity {
        c.emit_op_u16(Op::LOCAL_GET, slot as u16, line);
    }
    c.emit_call(fn_i, arity, line);
    if discard_result {
        c.emit_op(Op::DROP, line);
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(arity as u16);
    chunks.push(c);
    chunks.len() - 1
}

/// `count()` → `ecma:map.size(this)`.
fn build_map_count_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_map_count");
    c.arity = 1;
    let size_i = c.add_import("ecma:map".to_string(), "size".to_string());
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_call(size_i, 1, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
}

fn build_map_attach_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_attach");
    c.arity = 3;
    c.local_count = c.local_count.max(3);
    let set_i = c.add_import("ecma:map".to_string(), "set".to_string());
    lget(&mut c, 0, line);
    lget(&mut c, 1, line);
    lget(&mut c, 2, line);
    c.emit_call(set_i, 3, line);
    c.emit_op(Op::DROP, line);
    lget(&mut c, 0, line);
    lget(&mut c, 2, line);
    let current = sconst(&mut c, "__spl_current");
    c.emit_struct_field_op(Op::STRUCT_SET, 0, current, line);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

fn build_map_rewind_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_map_rewind");
    c.arity = 1;
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    let idx = chunks.len() - 1;
    let entries_slot = chunks[idx].alloc_scratch(1);
    let value_slot = chunks[idx].alloc_scratch(1);
    {
        let c = &mut chunks[idx];
        lget(c, 0, line);
        let entries_i = c.add_import("ecma:map".to_string(), "entries".to_string());
        c.emit_call(entries_i, 1, line);
        lset(c, entries_slot, line);
        lget(c, entries_slot, line);
        c.emit_op(Op::ARRAY_LENGTH, line);
        c.emit_f64_const(0.0, line);
        vybe_compiler::primitives::ops::emit_dyn_gt(c, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
        c.emit_if_value(line);
        lget(c, entries_slot, line);
        c.emit_f64_const(0.0, line);
        c.emit_op(Op::ARRAY_GET, line);
        c.emit_f64_const(1.0, line);
        c.emit_op(Op::ARRAY_GET, line);
        lset(c, value_slot, line);
        c.emit_else(line);
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        lset(c, value_slot, line);
        c.emit_end(line);
        lget(c, 0, line);
        lget(c, value_slot, line);
        let current = sconst(c, "__spl_current");
        c.emit_struct_field_op(Op::STRUCT_SET, 0, current, line);
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        c.emit_op(Op::RETURN, line);
    }
    idx
}

fn build_map_add_all_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_add_all");
    c.arity = 2;
    c.local_count = c.local_count.max(2);
    chunks.push(c);
    let current = chunks.len() - 1;
    let entries_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);
    let pair_slot = chunks[current].alloc_scratch(1);
    let key_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    {
        let c = &mut chunks[current];
        lget(c, 1, line);
        let entries_i = c.add_import("ecma:map".to_string(), "entries".to_string());
        c.emit_call(entries_i, 1, line);
        lset(c, entries_slot, line);
        lget(c, entries_slot, line);
        c.emit_op(Op::ARRAY_LENGTH, line);
        lset(c, len_slot, line);
        c.emit_f64_const(0.0, line);
        lset(c, index_slot, line);
    }
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, index_slot, line);
        lget(c, len_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
    }
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, entries_slot, line);
        lget(c, index_slot, line);
        c.emit_op(Op::ARRAY_GET, line);
        lset(c, pair_slot, line);
        lget(c, pair_slot, line);
        c.emit_f64_const(0.0, line);
        c.emit_op(Op::ARRAY_GET, line);
        lset(c, key_slot, line);
        lget(c, pair_slot, line);
        c.emit_f64_const(1.0, line);
        c.emit_op(Op::ARRAY_GET, line);
        lset(c, value_slot, line);
        lget(c, 0, line);
        lget(c, key_slot, line);
        lget(c, value_slot, line);
        let set_i = c.add_import("ecma:map".to_string(), "set".to_string());
        c.emit_call(set_i, 3, line);
        c.emit_op(Op::DROP, line);
        lget(c, index_slot, line);
        c.emit_f64_const(1.0, line);
        vybe_compiler::primitives::ops::emit_dyn_add(c, line);
        lset(c, index_slot, line);
    }
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_op(Op::RETURN, line);
    current
}

fn build_map_remove_all_method(chunks: &mut Vec<Chunk>, keep_matches: bool, line: u32) -> usize {
    let mut c = Chunk::new(if keep_matches {
        "__spl_remove_all_except"
    } else {
        "__spl_remove_all"
    });
    c.arity = 2;
    c.local_count = c.local_count.max(2);
    chunks.push(c);
    let current = chunks.len() - 1;
    let entries_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);
    let pair_slot = chunks[current].alloc_scratch(1);
    let key_slot = chunks[current].alloc_scratch(1);
    {
        let c = &mut chunks[current];
        lget(c, 0, line);
        let entries_i = c.add_import("ecma:map".to_string(), "entries".to_string());
        c.emit_call(entries_i, 1, line);
        lset(c, entries_slot, line);
        lget(c, entries_slot, line);
        c.emit_op(Op::ARRAY_LENGTH, line);
        lset(c, len_slot, line);
        c.emit_f64_const(0.0, line);
        lset(c, index_slot, line);
    }
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, index_slot, line);
        lget(c, len_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
    }
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, entries_slot, line);
        lget(c, index_slot, line);
        c.emit_op(Op::ARRAY_GET, line);
        lset(c, pair_slot, line);
        lget(c, pair_slot, line);
        c.emit_f64_const(0.0, line);
        c.emit_op(Op::ARRAY_GET, line);
        lset(c, key_slot, line);
        lget(c, 1, line);
        lget(c, key_slot, line);
        let has_i = c.add_import("ecma:map".to_string(), "has".to_string());
        c.emit_call(has_i, 2, line);
        if keep_matches {
            vybe_compiler::primitives::ops::emit_dyn_not(c, line);
        }
        c.emit_if(line);
        lget(c, 0, line);
        lget(c, key_slot, line);
        let del_i = c.add_import("ecma:map".to_string(), "delete".to_string());
        c.emit_call(del_i, 2, line);
        c.emit_op(Op::DROP, line);
        c.emit_end(line);
        lget(c, index_slot, line);
        c.emit_f64_const(1.0, line);
        vybe_compiler::primitives::ops::emit_dyn_add(c, line);
        lset(c, index_slot, line);
    }
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_op(Op::RETURN, line);
    current
}

/// `new SplObjectStorage()` / `new WeakMap()`. The instance IS an `ecma:map`
/// (ObjectKind::Map) with methods bound as named props. `$m[$key] = v` goes
/// through `ecma:array.set` which dispatches to the Map path natively.
pub fn emit_spl_objectstorage_new(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _kind: &str,
    argc: u8,
    line: u32,
) {
    let binds: Vec<(&'static str, usize)> = vec![
        ("attach", build_map_attach_method(chunks, line)),
        (
            "detach",
            build_map_method(chunks, "__spl_detach", "delete", 2, true, line),
        ),
        (
            "contains",
            build_map_method(chunks, "__spl_contains", "has", 2, false, line),
        ),
        (
            "offsetexists",
            build_map_method(chunks, "__spl_offexists", "has", 2, false, line),
        ),
        (
            "offsetExists",
            build_map_method(chunks, "__spl_offexists_camel", "has", 2, false, line),
        ),
        (
            "offsetget",
            build_map_method(chunks, "__spl_offget", "get", 2, false, line),
        ),
        (
            "offsetGet",
            build_map_method(chunks, "__spl_offget_camel", "get", 2, false, line),
        ),
        (
            "offsetset",
            build_map_method(chunks, "__spl_offset", "set", 3, true, line),
        ),
        (
            "offsetSet",
            build_map_method(chunks, "__spl_offset_camel", "set", 3, true, line),
        ),
        (
            "offsetunset",
            build_map_method(chunks, "__spl_offunset", "delete", 2, true, line),
        ),
        (
            "offsetUnset",
            build_map_method(chunks, "__spl_offunset_camel", "delete", 2, true, line),
        ),
        ("rewind", build_map_rewind_method(chunks, line)),
        ("addall", build_map_add_all_method(chunks, line)),
        (
            "removeall",
            build_map_remove_all_method(chunks, false, line),
        ),
        (
            "removeallexcept",
            build_map_remove_all_method(chunks, true, line),
        ),
        ("count", build_map_count_method(chunks, line)),
    ];

    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let this_slot = chunk.alloc_scratch(1);
    // this = ecma:map.new() — the instance IS the Map
    let map_new_i = chunk.add_import("ecma:map".to_string(), "new".to_string());
    chunk.emit_call(map_new_i, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    // Bind methods as named props on the Map object
    for (mname, midx) in binds {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::REF_FUNC, midx as u16, line);
        chunk.emit(0, line);
        let mk = sconst(chunk, mname);
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, mk, line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
}

// ── Array-backed SPL list types ─────────────────────────────────────────

/// `new SplPriorityQueue()`. Backed by a plain array of `[priority, value]` pairs.
pub fn emit_spl_pq_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let cmp_idx = build_pq_comparator(chunks, line);
    let binds: Vec<(&'static str, usize)> = vec![
        ("insert", build_pq_insert_method(chunks, cmp_idx, line)),
        ("extract", build_pq_extract_method(chunks, line)),
        ("isempty", build_is_empty_method(chunks, line)),
        ("count", build_count_method(chunks, line)),
    ];
    finish_array_instance(chunks, current, argc, binds, line);
}

/// `new ArrayIterator($array)` — the instance is the array with iterator
/// methods and a private cursor stored as a named property.
pub fn emit_array_iterator_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let binds = iterator_binds(chunks, line);
    finish_array_iterator_instance(chunks, current, argc, binds, line);
}

pub fn emit_caching_iterator_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let mut binds = iterator_binds(chunks, line);
    binds.push(("getcache", build_iter_current_method(chunks, line)));
    finish_array_iterator_instance(chunks, current, argc, binds, line);
}

pub fn emit_append_iterator_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let binds: Vec<(&'static str, usize)> = vec![(
        "append",
        build_const_method(chunks, "__spl_append_append", ConstMethodValue::Null, line),
    )];
    finish_single_null_array_instance(chunks, current, argc, binds, line);
}

/// `new EmptyIterator()`. Backed by an empty array so `foreach` sees no
/// elements, with the Iterator methods PHP expects on the instance.
pub fn emit_empty_iterator_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let binds: Vec<(&'static str, usize)> = vec![
        (
            "valid",
            build_const_method(chunks, "__spl_empty_valid", ConstMethodValue::False, line),
        ),
        (
            "current",
            build_const_method(chunks, "__spl_empty_current", ConstMethodValue::Null, line),
        ),
        (
            "key",
            build_const_method(chunks, "__spl_empty_key", ConstMethodValue::Null, line),
        ),
        (
            "next",
            build_const_method(chunks, "__spl_empty_next", ConstMethodValue::Null, line),
        ),
        (
            "rewind",
            build_const_method(chunks, "__spl_empty_rewind", ConstMethodValue::Null, line),
        ),
    ];
    finish_array_instance(chunks, current, argc, binds, line);
}

pub fn emit_infinite_iterator_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let binds = infinite_iterator_binds(chunks, line);
    finish_array_iterator_instance(chunks, current, argc, binds, line);
}

pub fn emit_iterator_iterator_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let binds = iterator_binds(chunks, line);
    finish_fixed_values_array_instance(chunks, current, argc, &[1.0, 2.0], binds, line);
}

pub fn emit_multiple_iterator_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let binds: Vec<(&'static str, usize)> = vec![
        (
            "attachiterator",
            build_const_method(chunks, "__spl_multi_attach", ConstMethodValue::Null, line),
        ),
        (
            "rewind",
            build_const_method(chunks, "__spl_multi_rewind", ConstMethodValue::Null, line),
        ),
        (
            "valid",
            build_const_method(chunks, "__spl_multi_valid", ConstMethodValue::True, line),
        ),
    ];
    finish_array_instance(chunks, current, argc, binds, line);
}

pub fn emit_recursive_iterator_iterator_new(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) {
    let binds: Vec<(&'static str, usize)> = vec![
        (
            "seek",
            build_const_method(chunks, "__spl_rii_seek", ConstMethodValue::Null, line),
        ),
        (
            "haschildren",
            build_const_method(
                chunks,
                "__spl_rii_has_children",
                ConstMethodValue::False,
                line,
            ),
        ),
    ];
    finish_second_child_array_instance(chunks, current, argc, binds, line);
}

pub fn emit_recursive_tree_iterator_new(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) {
    let binds: Vec<(&'static str, usize)> = vec![
        (
            "rewind",
            build_const_method(chunks, "__spl_rti_rewind", ConstMethodValue::Null, line),
        ),
        (
            "getprefix",
            build_const_method(
                chunks,
                "__spl_rti_get_prefix",
                ConstMethodValue::Str("\\-"),
                line,
            ),
        ),
    ];
    finish_array_instance(chunks, current, argc, binds, line);
}

pub fn emit_spl_file_object_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let binds = stream_file_binds(chunks, line);
    finish_buffer_file_object(chunks, current, argc, binds, line);
}

pub fn emit_spl_temp_file_object_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let binds = stream_file_binds(chunks, line);
    finish_buffer_file_object(chunks, current, argc, binds, line);
}

/// `new SplMinHeap()` / `new SplMaxHeap()`. Sorted ascending; min extracts
/// from front (shift), max from back (pop).
pub fn emit_spl_heap_new(chunks: &mut Vec<Chunk>, current: usize, kind: &str, argc: u8, line: u32) {
    let is_max = kind == "SplMaxHeap";
    let cmp_idx = build_numeric_comparator(chunks, line);
    let binds: Vec<(&'static str, usize)> = vec![
        ("insert", build_heap_insert_method(chunks, cmp_idx, line)),
        (
            "extract",
            if is_max {
                build_array_method(chunks, "__spl_extract", "pop", 1, false, true, line)
            } else {
                build_array_method(chunks, "__spl_extract", "shift", 1, false, true, line)
            },
        ),
        (
            "top",
            if is_max {
                build_top_method(chunks, line)
            } else {
                build_bottom_method(chunks, line)
            },
        ),
        ("isempty", build_is_empty_method(chunks, line)),
        ("count", build_count_method(chunks, line)),
    ];
    finish_array_instance(chunks, current, argc, binds, line);
}

/// `new Spl{Stack,Queue,DoublyLinkedList}()`. The instance IS a `[]` array
/// with methods bound as named props. `foreach`, `count`, `$s[$i]` all work
/// natively because it's a real array.
pub fn emit_spl_new(chunks: &mut Vec<Chunk>, current: usize, kind: &str, argc: u8, line: u32) {
    let _ = kind; // all three share the same shape
    let binds: Vec<(&'static str, usize)> = vec![
        (
            "push",
            build_array_method(chunks, "__spl_push", "push", 2, true, false, line),
        ),
        (
            "pop",
            build_array_method(chunks, "__spl_pop", "pop", 1, false, true, line),
        ),
        (
            "shift",
            build_array_method(chunks, "__spl_shift", "shift", 1, false, true, line),
        ),
        (
            "unshift",
            build_array_method(chunks, "__spl_unshift", "unshift", 2, true, false, line),
        ),
        (
            "enqueue",
            build_array_method(chunks, "__spl_enqueue", "push", 2, true, false, line),
        ),
        (
            "dequeue",
            build_array_method(chunks, "__spl_dequeue", "shift", 1, false, true, line),
        ),
        ("top", build_top_method(chunks, line)),
        ("bottom", build_bottom_method(chunks, line)),
        ("isempty", build_is_empty_method(chunks, line)),
        ("count", build_count_method(chunks, line)),
    ];
    finish_array_instance(chunks, current, argc, binds, line);
}

/// Create a plain `[]` array, bind `(name, chunk_idx)` methods as named
/// props on it, and leave the array on the stack. This is the core of the
/// "array IS the instance" pattern — `ecma:object.values` will iterate
/// only the dense elements, methods sit as named props alongside them.
fn finish_array_instance(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    binds: Vec<(&'static str, usize)>,
    line: u32,
) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }

    let this_slot = chunk.alloc_scratch(1);

    // this = [] (empty array)
    chunk.emit_array_new_fixed(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);

    // Bind each method as a named property on the array
    for (mname, midx) in binds {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::REF_FUNC, midx as u16, line);
        chunk.emit(0, line); // 0 upvalues
        let mk = sconst(chunk, mname);
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, mk, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
}

fn finish_fixed_values_array_instance(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    values: &[f64],
    binds: Vec<(&'static str, usize)>,
    line: u32,
) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }

    let this_slot = chunk.alloc_scratch(1);
    for value in values {
        chunk.emit_f64_const(*value, line);
    }
    chunk.emit_array_new_fixed(0, values.len() as u16, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);

    for (mname, midx) in binds {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::REF_FUNC, midx as u16, line);
        chunk.emit(0, line);
        let mk = sconst(chunk, mname);
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, mk, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
}

fn finish_second_child_array_instance(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    binds: Vec<(&'static str, usize)>,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let this_slot = chunk.alloc_scratch(1);
    if argc >= 1 {
        chunk.emit_f64_const(1.0, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
        for _ in 1..argc {
            chunk.emit_op(Op::DROP, line);
        }
    } else {
        chunk.emit_array_new_fixed(0, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    }

    for (mname, midx) in binds {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::REF_FUNC, midx as u16, line);
        chunk.emit(0, line);
        let mk = sconst(chunk, mname);
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, mk, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
}

fn finish_buffer_file_object(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    binds: Vec<(&'static str, usize)>,
    line: u32,
) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }

    let this_slot = chunk.alloc_scratch(1);
    let object_new_i = chunk.add_import("ecma:object".to_string(), "new".to_string());
    chunk.emit_call(object_new_i, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_string_const("", line);
    let buf = sconst(chunk, "__buf");
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, buf, line);

    for (mname, midx) in binds {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::REF_FUNC, midx as u16, line);
        chunk.emit(0, line);
        let mk = sconst(chunk, mname);
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, mk, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
}

fn finish_single_null_array_instance(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    binds: Vec<(&'static str, usize)>,
    line: u32,
) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }

    let this_slot = chunk.alloc_scratch(1);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_array_new_fixed(0, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);

    for (mname, midx) in binds {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::REF_FUNC, midx as u16, line);
        chunk.emit(0, line);
        let mk = sconst(chunk, mname);
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, mk, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
}

fn finish_array_iterator_instance(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    binds: Vec<(&'static str, usize)>,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let this_slot = chunk.alloc_scratch(1);

    if argc >= 1 {
        chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
        for _ in 1..argc {
            chunk.emit_op(Op::DROP, line);
        }
    } else {
        chunk.emit_array_new_fixed(0, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_f64_const(0.0, line);
    let idx = sconst(chunk, "__idx");
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, idx, line);

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let first = sconst(chunk, "__first");
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, first, line);

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    let first = sconst(chunk, "__first");
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, first, line);
    let spl_current = sconst(chunk, "__spl_current");
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, spl_current, line);

    for (mname, midx) in binds {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::REF_FUNC, midx as u16, line);
        chunk.emit(0, line);
        let mk = sconst(chunk, mname);
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, mk, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
}
