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
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

fn sconst(c: &mut Chunk, s: &str) -> u16 {
    c.add_constant(Value::String(Arc::from(s)))
}

// ── Array-backed method builders (this = the array itself) ──────────────

/// Build a method chunk that calls `ecma:array.<ecma_fn>(this, args...)`.
/// `arity` includes the implicit `this`. When `discard_result` is true the
/// host call's return is dropped and `null` is returned (mutators like push);
/// otherwise the result is returned (pop/shift).
fn build_array_method(
    chunks: &mut Vec<Chunk>,
    name: &str,
    ecma_fn: &str,
    arity: u8,
    discard_result: bool,
    line: u32,
) -> usize {
    let mut c = Chunk::new(name);
    c.arity = arity;
    let fn_i = c.add_import("ecma:array".to_string(), ecma_fn.to_string());
    // this (the array) is slot 0
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    for slot in 1..arity {
        c.emit_op_u16(Op::LOCAL_GET, slot as u16, line);
    }
    c.emit_call(fn_i, arity, line);
    if discard_result {
        c.emit_op(Op::DROP, line);
        c.emit_op(Op::NULL, line);
    }
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(arity as u16);
    chunks.push(c);
    chunks.len() - 1
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
    let zero = c.add_constant(Value::F64(0.0));
    c.emit_op_u16(Op::CONST, zero, line);
    crate::emitter::ops::emit_dyn_eq(&mut c, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
}

/// `top()` → `this.at(-1)` (last element).
fn build_top_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_top");
    c.arity = 1;
    let at_i = c.add_import("ecma:array".to_string(), "at".to_string());
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    let neg1 = c.add_constant(Value::F64(-1.0));
    c.emit_op_u16(Op::CONST, neg1, line);
    c.emit_call(at_i, 2, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
}

/// `bottom()` → `this.at(0)` (first element).
fn build_bottom_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_bottom");
    c.arity = 1;
    let at_i = c.add_import("ecma:array".to_string(), "at".to_string());
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    let zero = c.add_constant(Value::F64(0.0));
    c.emit_op_u16(Op::CONST, zero, line);
    c.emit_call(at_i, 2, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
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
    let push_i = c.add_import("ecma:array".to_string(), "push".to_string());
    let sort_i = c.add_import("ecma:array".to_string(), "sort".to_string());
    // this.push(v)
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::LOCAL_GET, 1, line);
    c.emit_call(push_i, 2, line);
    c.emit_op(Op::DROP, line);
    // this.sort(numcmp)
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::REF_FUNC, cmp_idx as u16, line);
    c.emit(0, line);
    c.emit_call(sort_i, 2, line);
    c.emit_op(Op::DROP, line);
    c.emit_op(Op::NULL, line);
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
    let zero = c.add_constant(Value::F64(0.0));
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::CONST, zero, line);
    c.emit_op(Op::ARRAY_GET, line);
    c.emit_op_u16(Op::LOCAL_GET, 1, line);
    c.emit_op_u16(Op::CONST, zero, line);
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
    let push_i = c.add_import("ecma:array".to_string(), "push".to_string());
    let sort_i = c.add_import("ecma:array".to_string(), "sort".to_string());
    // this.push([priority, value])
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::LOCAL_GET, 2, line); // priority → pair[0]
    c.emit_op_u16(Op::LOCAL_GET, 1, line); // value    → pair[1]
    c.emit_op_u16(Op::ARRAY_NEW_FIXED, 2, line);
    c.emit_call(push_i, 2, line);
    c.emit_op(Op::DROP, line);
    // this.sort(pqcmp)
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::REF_FUNC, cmp_idx as u16, line);
    c.emit(0, line);
    c.emit_call(sort_i, 2, line);
    c.emit_op(Op::DROP, line);
    c.emit_op(Op::NULL, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(3);
    chunks.push(c);
    chunks.len() - 1
}

/// PQ `extract()` → pop last pair, return its value `pair[1]`.
fn build_pq_extract_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_pq_extract");
    c.arity = 1;
    let pop_i = c.add_import("ecma:array".to_string(), "pop".to_string());
    let one = c.add_constant(Value::F64(1.0));
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_call(pop_i, 1, line); // → [priority, value]
    c.emit_op_u16(Op::CONST, one, line);
    c.emit_op(Op::ARRAY_GET, line); // pair[1] = value
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
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
        c.emit_op(Op::NULL, line);
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
        ("attach", build_map_method(chunks, "__spl_attach", "set", 3, true, line)),
        ("detach", build_map_method(chunks, "__spl_detach", "delete", 2, true, line)),
        ("contains", build_map_method(chunks, "__spl_contains", "has", 2, false, line)),
        ("offsetexists", build_map_method(chunks, "__spl_offexists", "has", 2, false, line)),
        ("offsetget", build_map_method(chunks, "__spl_offget", "get", 2, false, line)),
        ("offsetset", build_map_method(chunks, "__spl_offset", "set", 3, true, line)),
        ("offsetunset", build_map_method(chunks, "__spl_offunset", "delete", 2, true, line)),
        ("count", build_map_count_method(chunks, line)),
    ];

    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let this_slot = chunk.local_count;
    chunk.local_count += 1;
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
        chunk.emit_op_u16(Op::STRUCT_SET, mk, line);
        chunk.emit_op(Op::DROP, line);
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
                build_array_method(chunks, "__spl_extract", "pop", 1, false, line)
            } else {
                build_array_method(chunks, "__spl_extract", "shift", 1, false, line)
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
        ("push", build_array_method(chunks, "__spl_push", "push", 2, true, line)),
        ("pop", build_array_method(chunks, "__spl_pop", "pop", 1, false, line)),
        ("shift", build_array_method(chunks, "__spl_shift", "shift", 1, false, line)),
        ("unshift", build_array_method(chunks, "__spl_unshift", "unshift", 2, true, line)),
        ("enqueue", build_array_method(chunks, "__spl_enqueue", "push", 2, true, line)),
        ("dequeue", build_array_method(chunks, "__spl_dequeue", "shift", 1, false, line)),
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

    let this_slot = chunk.local_count;
    chunk.local_count += 1;

    // this = [] (empty array)
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);

    // Bind each method as a named property on the array
    for (mname, midx) in binds {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::REF_FUNC, midx as u16, line);
        chunk.emit(0, line); // 0 upvalues
        let mk = sconst(chunk, mname);
        chunk.emit_op_u16(Op::STRUCT_SET, mk, line);
        chunk.emit_op(Op::DROP, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
}
