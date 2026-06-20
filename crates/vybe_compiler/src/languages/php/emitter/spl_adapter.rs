//! PHP SPL data-structure classes — Rust inline opcode emitters.
//!
//! `new SplStack()` / `new SplQueue()` / `new SplDoublyLinkedList()` build a
//! plain struct `{ __spl_kind, items: [] }` and BIND their methods on the
//! instance (real function-ref methods, so `$s->count()` dispatches to the
//! object's own method — no name collision with `Countable` etc.). Each
//! method composes existing opcodes + `ecma:array.*`. No host fns, no
//! stdlib, no common/VM changes — same shape as `fiber_adapter.rs`.
//!
//! The walker rewrites `new SplStack(args)` → `__spl_new_splstack(args)`,
//! the profile binds that to `common:php.spl_splstack`, and the dispatcher
//! routes here.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

fn sconst(c: &mut Chunk, s: &str) -> u16 {
    c.add_constant(Value::String(Arc::from(s)))
}

/// Build a method chunk that forwards `this.items.<ecma_fn>(args...)`.
/// `arity` includes the implicit `this`. When `discard_result` is true the
/// host call's value is dropped and `null` is returned (mutators like
/// `push`/`unshift`); otherwise the result is returned (`pop`/`shift`).
fn build_forward_method(
    chunks: &mut Vec<Chunk>,
    name: &str,
    ecma_fn: &str,
    arity: u8,
    discard_result: bool,
    line: u32,
) -> usize {
    let mut c = Chunk::new(name);
    c.arity = arity;
    let items_k = sconst(&mut c, "items");
    let fn_i = c.add_import("ecma:array".to_string(), ecma_fn.to_string());
    // this.items
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, items_k, line);
    // push the user args (slots 1..arity)
    for slot in 1..arity {
        c.emit_op_u16(Op::LOCAL_GET, slot as u16, line);
    }
    c.emit_op_u16(Op::CALL_IMPORT, fn_i, line);
    c.emit(arity, line); // items + (arity-1) args
    if discard_result {
        c.emit_op(Op::DROP, line);
        c.emit_op(Op::NULL, line);
    }
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(arity as u16);
    chunks.push(c);
    chunks.len() - 1
}

/// `count()` → `this.items.length`.
fn build_count_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_count");
    c.arity = 1;
    let items_k = sconst(&mut c, "items");
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, items_k, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
}

/// `isEmpty()` → `this.items.length == 0`.
fn build_is_empty_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_is_empty");
    c.arity = 1;
    let items_k = sconst(&mut c, "items");
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, items_k, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    let zero = c.add_constant(Value::F64(0.0));
    c.emit_op_u16(Op::CONST, zero, line);
    crate::emitter::ops::emit_dyn_eq(&mut c, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
}

/// `top()` → last element `this.items.at(-1)`.
fn build_top_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_top");
    c.arity = 1;
    let items_k = sconst(&mut c, "items");
    let at_i = c.add_import("ecma:array".to_string(), "at".to_string());
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, items_k, line);
    let neg1 = c.add_constant(Value::F64(-1.0));
    c.emit_op_u16(Op::CONST, neg1, line);
    c.emit_op_u16(Op::CALL_IMPORT, at_i, line);
    c.emit(2, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
}

/// `bottom()` → first element `this.items.at(0)`.
fn build_bottom_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_bottom");
    c.arity = 1;
    let items_k = sconst(&mut c, "items");
    let at_i = c.add_import("ecma:array".to_string(), "at".to_string());
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, items_k, line);
    let zero = c.add_constant(Value::F64(0.0));
    c.emit_op_u16(Op::CONST, zero, line);
    c.emit_op_u16(Op::CALL_IMPORT, at_i, line);
    c.emit(2, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
}

/// Numeric comparator chunk `(a, b) => a - b`, for heap ordering.
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

/// Heap `insert(v)` → `this.items.push(v); this.items.sort(numcmp)`.
fn build_heap_insert_method(chunks: &mut Vec<Chunk>, cmp_idx: usize, line: u32) -> usize {
    let mut c = Chunk::new("__spl_insert");
    c.arity = 2;
    let items_k = sconst(&mut c, "items");
    let push_i = c.add_import("ecma:array".to_string(), "push".to_string());
    let sort_i = c.add_import("ecma:array".to_string(), "sort".to_string());
    // this.items.push(v)
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, items_k, line);
    c.emit_op_u16(Op::LOCAL_GET, 1, line);
    c.emit_op_u16(Op::CALL_IMPORT, push_i, line);
    c.emit(2, line);
    c.emit_op(Op::DROP, line);
    // this.items.sort(numcmp)  (ascending)
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, items_k, line);
    c.emit_op_u16(Op::REF_FUNC, cmp_idx as u16, line);
    c.emit(0, line);
    c.emit_op_u16(Op::CALL_IMPORT, sort_i, line);
    c.emit(2, line);
    c.emit_op(Op::DROP, line);
    c.emit_op(Op::NULL, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(2);
    chunks.push(c);
    chunks.len() - 1
}

/// Priority comparator `(a, b) => a[0] - b[0]` — orders `[priority, value]`
/// pairs ascending by priority.
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

/// PriorityQueue `insert(value, priority)` → push `[priority, value]`, sort
/// ascending by priority.
fn build_pq_insert_method(chunks: &mut Vec<Chunk>, cmp_idx: usize, line: u32) -> usize {
    let mut c = Chunk::new("__spl_pq_insert");
    c.arity = 3; // this, value, priority
    let items_k = sconst(&mut c, "items");
    let push_i = c.add_import("ecma:array".to_string(), "push".to_string());
    let sort_i = c.add_import("ecma:array".to_string(), "sort".to_string());
    // this.items.push([priority, value])
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, items_k, line);
    c.emit_op_u16(Op::LOCAL_GET, 2, line); // priority → pair[0]
    c.emit_op_u16(Op::LOCAL_GET, 1, line); // value    → pair[1]
    c.emit_op_u16(Op::ARRAY_NEW_FIXED, 2, line);
    c.emit_op_u16(Op::CALL_IMPORT, push_i, line);
    c.emit(2, line);
    c.emit_op(Op::DROP, line);
    // sort by priority
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, items_k, line);
    c.emit_op_u16(Op::REF_FUNC, cmp_idx as u16, line);
    c.emit(0, line);
    c.emit_op_u16(Op::CALL_IMPORT, sort_i, line);
    c.emit(2, line);
    c.emit_op(Op::DROP, line);
    c.emit_op(Op::NULL, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(3);
    chunks.push(c);
    chunks.len() - 1
}

/// PriorityQueue `extract()` → pop the highest-priority pair, return its value.
fn build_pq_extract_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_pq_extract");
    c.arity = 1;
    let items_k = sconst(&mut c, "items");
    let pop_i = c.add_import("ecma:array".to_string(), "pop".to_string());
    let one = c.add_constant(Value::F64(1.0));
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, items_k, line);
    c.emit_op_u16(Op::CALL_IMPORT, pop_i, line);
    c.emit(1, line); // → [priority, value]
    c.emit_op_u16(Op::CONST, one, line);
    c.emit_op(Op::ARRAY_GET, line); // pair[1] = value
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
}

// ── SplObjectStorage / WeakMap (ecma:map, object-identity keys) ─────────

/// Build a method forwarding `this.storage.<ecma_map_fn>(args...)`. `arity`
/// includes `this`; the map is passed first, then user args. When
/// `discard_result` is set, drops the result and returns null.
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
    let storage_k = sconst(&mut c, "storage");
    let fn_i = c.add_import("ecma:map".to_string(), map_fn.to_string());
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, storage_k, line);
    for slot in 1..arity {
        c.emit_op_u16(Op::LOCAL_GET, slot as u16, line);
    }
    c.emit_op_u16(Op::CALL_IMPORT, fn_i, line);
    c.emit(arity, line);
    if discard_result {
        c.emit_op(Op::DROP, line);
        c.emit_op(Op::NULL, line);
    }
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(arity as u16);
    chunks.push(c);
    chunks.len() - 1
}

/// `count()` → `this.storage.size`.
fn build_map_count_method(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_map_count");
    c.arity = 1;
    let storage_k = sconst(&mut c, "storage");
    let size_i = c.add_import("ecma:map".to_string(), "size".to_string());
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, storage_k, line);
    c.emit_op_u16(Op::CALL_IMPORT, size_i, line);
    c.emit(1, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
}

/// `new SplObjectStorage()` / `new WeakMap()`. Backed by an `ecma:map` whose
/// object keys compare by reference identity. WeakMap also exposes
/// `offsetGet`/`offsetSet` for `$wm[$obj]` ArrayAccess.
pub fn emit_spl_objectstorage_new(
    chunks: &mut Vec<Chunk>,
    current: usize,
    kind: &str,
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
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    chunk.emit_op(Op::DROP, line);
    // this.storage = ecma:map.new()
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    let map_new_i = chunk.add_import("ecma:map".to_string(), "new".to_string());
    chunk.emit_op_u16(Op::CALL_IMPORT, map_new_i, line);
    chunk.emit(0, line);
    let storage_k = sconst(chunk, "storage");
    chunk.emit_op_u16(Op::STRUCT_SET, storage_k, line);
    chunk.emit_op(Op::DROP, line);
    emit_kind_and_binds(chunk, this_slot, kind, binds, line);
}

// ── SplFixedArray (ArrayAccess) ─────────────────────────────────────────

/// `offsetGet($i)` → `this.items[$i]`.
fn build_offset_get(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_offsetget");
    c.arity = 2;
    let items_k = sconst(&mut c, "items");
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, items_k, line);
    c.emit_op_u16(Op::LOCAL_GET, 1, line);
    c.emit_op(Op::ARRAY_GET, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(2);
    chunks.push(c);
    chunks.len() - 1
}

/// `offsetSet($i, $v)` → `this.items[$i] = $v`.
fn build_offset_set(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut c = Chunk::new("__spl_offsetset");
    c.arity = 3;
    let items_k = sconst(&mut c, "items");
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, items_k, line);
    c.emit_op_u16(Op::LOCAL_GET, 1, line);
    c.emit_op_u16(Op::LOCAL_GET, 2, line);
    c.emit_op(Op::ARRAY_SET, line);
    c.emit_op(Op::DROP, line);
    c.emit_op(Op::NULL, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(3);
    chunks.push(c);
    chunks.len() - 1
}

/// `() -> this.<field>` reader (used for `getSize`/`count`/`toArray`).
fn build_field_reader(chunks: &mut Vec<Chunk>, name: &str, field: &str, line: u32) -> usize {
    let mut c = Chunk::new(name);
    c.arity = 1;
    let k = sconst(&mut c, field);
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op_u16(Op::STRUCT_GET, k, line);
    c.emit_op(Op::RETURN, line);
    c.local_count = c.local_count.max(1);
    chunks.push(c);
    chunks.len() - 1
}

fn fixedarray_binds(chunks: &mut Vec<Chunk>, line: u32) -> Vec<(&'static str, usize)> {
    vec![
        ("offsetget", build_offset_get(chunks, line)),
        ("offsetset", build_offset_set(chunks, line)),
        ("getsize", build_field_reader(chunks, "__spl_getsize", "size", line)),
        ("count", build_field_reader(chunks, "__spl_fa_count", "size", line)),
        ("toarray", build_field_reader(chunks, "__spl_toarray", "items", line)),
    ]
}

/// Emit the common tail: stamp `__spl_kind`, bind methods, leave instance.
/// `this_slot` must already hold a struct with `size`/`items` set.
fn emit_kind_and_binds(
    chunk: &mut Chunk,
    this_slot: u16,
    kind: &str,
    binds: Vec<(&'static str, usize)>,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    let kind_c = sconst(chunk, kind);
    chunk.emit_op_u16(Op::CONST, kind_c, line);
    let kind_k = sconst(chunk, "__spl_kind");
    chunk.emit_op_u16(Op::STRUCT_SET, kind_k, line);
    chunk.emit_op(Op::DROP, line);
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

/// `new SplFixedArray($size)`. Stack: `[size]` → `[instance]`.
pub fn emit_spl_fixedarray_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let binds = fixedarray_binds(chunks, line);
    let chunk = &mut chunks[current];
    let size_slot = chunk.local_count;
    let this_slot = chunk.local_count + 1;
    chunk.local_count += 2;
    // capture size (default 0)
    if argc >= 1 {
        chunk.emit_op_u16(Op::LOCAL_SET, size_slot, line);
        chunk.emit_op(Op::DROP, line);
    } else {
        let zero = chunk.add_constant(Value::F64(0.0));
        chunk.emit_op_u16(Op::CONST, zero, line);
        chunk.emit_op_u16(Op::LOCAL_SET, size_slot, line);
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    chunk.emit_op(Op::DROP, line);
    // this.size = size
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, size_slot, line);
    let size_k = sconst(chunk, "size");
    chunk.emit_op_u16(Op::STRUCT_SET, size_k, line);
    chunk.emit_op(Op::DROP, line);
    // this.items = []
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    let items_k = sconst(chunk, "items");
    chunk.emit_op_u16(Op::STRUCT_SET, items_k, line);
    chunk.emit_op(Op::DROP, line);
    emit_kind_and_binds(chunk, this_slot, "SplFixedArray", binds, line);
}

/// `SplFixedArray::fromArray($source)`. Stack: `[source]` → `[instance]`.
pub fn emit_spl_fixedarray_from_array(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    line: u32,
) {
    let binds = fixedarray_binds(chunks, line);
    let chunk = &mut chunks[current];
    let src_slot = chunk.local_count;
    let this_slot = chunk.local_count + 1;
    chunk.local_count += 2;
    chunk.emit_op_u16(Op::LOCAL_SET, src_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    chunk.emit_op(Op::DROP, line);
    // this.items = source
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, src_slot, line);
    let items_k = sconst(chunk, "items");
    chunk.emit_op_u16(Op::STRUCT_SET, items_k, line);
    chunk.emit_op(Op::DROP, line);
    // this.size = source.length
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, src_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    let size_k = sconst(chunk, "size");
    chunk.emit_op_u16(Op::STRUCT_SET, size_k, line);
    chunk.emit_op(Op::DROP, line);
    emit_kind_and_binds(chunk, this_slot, "SplFixedArray", binds, line);
}

/// `new SplPriorityQueue()`.
pub fn emit_spl_pq_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let cmp_idx = build_pq_comparator(chunks, line);
    let binds: Vec<(&'static str, usize)> = vec![
        ("insert", build_pq_insert_method(chunks, cmp_idx, line)),
        ("extract", build_pq_extract_method(chunks, line)),
        ("isempty", build_is_empty_method(chunks, line)),
        ("count", build_count_method(chunks, line)),
    ];
    finish_instance(chunks, current, "SplPriorityQueue", argc, binds, line);
}

/// `new SplMinHeap()` / `new SplMaxHeap()`. Items are kept ascending; the
/// min heap extracts from the front (`shift`), the max heap from the back
/// (`pop`).
pub fn emit_spl_heap_new(chunks: &mut Vec<Chunk>, current: usize, kind: &str, argc: u8, line: u32) {
    let is_max = kind == "SplMaxHeap";
    let cmp_idx = build_numeric_comparator(chunks, line);
    let binds: Vec<(&'static str, usize)> = vec![
        ("insert", build_heap_insert_method(chunks, cmp_idx, line)),
        (
            "extract",
            if is_max {
                build_forward_method(chunks, "__spl_extract", "pop", 1, false, line)
            } else {
                build_forward_method(chunks, "__spl_extract", "shift", 1, false, line)
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
    finish_instance(chunks, current, kind, argc, binds, line);
}

/// `new Spl{Stack,Queue,DoublyLinkedList}()`. Builds the instance struct and
/// binds its methods. Stack on entry: `[ctor args...]` (dropped — these
/// classes take no constructor args). Stack on exit: `[instance]`.
pub fn emit_spl_new(chunks: &mut Vec<Chunk>, current: usize, kind: &str, argc: u8, line: u32) {
    let binds: Vec<(&'static str, usize)> = vec![
        ("push", build_forward_method(chunks, "__spl_push", "push", 2, true, line)),
        ("pop", build_forward_method(chunks, "__spl_pop", "pop", 1, false, line)),
        ("shift", build_forward_method(chunks, "__spl_shift", "shift", 1, false, line)),
        ("unshift", build_forward_method(chunks, "__spl_unshift", "unshift", 2, true, line)),
        ("enqueue", build_forward_method(chunks, "__spl_enqueue", "push", 2, true, line)),
        ("dequeue", build_forward_method(chunks, "__spl_dequeue", "shift", 1, false, line)),
        ("top", build_top_method(chunks, line)),
        ("bottom", build_bottom_method(chunks, line)),
        ("isempty", build_is_empty_method(chunks, line)),
        ("count", build_count_method(chunks, line)),
    ];
    finish_instance(chunks, current, kind, argc, binds, line);
}

/// Emit the instance struct `{ __spl_kind, items: [] }` into `chunks[current]`,
/// bind each `(name, chunk_idx)` method, and leave the instance on the stack.
fn finish_instance(
    chunks: &mut [Chunk],
    current: usize,
    kind: &str,
    argc: u8,
    binds: Vec<(&'static str, usize)>,
    line: u32,
) {
    let chunk = &mut chunks[current];
    // SplStack/Queue/DLL take no constructor args — drop any.
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }

    let this_slot = chunk.local_count;
    chunk.local_count += 1;

    // this = struct_new {}
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    chunk.emit_op(Op::DROP, line);

    // this.__spl_kind = kind
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    let kind_c = sconst(chunk, kind);
    chunk.emit_op_u16(Op::CONST, kind_c, line);
    let kind_k = sconst(chunk, "__spl_kind");
    chunk.emit_op_u16(Op::STRUCT_SET, kind_k, line);
    chunk.emit_op(Op::DROP, line);

    // this.items = []
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    let items_k = sconst(chunk, "items");
    chunk.emit_op_u16(Op::STRUCT_SET, items_k, line);
    chunk.emit_op(Op::DROP, line);

    // Bind each method: this.<name> = ref_func(idx)
    for (mname, midx) in binds {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::REF_FUNC, midx as u16, line);
        chunk.emit(0, line); // 0 upvalues
        let mk = sconst(chunk, mname);
        chunk.emit_op_u16(Op::STRUCT_SET, mk, line);
        chunk.emit_op(Op::DROP, line);
    }

    // Leave the instance as the `new` result.
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
}
