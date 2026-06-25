//! Collection operations — arrays, sets, sorting, range.
//!
//! Every helper that emits a `wasm:js-*` import takes `chunks: &mut [Chunk]`
//! and `current: usize` so imports register on the chunk that emits the
//! bytecode. The VM resolves `CALL_IMPORT` against the executing chunk,
//! and the WASM writer can still aggregate those imports into a module
//! section.

#[allow(unused_imports)]
use crate::emitter::instructions::core_wasm;
use crate::emitter::Target;
use std::sync::Arc;
use vybe_bytecode::Chunk;
use vybe_bytecode::Value;
use vybe_bytecode::opcode::Op;

// ── `ecma:array.*` import helpers (Phase D) ─────────────────
//
// Every language's array surface funnels through these helpers, so the
// emitted .wasm asks for `ecma:array.*` imports whether it runs on
// Vybe's built-in handlers, on v8 (native JS glue), or on plain
// wasmtime with the polyfill module.

fn emit_import_call(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    let c = &mut chunks[current];
    c.emit_call(idx, argc, line);
}

/// Two-chunk variant for callers that have the imports chunk and the
/// code chunk as separate owned objects (notably runtime helper builders, where
/// `build_*` functions build a fresh local Chunk and later append it
/// to the program's chunks vec).
///
/// Same invariant as `emit_import_call`: imports register on the
/// passed `imports` chunk (caller ensures that's the module-level
/// imports chunk = `chunks[0]` of the final program), and the
/// CALL_IMPORT opcode emits in `code`.
pub(crate) fn emit_import_call_into(
    _imports: &mut Chunk,
    code: &mut Chunk,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = code.add_import(module, name);
    code.emit_call(idx, argc, line);
}

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

#[allow(dead_code)]
fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    chunk.emit_string_const(value, line);
}

#[allow(dead_code)]
fn push_f64(chunk: &mut Chunk, value: f64, line: u32) {
    chunk.emit_f64_const(value, line);
}

fn emit_php_array_iter(chunks: &mut [Chunk], current: usize, line: u32, want_entries: bool) {
    // Unified iteration via the host — the single iterator every language
    // shares. `ecma:object.entries` / `iterForOf` dispatch per `ObjectKind`:
    //   - Map    → native IndexMap insertion order
    //   - Array  → index order
    //   - Object → `__keys`-tracked order (handled host-side in
    //              `ordinary_ordered_keys`)
    // The old compiler-side `__keys` / `vybe$assoc_keys_csv` source-selection
    // duplicated exactly this and was the source of the assoc-iteration bugs,
    // so it's retired. Stack: [iterable] → [entries-or-values array].
    emit_import_call(
        chunks,
        current,
        "ecma:object",
        if want_entries { "entries" } else { "iterForOf" },
        1,
        line,
    );
}

/// Create an empty array (common case). Stack: [] → [array] via
/// `vybe:js-array.newWithLength(0)`.
///
/// Non-zero counts still use `ARRAY_NEW_FIXED` because packing N
/// stack values into one array doesn't have a single-op ecma:array
/// equivalent; callers (stdlib/dict [k,v] pair building) migrate
/// incrementally. Each count>0 call site is a Phase E breadcrumb.
pub fn emit_array_new(chunks: &mut [Chunk], current: usize, count: u16, line: u32) {
    if count == 0 {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        emit_import_call(chunks, current, "vybe:js-array", "newWithLength", 1, line);
    } else {
        chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, count, line);
    }
}

/// Create a length-N null-filled array. Stack: [length_i32] → [array]
/// via `vybe:js-array.newWithLength`.
pub fn emit_new_with_length(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "vybe:js-array", "newWithLength", 1, line);
}

/// Length of a collection OR string — runtime-dispatched between
/// `wasm:js-string.length` and `ecma:array.length`.
///
/// `__len__` canonicalises every language's `.length` / `len()` /
/// `.size` / `Length()` into one call; the spec splits arrays
/// (`ecma:array.length`, ECMA-262 §23.1.3.12) from strings
/// (`wasm:js-string.length`, js-string-builtins). A `REF_IS_STRING`
/// branch selects the right import — same pattern v8 uses for
/// property dispatch on auto-boxed primitives.
pub fn emit_len(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], value_slot, line);

    lget(&mut chunks[current], value_slot, line);
    { let idx = chunks[current].add_import("wasm:js-string", "test"); chunks[current].emit_call(idx, 1, line); }
    // wasm:js-string.test already returns I32(0/1) — use it directly as the if
    // condition. Do NOT call emit_dyn_to_bool here: that registers imports on
    // chunks[current] via chunk.add_import, which collides with the global
    // import indices emitted by emit_import_call (chunks[0]-based) below,
    // causing CALL_IMPORT to resolve the wrong host function at runtime.
    chunks[current].emit_if_value(line);

    // String — wasm:js-string.length.
    lget(&mut chunks[current], value_slot, line);
    emit_import_call(chunks, current, "wasm:js-string", "length", 1, line);

    chunks[current].emit_else(line);

    // Not a string — try array length first. A null result means the
    // value wasn't array-like; a numeric 0 is a real empty length and
    // must not fall through to Object.keys(...).length.
    lget(&mut chunks[current], value_slot, line);
    emit_import_call(chunks, current, "ecma:array", "length", 1, line);
    let arr_len_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_len_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_len_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    // REF_IS_NULL returns I32(0/1) — use directly, same reason as above.
    chunks[current].emit_if_value(line);

    lget(&mut chunks[current], value_slot, line);
    emit_import_call(chunks, current, "ecma:map", "size", 1, line);
    let map_len_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map_len_slot, line);

    // map_len != 0: ecma:map.size returns I32 — use I32_NE directly.
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_NE, line);
    chunks[current].emit_if_value(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, map_len_slot, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], value_slot, line);
    emit_import_call(chunks, current, "ecma:object", "keys", 1, line);
    emit_import_call(chunks, current, "ecma:array", "length", 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_len_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// Direct WASM `array.length` — the GC bytecode `0xFB 0x0F` opcode.
/// Stack: [array] → [i32]. Use whenever the operand is statically
/// known to be an array (for-in loops, reduce, polyfills) — this
/// avoids the polymorphic string-or-array dispatch in `emit_len`,
/// which composes flat byte-offset jumps inside structured loops.
pub fn emit_array_length(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::ARRAY_LENGTH, line);
}

/// Array push (spec contract). Stack: [array, value] → [new_length_i32]
/// via `ecma:array.push` — matches ECMA-262 §23.1.3.20.
///
/// Callers that need the array back must stash it in a local before the
/// push loop and reload afterwards. See the `rest_arr` pattern in
/// `compile_function_decl` for the canonical template.
pub fn emit_push(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:array", "push", 2, line);
}

/// Array pop. Stack: [array] → [value] via `ecma:array.pop`.
pub fn emit_pop(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:array", "pop", 1, line);
}

/// Iteration primitives — polymorphic over `Array`, `Map`, and
/// `Ordinary` objects. Every language's iteration (PHP `foreach`,
/// Python `for ... in`, Ruby `each`, JS `for...of`, C# `foreach`)
/// routes through these three functions.
///
/// Provider: `ecma:object` — the Component-Model-portable
/// import set. Modules compiled through these helpers run on any
/// engine that implements `ecma:object` (V8, SpiderMonkey, wasmtime
/// with the polyfill). If the provider ever changes (e.g. inline
/// opcodes for perf), it changes HERE, not in language profiles or walkers.

/// Push an array of keys. Stack: [iterable] → [array_of_keys].
///
/// Uses `iterForIn` so JS `for...in` semantics walk the prototype chain
/// (ECMA-262 §14.7.5.6 step 8.b). `Object.keys(...)` user calls go
/// directly to `ecma:object.keys` (own-only) — that's a separate entry.
pub fn emit_iter_keys(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:object", "iterForIn", 1, line);
}

/// Push an array of values. Stack: [iterable] → [array_of_values].
///
/// `ecma:object.values` yields the entry VALUES uniformly: Array → elements,
/// Map → `m.values()`, Object → property values. This is the correct primitive
/// for value-iteration (`foreach ($a as $v)`, Python `for v in d`). JS `for-of`
/// of a Map (which yields `[k,v]` pairs) pre-drains the Map to a pair-array
/// before reaching here (see the `__vybe_iter_drain` path), so its pairs are
/// preserved as the array's elements — `values` returns them unchanged.
pub fn emit_iter_values(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:object", "values", 1, line);
}

/// Push an array of [key, value] pair arrays. Stack: [iterable] →
/// [array_of_pairs]. Used for `foreach ($m as $k => $v)` in PHP and
/// equivalents in other languages.
pub fn emit_iter_entries(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_php_array_iter(chunks, current, line, true);
}

/// Create an empty Map (ordered associative: IndexMap<Value, Value>).
/// Stack: [] → [map] via `ecma:map.new`. Used by languages with
/// keyed literals (PHP `['k'=>v]`, Python `{'k':v}`, Ruby `{k=>v}`) —
/// same backing across every language, same accessor imports
/// (`ecma:array.get/.set` dispatch polymorphically on
/// `ObjectKind::Map`).
pub fn emit_map_new(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:map", "new", 0, line);
}

/// Polymorphic indexed read. Stack: [collection, index] → [value].
///
/// Emits the WASM GC `array.get` opcode (`Op::ARRAY_GET`, byte 0xFB 0x0B).
/// The VM dispatch routes per `ObjectKind`:
///   - `Array` → indexed Vec access (with negative-index wrap)
///   - `Map`   → IndexMap lookup by Value key (mirrors `ecma:map.get`,
///              ECMA-262 §24.1.3.4)
///   - plain `Object` → property-bag lookup (mirrors `ecma:object.get`)
///   - `Value::String` → char-by-index
///
/// Single bytecode instruction, no host-call overhead. Spec-clean
/// per ObjectKind; the VM is the unified dispatcher. Language-specific
/// wraps (Python negative-index normalization, etc.) layer on top via
/// `Compiler::emit_negative_index_wrap` BEFORE this emit.
pub fn emit_get(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::ARRAY_GET, line);
}

/// Polymorphic indexed write. Stack: [collection, index, value] → [value].
///
/// Emits the WASM GC `array.set` opcode (`Op::ARRAY_SET`, byte 0xFB 0x0E).
/// VM dispatch per ObjectKind: Array element store (with auto-extend
/// + length sync), Map IndexMap insert (mirrors `ecma:map.set`,
/// ECMA-262 §24.1.3.9), or plain Object property-bag write.
///
/// Set element / property. Stack: [obj, key, val] → [retval] via
/// `ecma:array.set`. The host fn is the single dispatch point: Array
/// (with ECMA-262 §6.1.7.2 sparse-fill semantics — holes are
/// Undefined), Map (Value-keyed IndexMap insert), and Ordinary
/// (property bag). Routing through one place keeps PHP `$a[$k]=v`,
/// Python `d[k]=v`, JS `a[i]=v`, Ruby `h[k]=v` etc. on identical
/// runtime semantics regardless of source language.
///
/// The host fn returns `Null` per its spec; existing `emit(Op::DROP)`
/// calls after `emit_set` discard it without breaking the void-style
/// statement form `arr[i] = v;`.
pub fn emit_set(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:array", "set", 3, line);
}

/// Array slice. Stack: [array, start, end] → [array] via `ecma:array.slice`.
/// For polymorphic (string OR array) slicing, prefer the
/// `__vybe_slice` runtime-helper func-ref path.
pub fn emit_slice(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:array", "slice", 3, line);
}

/// Push the __vybe_slice func ref. Use BEFORE compiling the object/start/end.
/// Pure WASM — bundle wires `__vybe_slice` to `build_slice` runtime helper,
/// which dispatches at runtime to either `str_substring` or `array_slice`
/// depending on the operand type. Works uniformly across every language whose
/// surface syntax is `obj[start..end]`.
pub fn emit_slice_push_func(chunk: &mut Chunk, line: u32) {
    let name = chunk.add_constant(Value::String(Arc::from("__vybe_slice")));
    chunk.emit_op_u16(Op::GLOBAL_GET, name, line);
}

/// Invoke __vybe_slice after [func, obj, start, end] are on the stack.
pub fn emit_slice_invoke(chunk: &mut Chunk, line: u32) {
    chunk.emit_op_u8(Op::CALL_REF, 3, line);
}

/// Array join. Stack: [array, delimiter] → [string] via `ecma:array.join`.
pub fn emit_join(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:array", "join", 2, line);
}

/// .NET-style `String.Join(separator, array)` — separator-first arg
/// order. Swaps to the array-first order `ecma:array.join` expects,
/// then dispatches. Also handles the `params object[]` 1-arg form by
/// JS-spec coercion (caller arity check guarantees 2 args by the time
/// we get here). Stack: [separator, array] → [string].
pub fn emit_join_sep_first(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let sep_slot = chunk.alloc_scratch(2);
    let arr_slot = sep_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, sep_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sep_slot, line);
    emit_import_call(chunks, current, "ecma:array", "join", 2, line);
}

/// Array reverse (in-place). Stack: [array] → [array] via `ecma:array.reverse`.
pub fn emit_reverse(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:array", "reverse", 1, line);
}

/// Array contains / JS `.includes`. Stack: [array, value] → [bool] via
/// `ecma:array.includes`.
pub fn emit_contains(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:array", "includes", 2, line);
}

/// Array indexOf. Stack: [array, value] → [i32] via `ecma:array.indexOf`.
pub fn emit_index_of(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:array", "indexOf", 2, line);
}

/// Array concat. Stack: [array, array] → [array] via `ecma:array.concat`.
pub fn emit_concat(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:array", "concat", 2, line);
}

/// Array shift (remove first). Stack: [array] → [value] via `ecma:array.shift`.
pub fn emit_shift(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:array", "shift", 1, line);
}

/// Array fill. Stack: [array, value, start, end] → [array] via `vybe:js-array.fill`.
pub fn emit_fill(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "vybe:js-array", "fill", 4, line);
}

/// Array sort (in-place). Stack: [array] → [array] via the shared
/// `__vybe_sort_in_place` helper so language-level compare semantics
/// stay aligned across collection surfaces.
pub fn emit_sort(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_runtime_helper_call(chunks, current, "__vybe_sort_in_place", 1, line);
}

/// Array lastIndexOf. Stack: [array, value] → [i32] via `ecma:array.lastIndexOf`.
pub fn emit_last_index_of(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:array", "lastIndexOf", 2, line);
}

/// Array removeAt (splice). Stack: [array, index] → [null].
/// splice(arr, index, 1) — deletes 1 element at index.
pub fn emit_remove_at(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::i32_const(&mut chunks[current], line, 1);
    emit_import_call(chunks, current, "ecma:array", "splice", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// Array insert. Stack: [array, index, deleteCount=0, value] → [null].
/// Caller must push 0 as deleteCount before value.
/// splice(arr, index, 0, value) — inserts value at index.
pub fn emit_insert(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:array", "splice", 4, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// indexOf with fromIndex. Stack: [array, value, fromIndex] → [i32].
pub fn emit_index_of_from(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:array", "indexOf", 3, line);
}

/// lastIndexOf with fromIndex. Stack: [array, value, fromIndex] → [i32].
pub fn emit_last_index_of_from(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:array", "lastIndexOf", 3, line);
}

/// RemoveRange. Stack: [array, index, count] → [null].
/// splice(arr, index, count) — removes count elements at index.
pub fn emit_remove_range(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:array", "splice", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// GetRange. Stack: [array, index, count] → [new_array].
/// Computes end = index + count, then slice(arr, index, end).
pub fn emit_get_range(chunks: &mut [Chunk], current: usize, line: u32) {
    let count_local = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    // save count
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_local, line);
    // stack: [arr, index]
    chunks[current].emit_dup(line); // [arr, index, index]
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_local, line); // [arr, index, index, count]
    crate::emitter::ops::emit_dyn_add(&mut chunks[current], line); // [arr, index, end]
    emit_import_call(chunks, current, "ecma:array", "slice", 3, line);
}

/// Clone (full copy). Stack: [array] → [new_array].
/// Pushes 0 and i32::MAX as start/end, then slice.
pub fn emit_clone(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_i32_const(i32::MAX, line);
    emit_import_call(chunks, current, "ecma:array", "slice", 3, line);
}

/// Sequence equality. Stack: [left_array, right_array] -> [bool].
pub fn emit_sequence_equal(chunks: &mut [Chunk], current: usize, line: u32) {
    let left_slot = alloc_local(&mut chunks[current]);
    let right_slot = alloc_local(&mut chunks[current]);
    let len_slot = alloc_local(&mut chunks[current]);
    let idx_slot = alloc_local(&mut chunks[current]);
    let right_elem_slot = alloc_local(&mut chunks[current]);
    let result_slot = alloc_local(&mut chunks[current]);

    lset(&mut chunks[current], right_slot, line);
    lset(&mut chunks[current], left_slot, line);

    chunks[current].emit_bool_const(true, line);
    lset(&mut chunks[current], result_slot, line);

    lget(&mut chunks[current], left_slot, line);
    emit_len(chunks, current, line);
    lset(&mut chunks[current], len_slot, line);

    lget(&mut chunks[current], right_slot, line);
    emit_len(chunks, current, line);
    lget(&mut chunks[current], len_slot, line);
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    lset(&mut chunks[current], idx_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    lget(&mut chunks[current], idx_slot, line);
    lget(&mut chunks[current], len_slot, line);
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    lget(&mut chunks[current], right_slot, line);
    lget(&mut chunks[current], idx_slot, line);
    emit_get(chunks, current, line);
    lset(&mut chunks[current], right_elem_slot, line);

    lget(&mut chunks[current], left_slot, line);
    lget(&mut chunks[current], idx_slot, line);
    emit_get(chunks, current, line);
    lget(&mut chunks[current], right_elem_slot, line);
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    let equal_values = chunks[current].emit_block(line);
    chunks[current].emit_br_if(0, line);

    chunks[current].emit_bool_const(false, line);
    lset(&mut chunks[current], result_slot, line);
    chunks[current].emit_br(2, line);

    chunks[current].emit_end(line);
    chunks[current].patch_block(equal_values);

    lget(&mut chunks[current], idx_slot, line);
    chunks[current].emit_i32_const(1, line);
    crate::emitter::ops::emit_dyn_add(&mut chunks[current], line);
    lset(&mut chunks[current], idx_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    lset(&mut chunks[current], result_slot, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], result_slot, line);
}

/// InsertRange. Stack: [array, index, src_array] → [null].
/// Calls __vybe_array_insert_range stdlib (func below args via local reorder).
pub fn emit_insert_range(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_runtime_helper_call(chunks, current, "__vybe_array_insert_range", 3, line);
}

/// SetRange. Stack: [array, index, src_array] → [null].
pub fn emit_set_range(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_runtime_helper_call(chunks, current, "__vybe_array_set_range", 3, line);
}

/// BinarySearch. Stack: [array, value] → [i32].
pub fn emit_binary_search(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_runtime_helper_call(chunks, current, "__vybe_array_binary_search", 2, line);
}

/// ReverseRange. Stack: [array, index, count] → [null].
pub fn emit_reverse_range(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_runtime_helper_call(chunks, current, "__vybe_array_reverse_range", 3, line);
}

/// Remove by value. Stack: [array, value] → [bool].
pub fn emit_remove_value(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_runtime_helper_call(chunks, current, "__vybe_array_remove_value", 2, line);
}

/// Insert at index. Stack: [array, index, value] → [null].
pub fn emit_insert_at(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_runtime_helper_call(chunks, current, "__vybe_array_insert", 3, line);
}

/// Clear array. Stack: [array] → [null].
/// splice(arr, 0, MAX_INT) removes all elements.
pub fn emit_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_i32_const(i32::MAX, line);
    emit_import_call(chunks, current, "ecma:array", "splice", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// Generic stash-and-call: pop `argc` values into scratch locals,
/// GLOBAL_GET the polyfill func by name, push the args back in order,
/// CALL_REF into the bundled `__vybe_*` runtime helper.
///
/// Slots are appended at `chunks[current].local_count`; locals grow
/// monotonically per call site (no reuse across calls) which trades
/// a few extra Null slots per frame for not requiring Compiler-level
/// scope tracking from this helper.
pub fn emit_runtime_helper_call(
    chunks: &mut [Chunk],
    current: usize,
    global: &'static str,
    argc: u8,
    line: u32,
) {
    let base = chunks[current].local_count;
    chunks[current].alloc_scratch(argc as u16);
    // Stash args (top-of-stack is arg N-1 → highest slot).
    for i in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + i, line);
    }
    let name_c = chunks[current].add_constant(Value::String(Arc::from(global)));
    chunks[current].emit_op_u16(Op::GLOBAL_GET, name_c, line);
    for i in 0..argc as u16 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + i, line);
    }
    chunks[current].emit_op_u8(Op::CALL_REF, argc, line);
}

/// Pack N consecutive stack values into a new array (was the
/// `ARRAY_NEW_FIXED N` opcode). Stack: [v0, v1, …, v(N-1)] → [array].
///
/// There's no single `ecma:array.*` import that consumes N unknown
/// stack values, so this stashes each value into a caller-provided
/// block of consecutive locals, calls `newWithLength(0)`, then pushes
/// each local back in order.
///
/// `slot_base` must be the index of the first of N consecutive caller-
/// allocated local slots (typically via `scope.define()` in the
/// vybex compiler). The caller owns the slots; this helper only
/// reads/writes them.
pub fn emit_pack_n(chunks: &mut [Chunk], current: usize, n: u16, slot_base: u16, line: u32) {
    if n == 0 {
        emit_array_new(chunks, current, 0, line);
        return;
    }
    // Stash in reverse (stack top = v(N-1) goes into slot_base + N-1).
    for i in (0..n).rev() {
        let slot = slot_base + i;
        chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    // Build empty array, push each in forward order.
    emit_array_new(chunks, current, 0, line);
    for i in 0..n {
        chunks[current].emit_dup(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot_base + i, line);
        emit_import_call(chunks, current, "ecma:array", "push", 2, line);
        chunks[current].emit_op(Op::DROP, line); // drop new_length
    }
}

/// Pack two values from stack into a new two-element array.
/// Stack: [v1, v2] → [array_of_two]. Used by dict building etc.
/// See `emit_array_pair_into` for the two-chunk variant.
pub fn emit_array_pair(chunks: &mut [Chunk], current: usize, line: u32) {
    let v2 = chunks[current].local_count;
    let v1 = chunks[current].local_count + 1;
    chunks[current].alloc_scratch(2);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v1, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    emit_import_call(chunks, current, "ecma:array", "newWithLength", 1, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v1, line);
    emit_import_call(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v2, line);
    emit_import_call(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
}

// ── Two-chunk `_into` variants ─────────────────────────────────
//
// For callers that hold the imports chunk and the code chunk as
// separate owned objects — runtime helper builders are the main consumers (their
// `build_*` functions build a fresh local Chunk and return it).
// Each one mirrors the slice-based API above.

/// `ecma:array.newWithLength(0)` → empty Array on `code`'s stack.
/// Import registers on `imports`.
pub fn emit_array_new_into(imports: &mut Chunk, code: &mut Chunk, count: u16, line: u32) {
    if count == 0 {
        core_wasm::i32_const(code, line, 0);
        emit_import_call_into(imports, code, "vybe:js-array", "newWithLength", 1, line);
    } else {
        code.emit_op_u16(Op::ARRAY_NEW_FIXED, count, line);
    }
}

pub fn emit_new_with_length_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "vybe:js-array", "newWithLength", 1, line);
}

pub fn emit_len_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    // Runtime-typed length implemented with structured control flow.
    //
    // WASM blocks with `() -> ()` type can't leave a value on the stack
    // at the end — to hand the computed length out of the dispatch we
    // stash it in a scratch local and reload after the block closes.
    //
    // Stack at entry: [val]. Exit: [len].
    let scratch_val = code.local_count;
    let scratch_len = code.local_count + 1;
    code.local_count = code
        .local_count
        .checked_add(2)
        .expect("emit_len_into: local slot overflow");
    code.emit_op_u16(Op::LOCAL_SET, scratch_val, line);

    let outer = code.emit_block(line);
    let str_block = code.emit_block(line);
    code.emit_op_u16(Op::LOCAL_GET, scratch_val, line);
    { let idx = code.add_import("wasm:js-string", "test"); code.emit_call(idx, 1, line); }
    // wasm:js-string.test already yields i32(0/1); invert with I32_EQZ. Do NOT use
    // emit_dyn_not here — it calls emit_dyn_to_bool, which registers imports
    // on `code` via add_import. Those collide with the chunks[0]-based global
    // import indices used by emit_import_call_into below, making CALL_IMPORT
    // resolve the wrong host fn (ecma:array.length → wasm:js-string.length).
    code.emit_op(Op::I32_EQZ, line);
    // `br_if 0` pops the i32. When `!is_string` is true we jump out of
    // `str_block`, falling through to the array-length branch.
    code.emit_br_if(0, line);
    // String path — stash length and exit outer block.
    code.emit_op_u16(Op::LOCAL_GET, scratch_val, line);
    emit_import_call_into(imports, code, "wasm:js-string", "length", 1, line);
    code.emit_op_u16(Op::LOCAL_SET, scratch_len, line);
    code.emit_br(1, line);
    code.emit_end(line);
    code.patch_block(str_block);
    // Array path — fallthrough from the `br_if 0`.
    code.emit_op_u16(Op::LOCAL_GET, scratch_val, line);
    emit_import_call_into(imports, code, "ecma:array", "length", 1, line);
    code.emit_op_u16(Op::LOCAL_SET, scratch_len, line);
    code.emit_end(line);
    code.patch_block(outer);
    code.emit_op_u16(Op::LOCAL_GET, scratch_len, line);
}

pub fn emit_push_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "ecma:array", "push", 2, line);
}

pub fn emit_pop_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "ecma:array", "pop", 1, line);
}

pub fn emit_get_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "ecma:array", "get", 2, line);
}

pub fn emit_set_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "ecma:array", "set", 3, line);
}

pub fn emit_slice_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "ecma:array", "slice", 3, line);
}

pub fn emit_join_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "ecma:array", "join", 2, line);
}

pub fn emit_reverse_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "ecma:array", "reverse", 1, line);
}

pub fn emit_contains_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "ecma:array", "includes", 2, line);
}

pub fn emit_index_of_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "ecma:array", "indexOf", 2, line);
}

pub fn emit_concat_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "ecma:array", "concat", 2, line);
}

pub fn emit_shift_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "ecma:array", "shift", 1, line);
}

pub fn emit_fill_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "vybe:js-array", "fill", 4, line);
}

/// Pack two values from stack into a new two-element array.
/// Stack: [v1, v2] → [array_of_two]. Used by stdlib for `[k, v]` /
/// `[i, arr[i]]` pair construction — the one pattern without a single
/// `ecma:array.*` equivalent. Allocates 2 scratch slots via
/// `chunk.local_count` (safe in stdlib because these chunks don't
/// share slot space with a scope).
pub fn emit_array_pair_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    let v2 = code.local_count;
    let v1 = code.local_count + 1;
    code.alloc_scratch(2);
    // Stack: [v1, v2] — stash both into temp slots (peek-set + drop).
    code.emit_op_u16(Op::LOCAL_SET, v2, line);
    code.emit_op_u16(Op::LOCAL_SET, v1, line);
    // arr = ecma:array.newWithLength(0)
    core_wasm::i32_const(code, line, 0);
    emit_import_call_into(imports, code, "ecma:array", "newWithLength", 1, line);
    // arr.push(v1)
    code.emit_dup(line);
    code.emit_op_u16(Op::LOCAL_GET, v1, line);
    emit_import_call_into(imports, code, "ecma:array", "push", 2, line);
    code.emit_op(Op::DROP, line);
    // arr.push(v2)
    code.emit_dup(line);
    code.emit_op_u16(Op::LOCAL_GET, v2, line);
    emit_import_call_into(imports, code, "ecma:array", "push", 2, line);
    code.emit_op(Op::DROP, line);
}

// ── Host imports (higher-level operations) ──────────────────

/// range(stop) or range(start, stop) or range(start, stop, step).
/// Stack: [args...] → [array]
///
/// Routes through the bundled `__vybe_range` runtime helper via GLOBAL_GET + CALL_REF.
pub fn emit_range(chunks: &mut [Chunk], current: usize, arg_count: u8, line: u32) {
    emit_runtime_helper_call(chunks, current, "__vybe_range", arg_count, line);
}

/// Target-aware range — inline loop on pure WASM (saves a chunk call),
/// `__vybe_range` polyfill otherwise. Single-arg case stays inlined
/// since it's the most common shape; multi-arg routes to the polyfill.
pub fn emit_range_targeted(
    chunks: &mut [Chunk],
    current: usize,
    arg_count: u8,
    _target: &Target,
    line: u32,
) {
    {
        let chunk = &mut chunks[current];
        if arg_count == 1 {
            let stop_local = chunk.local_count;
            chunk.alloc_scratch(3);
            let i_local = stop_local + 1;
            let result_local = stop_local + 2;

            chunk.emit_op_u16(Op::LOCAL_SET, stop_local, line);
            chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
            chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);
            core_wasm::i32_const(chunk, line, 0);
            chunk.emit_op_u16(Op::LOCAL_SET, i_local, line);

            let block_patch = chunk.emit_block(line);
            let (loop_patch, _) = chunk.emit_loop_s(line);
            chunk.emit_op_u16(Op::LOCAL_GET, i_local, line);
            chunk.emit_op_u16(Op::LOCAL_GET, stop_local, line);
            crate::emitter::ops::emit_dyn_lt(chunk, line);
            crate::emitter::ops::emit_dyn_to_bool(chunk, line);
            chunk.emit_op(Op::I32_EQZ, line);
            chunk.emit_br_if(1, line);

            chunk.emit_op_u16(Op::LOCAL_GET, result_local, line);
            chunk.emit_op_u16(Op::LOCAL_GET, i_local, line);
            let push_idx = chunk.add_import("ecma:array", "push");
            chunk.emit_call(push_idx, 2, line);
            chunk.emit_op(Op::DROP, line);

            chunk.emit_op_u16(Op::LOCAL_GET, i_local, line);
            core_wasm::i32_const(chunk, line, 1);
            chunk.emit_op(Op::I32_ADD, line);
            chunk.emit_op_u16(Op::LOCAL_SET, i_local, line);

            chunk.emit_br(0, line);
            chunk.emit_end(line);
            chunk.patch_loop(loop_patch);
            chunk.emit_end(line);
            chunk.patch_block(block_patch);

            chunk.emit_op_u16(Op::LOCAL_GET, result_local, line);
        } else {
            emit_runtime_helper_call(chunks, current, "__vybe_range", arg_count, line);
        }
    }
}

/// sorted(iterable). Stack: [array] → [sorted_array]
/// Direct call into the bundled `__vybe_sorted` polyfill chunk — same
/// pattern as `emit_sorted_push_func` but consolidated for callers
/// that already have the arg on the stack.
pub fn emit_sorted(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_runtime_helper_call(chunks, current, "__vybe_sorted", 1, line);
}

/// Push the __vybe_sorted func ref. Use BEFORE compiling arg.
pub fn emit_sorted_push_func(chunk: &mut Chunk, line: u32) {
    let name = chunk.add_constant(Value::String(Arc::from("__vybe_sorted")));
    chunk.emit_op_u16(Op::GLOBAL_GET, name, line);
}

/// Invoke __vybe_sorted after [func, arg] are on stack.
pub fn emit_sorted_invoke(chunk: &mut Chunk, line: u32) {
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
}

/// reversed(iterable). Stack: [array] → [reversed_array]
pub fn emit_reversed(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_runtime_helper_call(chunks, current, "__vybe_reversed", 1, line);
}

/// enumerate(iterable). Stack: [array] → [array_of_pairs]
pub fn emit_enumerate(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_runtime_helper_call(chunks, current, "__vybe_enumerate", 1, line);
}

/// zip(a, b). Stack: [a, b] → [pairs]
pub fn emit_zip(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_runtime_helper_call(chunks, current, "__vybe_zip", 2, line);
}

/// sum(array). Stack: [array] → [number]
pub fn emit_sum(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_runtime_helper_call(chunks, current, "__vybe_sum", 1, line);
}

/// Python min(iterable). Stack: [array] → [value]
pub fn emit_pymin(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_runtime_helper_call(chunks, current, "__vybe_min", 1, line);
}

/// Python max(iterable). Stack: [array] → [value]
pub fn emit_pymax(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_runtime_helper_call(chunks, current, "__vybe_max", 1, line);
}
