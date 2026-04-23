//! Collection operations — arrays, sets, sorting, range.
//!
//! Every helper that emits a `wasm:js-*` import takes `chunks: &mut [Chunk]`
//! and `current: usize` so imports register on `chunks[0]` (the single
//! module-level import section per WASM semantics) while bytecode emits
//! on `chunks[current]`. Helpers that don't need imports still take
//! `&mut Chunk` directly.

use std::sync::Arc;
use vybe_bytecode::Chunk;
use vybe_bytecode::Value;
use vybe_bytecode::opcode::Op;
#[allow(unused_imports)]
use crate::emitter::Target;

// ── `wasm:js-array.*` import helpers (Phase D) ─────────────────
//
// Every language's array surface funnels through these helpers, so the
// emitted .wasm asks for `wasm:js-array.*` imports whether it runs on
// Vybe's built-in handlers, on v8 (native JS glue), or on plain
// wasmtime with the polyfill module.
//
// **WASM import sections are module-level, not per-function.** Vybe
// represents a single user module as many chunks (one per function),
// but the imports section is stored by convention on `chunks[0]`.
// Every helper here adds imports to `chunks[0]` and emits code to
// `chunks[current]` — passing them as `(chunks, current)` gives safe
// disjoint mutable access via array indexing even when `current == 0`.

fn emit_import_call(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[0].add_import(module, name);
    let c = &mut chunks[current];
    c.emit_op_u16(Op::CALL_IMPORT, idx, line);
    c.emit(argc, line);
}

/// Two-chunk variant for callers that have the imports chunk and the
/// code chunk as separate owned objects (notably stdlib.rs, where
/// `build_*` functions build a fresh local Chunk and later append it
/// to the program's chunks vec).
///
/// Same invariant as `emit_import_call`: imports register on the
/// passed `imports` chunk (caller ensures that's the module-level
/// imports chunk = `chunks[0]` of the final program), and the
/// CALL_IMPORT opcode emits in `code`.
fn emit_import_call_into(imports: &mut Chunk, code: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = imports.add_import(module, name);
    code.emit_op_u16(Op::CALL_IMPORT, idx, line);
    code.emit(argc, line);
}

/// Create an empty array (common case). Stack: [] → [array] via
/// `wasm:js-array.newWithLength(0)`.
///
/// Non-zero counts still use `ARRAY_NEW_FIXED` because packing N
/// stack values into one array doesn't have a single-op wasm:js-array
/// equivalent; callers (stdlib/dict [k,v] pair building) migrate
/// incrementally. Each count>0 call site is a Phase E breadcrumb.
pub fn emit_array_new(chunks: &mut [Chunk], current: usize, count: u16, line: u32) {
    if count == 0 {
        chunks[current].emit_op(Op::I32_CONST_0, line);
        emit_import_call(chunks, current, "wasm:js-array", "newWithLength", 1, line);
    } else {
        chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, count, line);
    }
}

/// Create a length-N null-filled array. Stack: [length_i32] → [array]
/// via `wasm:js-array.newWithLength`.
pub fn emit_new_with_length(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "newWithLength", 1, line);
}

/// Length of a collection OR string — runtime-dispatched between
/// `wasm:js-string.length` and `wasm:js-array.length`.
///
/// `__len__` canonicalises every language's `.length` / `len()` /
/// `.size` / `Length()` into one call; the spec splits arrays
/// (`wasm:js-array.length`, ECMA-262 §23.1.3.12) from strings
/// (`wasm:js-string.length`, js-string-builtins). A `REF_IS_STRING`
/// branch selects the right import — same pattern v8 uses for
/// property dispatch on auto-boxed primitives.
pub fn emit_len(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    c.emit_op(Op::DUP, line);                     // [v, v]
    c.emit_op(Op::REF_IS_STRING, line);            // [v, is_string]
    let to_str = c.emit_jump(Op::BR_IF_TRUE, line); // consumes bool
    // Not a string — wasm:js-array.length.
    emit_import_call(chunks, current, "wasm:js-array", "length", 1, line);
    let end = chunks[current].emit_jump(Op::BR, line);
    chunks[current].patch_jump(to_str);
    // String — wasm:js-string.length.
    emit_import_call(chunks, current, "wasm:js-string", "length", 1, line);
    chunks[current].patch_jump(end);
}

/// Array push (spec contract). Stack: [array, value] → [new_length_i32]
/// via `wasm:js-array.push` — matches ECMA-262 §23.1.3.20.
///
/// Callers that need the array back must stash it in a local before the
/// push loop and reload afterwards. See the `rest_arr` pattern in
/// `compile_function_decl` for the canonical template.
pub fn emit_push(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "push", 2, line);
}

/// Array pop. Stack: [array] → [value] via `wasm:js-array.pop`.
pub fn emit_pop(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "pop", 1, line);
}

/// Iteration primitives — polymorphic over `Array`, `Map`, and
/// `Ordinary` objects. Every language's iteration (PHP `foreach`,
/// Python `for ... in`, Ruby `each`, JS `for...of`, C# `foreach`)
/// routes through these three functions. Swapping the host provider
/// (e.g. `vybe:object` → a different iter surface) happens here.

/// Push an array of keys. Stack: [iterable] → [array_of_keys].
pub fn emit_iter_keys(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "vybe:object", "keys", 1, line);
}

/// Push an array of values. Stack: [iterable] → [array_of_values].
pub fn emit_iter_values(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "vybe:object", "values", 1, line);
}

/// Push an array of [key, value] pair arrays. Stack: [iterable] →
/// [array_of_pairs]. Used for `foreach ($m as $k => $v)` in PHP and
/// equivalents in other languages.
pub fn emit_iter_entries(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "vybe:object", "entries", 1, line);
}

/// Create an empty Map (ordered associative: IndexMap<Value, Value>).
/// Stack: [] → [map] via `wasm:js-map.new`. Used by languages with
/// keyed literals (PHP `['k'=>v]`, Python `{'k':v}`, Ruby `{k=>v}`) —
/// same backing across every language, same accessor imports
/// (`wasm:js-array.get/.set` dispatch polymorphically on
/// `ObjectKind::Map`).
pub fn emit_map_new(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-map", "new", 0, line);
}

/// Array get. Stack: [array, index] → [value] via `wasm:js-array.get`.
pub fn emit_get(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "get", 2, line);
}

/// Array set (spec contract). Stack: [array, index, value] → [null]
/// via `wasm:js-array.set` — the import is void (mutates in place).
/// Callers that need the assigned value back must DUP it before
/// emit_set and DROP the returned null.
pub fn emit_set(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "set", 3, line);
}

/// Array slice. Stack: [array, start, end] → [array] via `wasm:js-array.slice`.
/// For polymorphic (string OR array) slicing, prefer the
/// `__vybe_slice` stdlib func-ref path.
pub fn emit_slice(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "slice", 3, line);
}

/// Push the __vybe_slice func ref. Use BEFORE compiling the object/start/end.
/// Pure WASM — bundle wires `__vybe_slice` to `build_slice` stdlib chunk,
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

/// Array join. Stack: [array, delimiter] → [string] via `wasm:js-array.join`.
pub fn emit_join(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "join", 2, line);
}

/// Array reverse (in-place). Stack: [array] → [array] via `wasm:js-array.reverse`.
pub fn emit_reverse(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "reverse", 1, line);
}

/// Array contains / JS `.includes`. Stack: [array, value] → [bool] via
/// `wasm:js-array.includes`.
pub fn emit_contains(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "includes", 2, line);
}

/// Array indexOf. Stack: [array, value] → [i32] via `wasm:js-array.indexOf`.
pub fn emit_index_of(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "indexOf", 2, line);
}

/// Array concat. Stack: [array, array] → [array] via `wasm:js-array.concat`.
pub fn emit_concat(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "concat", 2, line);
}

/// Array shift (remove first). Stack: [array] → [value] via `wasm:js-array.shift`.
pub fn emit_shift(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "shift", 1, line);
}

/// Array fill. Stack: [array, value, start, end] → [array] via `wasm:js-array.fill`.
pub fn emit_fill(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "fill", 4, line);
}

/// Array sort (in-place). Stack: [array] → [array] via `wasm:js-array.sort`.
pub fn emit_sort(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "sort", 1, line);
}

/// Array lastIndexOf. Stack: [array, value] → [i32] via `wasm:js-array.lastIndexOf`.
pub fn emit_last_index_of(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "lastIndexOf", 2, line);
}

/// Array removeAt (splice). Stack: [array, index] → [null].
/// splice(arr, index, 1) — deletes 1 element at index.
pub fn emit_remove_at(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::I32_CONST_1, line);
    emit_import_call(chunks, current, "wasm:js-array", "splice", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// Array insert. Stack: [array, index, deleteCount=0, value] → [null].
/// Caller must push 0 as deleteCount before value.
/// splice(arr, index, 0, value) — inserts value at index.
pub fn emit_insert(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "splice", 4, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// indexOf with fromIndex. Stack: [array, value, fromIndex] → [i32].
pub fn emit_index_of_from(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "indexOf", 3, line);
}

/// lastIndexOf with fromIndex. Stack: [array, value, fromIndex] → [i32].
pub fn emit_last_index_of_from(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "lastIndexOf", 3, line);
}

/// RemoveRange. Stack: [array, index, count] → [null].
/// splice(arr, index, count) — removes count elements at index.
pub fn emit_remove_range(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "wasm:js-array", "splice", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// GetRange. Stack: [array, index, count] → [new_array].
/// Computes end = index + count, then slice(arr, index, end).
pub fn emit_get_range(chunks: &mut [Chunk], current: usize, line: u32) {
    let count_local = chunks[current].local_count;
    chunks[current].local_count += 1;
    // save count
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_local, line);
    chunks[current].emit_op(Op::DROP, line);
    // stack: [arr, index]
    chunks[current].emit_op(Op::DUP, line); // [arr, index, index]
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_local, line); // [arr, index, index, count]
    chunks[current].emit_op(Op::DYN_ADD, line); // [arr, index, end]
    emit_import_call(chunks, current, "wasm:js-array", "slice", 3, line);
}

/// Clone (full copy). Stack: [array] → [new_array].
/// Pushes 0 and i32::MAX as start/end, then slice.
pub fn emit_clone(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::I32_CONST_0, line);
    let max_c = chunks[current].add_constant(Value::I32(i32::MAX));
    chunks[current].emit_op_u16(Op::CONST, max_c, line);
    emit_import_call(chunks, current, "wasm:js-array", "slice", 3, line);
}

/// InsertRange. Stack: [array, index, src_array] → [null].
/// Calls __vybe_array_insert_range stdlib (func below args via local reorder).
pub fn emit_insert_range(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_stdlib_call_3(chunks, current, "__vybe_array_insert_range", line);
}

/// SetRange. Stack: [array, index, src_array] → [null].
pub fn emit_set_range(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_stdlib_call_3(chunks, current, "__vybe_array_set_range", line);
}

/// BinarySearch. Stack: [array, value] → [i32].
pub fn emit_binary_search(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_stdlib_call_2(chunks, current, "__vybe_array_binary_search", line);
}

/// ReverseRange. Stack: [array, index, count] → [null].
pub fn emit_reverse_range(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_stdlib_call_3(chunks, current, "__vybe_array_reverse_range", line);
}

/// Remove by value. Stack: [array, value] → [bool].
pub fn emit_remove_value(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_stdlib_call_2(chunks, current, "__vybe_array_remove_value", line);
}

/// Insert at index. Stack: [array, index, value] → [null].
pub fn emit_insert_at(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_stdlib_call_3(chunks, current, "__vybe_array_insert", line);
}

/// Clear array. Stack: [array] → [null].
/// splice(arr, 0, MAX_INT) removes all elements.
pub fn emit_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::I32_CONST_0, line);
    let max_c = chunks[current].add_constant(Value::I32(i32::MAX));
    chunks[current].emit_op_u16(Op::CONST, max_c, line);
    emit_import_call(chunks, current, "wasm:js-array", "splice", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// Helper: stash 3 args to locals, GLOBAL_GET func, restore args, CALL_REF 3.
fn emit_stdlib_call_3(chunks: &mut [Chunk], current: usize, global: &'static str, line: u32) {
    let a2 = chunks[current].local_count;
    let a1 = chunks[current].local_count + 1;
    let a0 = chunks[current].local_count + 2;
    chunks[current].local_count += 3;
    chunks[current].emit_op_u16(Op::LOCAL_SET, a2, line); chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a1, line); chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a0, line); chunks[current].emit_op(Op::DROP, line);
    let name_c = chunks[current].add_constant(Value::String(Arc::from(global)));
    chunks[current].emit_op_u16(Op::GLOBAL_GET, name_c, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a2, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 3, line);
}

/// Helper: stash 2 args to locals, GLOBAL_GET func, restore args, CALL_REF 2.
fn emit_stdlib_call_2(chunks: &mut [Chunk], current: usize, global: &'static str, line: u32) {
    let a1 = chunks[current].local_count;
    let a0 = chunks[current].local_count + 1;
    chunks[current].local_count += 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, a1, line); chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a0, line); chunks[current].emit_op(Op::DROP, line);
    let name_c = chunks[current].add_constant(Value::String(Arc::from(global)));
    chunks[current].emit_op_u16(Op::GLOBAL_GET, name_c, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a1, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
}

/// Pack N consecutive stack values into a new array (was the
/// `ARRAY_NEW_FIXED N` opcode). Stack: [v0, v1, …, v(N-1)] → [array].
///
/// There's no single `wasm:js-array.*` import that consumes N unknown
/// stack values, so this stashes each value into a caller-provided
/// block of consecutive locals, calls `newWithLength(0)`, then pushes
/// each local back in order.
///
/// `slot_base` must be the index of the first of N consecutive caller-
/// allocated local slots (typically via `scope.define()` in the
/// vybex compiler). The caller owns the slots; this helper only
/// reads/writes them.
pub fn emit_pack_n(
    chunks: &mut [Chunk],
    current: usize,
    n: u16,
    slot_base: u16,
    line: u32,
) {
    if n == 0 {
        emit_array_new(chunks, current, 0, line);
        return;
    }
    // Stash in reverse (stack top = v(N-1) goes into slot_base + N-1).
    for i in (0..n).rev() {
        let slot = slot_base + i;
        chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    // Build empty array, push each in forward order.
    emit_array_new(chunks, current, 0, line);
    for i in 0..n {
        chunks[current].emit_op(Op::DUP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot_base + i, line);
        emit_import_call(chunks, current, "wasm:js-array", "push", 2, line);
        chunks[current].emit_op(Op::DROP, line); // drop new_length
    }
}

/// Pack two values from stack into a new two-element array.
/// Stack: [v1, v2] → [array_of_two]. Used by dict building etc.
/// See `emit_array_pair_into` for the two-chunk variant.
pub fn emit_array_pair(chunks: &mut [Chunk], current: usize, line: u32) {
    let v2 = chunks[current].local_count;
    let v1 = chunks[current].local_count + 1;
    chunks[current].local_count += 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, v2, line); chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v1, line); chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    emit_import_call(chunks, current, "wasm:js-array", "newWithLength", 1, line);
    chunks[current].emit_op(Op::DUP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v1, line);
    emit_import_call(chunks, current, "wasm:js-array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::DUP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v2, line);
    emit_import_call(chunks, current, "wasm:js-array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
}

// ── Two-chunk `_into` variants ─────────────────────────────────
//
// For callers that hold the imports chunk and the code chunk as
// separate owned objects — stdlib.rs is the main consumer (its
// `build_*` functions build a fresh local Chunk and return it).
// Each one mirrors the slice-based API above.

/// `wasm:js-array.newWithLength(0)` → empty Array on `code`'s stack.
/// Import registers on `imports`.
pub fn emit_array_new_into(imports: &mut Chunk, code: &mut Chunk, count: u16, line: u32) {
    if count == 0 {
        code.emit_op(Op::I32_CONST_0, line);
        emit_import_call_into(imports, code, "wasm:js-array", "newWithLength", 1, line);
    } else {
        code.emit_op_u16(Op::ARRAY_NEW_FIXED, count, line);
    }
}

pub fn emit_new_with_length_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "wasm:js-array", "newWithLength", 1, line);
}

pub fn emit_len_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    // Runtime-typed length. `emit_jump(Op::BR_IF_TRUE)` uses flat
    // offsets that the wasm emitter writes as `nop`, so the jump would
    // vanish on v8. Use structured control flow instead.
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
    code.emit_op(Op::DROP, line);

    let outer = code.emit_block(line);
    let str_block = code.emit_block(line);
    code.emit_op_u16(Op::LOCAL_GET, scratch_val, line);
    code.emit_op(Op::REF_IS_STRING, line);
    code.emit_op(Op::DYN_NOT, line);
    // `br_if 0` pops the bool. When `!is_string` is true we jump out of
    // `str_block`, falling through to the array-length branch.
    code.emit_br_if(0, line);
    // String path — stash length and exit outer block.
    code.emit_op_u16(Op::LOCAL_GET, scratch_val, line);
    emit_import_call_into(imports, code, "wasm:js-string", "length", 1, line);
    code.emit_op_u16(Op::LOCAL_SET, scratch_len, line);
    code.emit_op(Op::DROP, line);
    code.emit_br(1, line);
    code.emit_end(line);
    code.patch_block(str_block);
    // Array path — fallthrough from the `br_if 0`.
    code.emit_op_u16(Op::LOCAL_GET, scratch_val, line);
    emit_import_call_into(imports, code, "wasm:js-array", "length", 1, line);
    code.emit_op_u16(Op::LOCAL_SET, scratch_len, line);
    code.emit_op(Op::DROP, line);
    code.emit_end(line);
    code.patch_block(outer);
    code.emit_op_u16(Op::LOCAL_GET, scratch_len, line);
}

pub fn emit_push_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "wasm:js-array", "push", 2, line);
}

pub fn emit_pop_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "wasm:js-array", "pop", 1, line);
}

pub fn emit_get_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "wasm:js-array", "get", 2, line);
}

pub fn emit_set_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "wasm:js-array", "set", 3, line);
}

pub fn emit_slice_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "wasm:js-array", "slice", 3, line);
}

pub fn emit_join_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "wasm:js-array", "join", 2, line);
}

pub fn emit_reverse_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "wasm:js-array", "reverse", 1, line);
}

pub fn emit_contains_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "wasm:js-array", "includes", 2, line);
}

pub fn emit_index_of_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "wasm:js-array", "indexOf", 2, line);
}

pub fn emit_concat_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "wasm:js-array", "concat", 2, line);
}

pub fn emit_shift_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "wasm:js-array", "shift", 1, line);
}

pub fn emit_fill_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_import_call_into(imports, code, "wasm:js-array", "fill", 4, line);
}

/// Pack two values from stack into a new two-element array.
/// Stack: [v1, v2] → [array_of_two]. Used by stdlib for `[k, v]` /
/// `[i, arr[i]]` pair construction — the one pattern without a single
/// `wasm:js-array.*` equivalent. Allocates 2 scratch slots via
/// `chunk.local_count` (safe in stdlib because these chunks don't
/// share slot space with a scope).
pub fn emit_array_pair_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    let v2 = code.local_count;
    let v1 = code.local_count + 1;
    code.local_count += 2;
    // Stack: [v1, v2] — stash both into temp slots (peek-set + drop).
    code.emit_op_u16(Op::LOCAL_SET, v2, line); code.emit_op(Op::DROP, line);
    code.emit_op_u16(Op::LOCAL_SET, v1, line); code.emit_op(Op::DROP, line);
    // arr = wasm:js-array.newWithLength(0)
    code.emit_op(Op::I32_CONST_0, line);
    emit_import_call_into(imports, code, "wasm:js-array", "newWithLength", 1, line);
    // arr.push(v1)
    code.emit_op(Op::DUP, line);
    code.emit_op_u16(Op::LOCAL_GET, v1, line);
    emit_import_call_into(imports, code, "wasm:js-array", "push", 2, line);
    code.emit_op(Op::DROP, line);
    // arr.push(v2)
    code.emit_op(Op::DUP, line);
    code.emit_op_u16(Op::LOCAL_GET, v2, line);
    emit_import_call_into(imports, code, "wasm:js-array", "push", 2, line);
    code.emit_op(Op::DROP, line);
}

// ── Host imports (higher-level operations) ──────────────────

/// range(stop) or range(start, stop) or range(start, stop, step).
/// Stack: [args...] → [array]
///
/// On Vybe: single host call. On standard WASM: inline loop.
pub fn emit_range(chunks: &mut [Chunk], current: usize, arg_count: u8, line: u32) {
    let idx = chunks[0].add_import("vybe:array", "range");
    let c = &mut chunks[current];
    c.emit_op_u16(Op::CALL_IMPORT, idx, line);
    c.emit(arg_count, line);
}

/// Target-aware range — uses host call on Vybe, inline loop on pure WASM.
/// Stack: [start, stop] → [array]
pub fn emit_range_targeted(chunks: &mut [Chunk], current: usize, arg_count: u8, target: &Target, line: u32) {
    if target.has_module("vybe:array") {
        let idx = chunks[0].add_import("vybe:array", "range");
        let c = &mut chunks[current];
        c.emit_op_u16(Op::CALL_IMPORT, idx, line);
        c.emit(arg_count, line);
    } else {
        let chunk = &mut chunks[current];
        if arg_count == 1 {
            let stop_local = chunk.local_count;
            chunk.local_count += 3;
            let i_local = stop_local + 1;
            let result_local = stop_local + 2;

            chunk.emit_op_u16(Op::LOCAL_SET, stop_local, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
            chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op(Op::I32_CONST_0, line);
            chunk.emit_op_u16(Op::LOCAL_SET, i_local, line);
            chunk.emit_op(Op::DROP, line);

            let loop_start = chunk.current_offset();
            chunk.emit_op_u16(Op::LOCAL_GET, i_local, line);
            chunk.emit_op_u16(Op::LOCAL_GET, stop_local, line);
            chunk.emit_op(Op::DYN_LT, line);
            let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

            chunk.emit_op_u16(Op::LOCAL_GET, result_local, line);
            chunk.emit_op_u16(Op::LOCAL_GET, i_local, line);
            let push_idx = chunk.add_import("wasm:js-array", "push");
            chunk.emit_op_u16(Op::CALL_IMPORT, push_idx, line);
            chunk.emit(2u8, line);
            chunk.emit_op(Op::DROP, line);

            chunk.emit_op_u16(Op::LOCAL_GET, i_local, line);
            chunk.emit_op(Op::I32_CONST_1, line);
            chunk.emit_op(Op::I32_ADD, line);
            chunk.emit_op_u16(Op::LOCAL_SET, i_local, line);
            chunk.emit_op(Op::DROP, line);

            chunk.emit_loop(loop_start, line);
            chunk.patch_jump(exit);

            chunk.emit_op_u16(Op::LOCAL_GET, result_local, line);
        } else {
            let idx = chunks[0].add_import("vybe:array", "range");
            let c = &mut chunks[current];
            c.emit_op_u16(Op::CALL_IMPORT, idx, line);
            c.emit(arg_count, line);
        }
    }
}

/// sorted(iterable). Stack: [array] → [sorted_array]
/// Legacy entry point — uses host import. The bundle aliases vybe:array sorted to __vybe_sorted.
/// Prefer using `emit_sorted_push_func` + args + `emit_sorted_invoke` for pure WASM bytecode.
pub fn emit_sorted(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[0].add_import("vybe:array", "sorted");
    let c = &mut chunks[current];
    c.emit_op_u16(Op::CALL_IMPORT, idx, line);
    c.emit(1, line);
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
    let idx = chunks[0].add_import("vybe:array", "reversed");
    let c = &mut chunks[current];
    c.emit_op_u16(Op::CALL_IMPORT, idx, line);
    c.emit(1, line);
}

/// enumerate(iterable). Stack: [array] → [array_of_pairs]
pub fn emit_enumerate(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[0].add_import("vybe:array", "enumerate");
    let c = &mut chunks[current];
    c.emit_op_u16(Op::CALL_IMPORT, idx, line);
    c.emit(1, line);
}

/// zip(a, b). Stack: [a, b] → [pairs]
pub fn emit_zip(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[0].add_import("vybe:array", "zip");
    let c = &mut chunks[current];
    c.emit_op_u16(Op::CALL_IMPORT, idx, line);
    c.emit(2, line);
}

/// sum(array). Stack: [array] → [number]
pub fn emit_sum(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[0].add_import("vybe:array", "sum");
    let c = &mut chunks[current];
    c.emit_op_u16(Op::CALL_IMPORT, idx, line);
    c.emit(1, line);
}

/// Python min(iterable). Stack: [array] → [value]
pub fn emit_pymin(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[0].add_import("vybe:array", "pymin");
    let c = &mut chunks[current];
    c.emit_op_u16(Op::CALL_IMPORT, idx, line);
    c.emit(1, line);
}

/// Python max(iterable). Stack: [array] → [value]
pub fn emit_pymax(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[0].add_import("vybe:array", "pymax");
    let c = &mut chunks[current];
    c.emit_op_u16(Op::CALL_IMPORT, idx, line);
    c.emit(1, line);
}
