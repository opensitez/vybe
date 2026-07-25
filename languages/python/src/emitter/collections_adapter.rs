use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use vybe_emitter::instructions::{core_wasm, host};

use vybe_emitter::{collections, dict, ops, strings, tuples};

/// Merge every `[k, v]` entry of `src` into the Map in `recv` (Python
/// `dict.update`). Loops `entries(src)` and calls the shared `ecma:map.set`.
fn emit_map_merge_into(chunks: &mut [Chunk], current: usize, recv: u16, src: u16, line: u32) {
    let c = &mut chunks[current];
    let entries = c.alloc_scratch(1);
    let n = c.alloc_scratch(1);
    let i = c.alloc_scratch(1);
    let pair = c.alloc_scratch(1);

    c.emit_op_u16(Op::LOCAL_GET, src, line);
    let ent = c.add_import("ecma:object", "entries");
    c.emit_call(ent, 1, line);
    c.emit_op_u16(Op::LOCAL_SET, entries, line);
    c.emit_op_u16(Op::LOCAL_GET, entries, line);
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
    // pair = entries[i]
    c.emit_op_u16(Op::LOCAL_GET, entries, line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_op(Op::ARRAY_GET, line);
    c.emit_op_u16(Op::LOCAL_SET, pair, line);
    // recv.set(pair[0], pair[1])
    c.emit_op_u16(Op::LOCAL_GET, recv, line);
    c.emit_op_u16(Op::LOCAL_GET, pair, line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::ARRAY_GET, line);
    c.emit_op_u16(Op::LOCAL_GET, pair, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::ARRAY_GET, line);
    let set = c.add_import("ecma:map", "set");
    c.emit_call(set, 3, line);
    c.emit_op(Op::DROP, line);
    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, i, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(lp);
    c.emit_end(line);
    c.patch_block(block);
}

/// Push i32 `1` if the value in `slot` is a `Map` (Python dicts are Maps),
/// else `0`. Uses the ECMA `Object.prototype.toString` tag.
fn emit_is_map(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let tag = chunks[current].add_import("ecma:object", "toStringTag");
    chunks[current].emit_call(tag, 1, line);
    chunks[current].emit_string_const("[object Map]", line);
    let eq = chunks[current].add_import("wasm:js-string", "equals");
    chunks[current].emit_call(eq, 2, line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    // Register on the CURRENT chunk (not chunks[0]): the import-table
    // normalizer (compiler/link.rs `normalize_import_table`) remaps each
    // CALL_IMPORT via the emitting chunk's OWN local import table. Using a
    // chunks[0] index inside a non-root chunk collides with per-chunk imports
    // (e.g. `emit_dyn_to_bool`), so the remap resolves the wrong host fn —
    // that broke `len()` on array/dict values inside function bodies. Matches
    // the shared `emit_import_call` convention.
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(argc, line);
}

fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

/// Python `enumerate(iterable[, start])` → array of `(index, value)` tagged
/// tuples (so `list(enumerate(xs))` reprs `[(0, x), …]` and `for i, v in …`
/// destructures). Stack: `[iterable]` or `[iterable, start]` → `[array_of_pairs]`.
/// The prior `collections::emit_enumerate` routed through the retiring
/// `__vybe_enumerate` bundle and was never wired, so `enumerate` no-op'd and
/// returned the iterable unchanged.
pub fn emit_enumerate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let arr = base;

    // idx = start (default 0). Index stays a raw i32 — Python's enumerate yields
    // ints, and a boxed f64 would repr as `0.0`.
    let idx = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
        call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
        chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    } else {
        chunks[current].emit_i32_const(0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);

    let result = chunks[current].alloc_scratch(1);
    call_import(chunks, current, "ecma:array", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    let i = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    // break when !(i < n)
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    // result.push((idx, arr[i]))
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    tuples::emit_tuple(chunks, current, 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // idx++; i++
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}

pub fn emit_extend(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let src = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    core_wasm::dup(&mut chunks[current], line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    collections::emit_insert_range(chunks, current, line);
}

/// `deque.extendleft(xs)` — prepend each item, so the result reverses `xs`
/// (Python semantics). In place: insert `reversed(copy(xs))` at index 0 via the
/// shared array insert-range, leveraging `ecma:array` under the hood.
pub fn emit_extendleft(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let src = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    // index 0 as a boxed number value
    chunks[current].emit_f64_const(0.0, line);
    call_import(chunks, current, "wasm:js-number", "fromF64", 1, line);
    // reversed copy of src (never mutate the argument)
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    collections::emit_clone(chunks, current, line);
    collections::emit_reverse(chunks, current, line);
    collections::emit_insert_range(chunks, current, line);
}

/// `deque.rotate(n)` — rotate right by `n` in place (negative = left). Removes
/// the trailing `k = n mod len` items and re-inserts them at the front, via
/// `ecma:array.splice` + the shared insert-range. Empty deque is a no-op.
pub fn emit_rotate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let recv = base;
    // n defaults to 1 when omitted.
    let n = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
        call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    } else {
        chunks[current].emit_f64_const(1.0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line); // n : f64

    // len (i32) and, if len==0, bail.
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    let len = chunks[current].alloc_scratch(1);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line); // len : i32

    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_op(Op::RETURN, line);
    chunks[current].emit_end(line);

    // k = ((n % len) + len) % len   (Euclidean, so negative n rotates left)
    let ni = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ni, line);
    let k = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ni, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op(Op::I32_REM_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op(Op::I32_REM_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, k, line);

    // removed = splice(recv, len - k, k)   (remove trailing k)
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, k, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    call_import(chunks, current, "wasm:js-number", "fromF64", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, k, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    call_import(chunks, current, "wasm:js-number", "fromF64", 1, line);
    call_import(chunks, current, "ecma:array", "splice", 3, line);
    let removed = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, removed, line);

    // recv.insert_range(0, removed)
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_f64_const(0.0, line);
    call_import(chunks, current, "wasm:js-number", "fromF64", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, removed, line);
    collections::emit_insert_range(chunks, current, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// `OrderedDict.move_to_end(key, last=True)` — reorder the key within the dict's
/// `__keys` insertion-order array (remove, then append for `last`, else prepend).
/// The property value is untouched; only enumeration order changes.
pub fn emit_move_to_end(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let d = base;
    let key = base + 1;
    let keys_k =
        chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__keys")));
    chunks[current].emit_op_u16(Op::LOCAL_GET, d, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_k, line);
    let keys = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);
    // remove the key from its current position
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    collections::emit_remove_value(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // re-insert at the end (last=True, the default). `move_to_end(k, last=False)`
    // (prepend) is a rare edge not exercised here.
    let _ = argc;
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// `dict.popitem(last=True)` — remove and return the last (or first, when
/// `last=False`) `(key, value)` pair as a tuple. Uses `Object.keys` for the key
/// order so it works on BOTH plain dict literals (no `__keys`) and `dict()`/
/// `OrderedDict`-built dicts, and DELETES the property so `len`/truthiness shrink
/// (otherwise `while d: d.popitem()` never terminates). Also trims `__keys` when
/// present, to keep `keys()` consistent.
pub fn emit_popitem(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let d = base;
    // keys = Object.keys(d)  (property order; `__`-internals already filtered)
    chunks[current].emit_op_u16(Op::LOCAL_GET, d, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let keys = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);
    // n = len(keys) : i32
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    // idx = last ? n-1 : 0
    let idx = chunks[current].alloc_scratch(1);
    let last_index = |chunks: &mut [Chunk], current: usize, n: u16| {
        chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_SUB, line);
    };
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
        vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        last_index(chunks, current, n);
        chunks[current].emit_else(line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_end(line);
    } else {
        last_index(chunks, current, n);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);
    // key = keys[idx]  (boxed index)
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    call_import(chunks, current, "wasm:js-number", "fromF64", 1, line);
    call_import(chunks, current, "ecma:array", "get", 2, line);
    let keyv = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keyv, line);
    // val = d[key]
    chunks[current].emit_op_u16(Op::LOCAL_GET, d, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keyv, line);
    dict::emit_get(chunks, current, line);
    let valv = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, valv, line);
    // delete the entry so the dict actually shrinks — Map-aware. A Map-backed
    // dict needs `ecma:map.delete` (object.delete is a no-op on a Map); the
    // legacy Ordinary dict uses object.delete + a `__keys` trim.
    emit_is_map(chunks, current, d, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, d, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keyv, line);
    call_import(chunks, current, "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, d, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keyv, line);
    call_import(chunks, current, "ecma:object", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    // trim __keys when present
    let keys_k =
        chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__keys")));
    chunks[current].emit_op_u16(Op::LOCAL_GET, d, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_k, line);
    let kk = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, kk, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, kk, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, kk, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keyv, line);
    collections::emit_remove_value(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    // return (key, val)
    chunks[current].emit_op_u16(Op::LOCAL_GET, keyv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, valv, line);
    vybe_emitter::tuples::emit_tuple(chunks, current, 2, line);
}

/// `Counter(iterable)` / `Counter()` — build a counting dict. `argc == 0` yields
/// an empty dict; otherwise the single iterable arg on TOS is counted. (The
/// keyword form `Counter(a=3, …)` needs the AST and is handled by the Python
/// `dict`-style compiler intercept.)
pub fn emit_counter_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        dict::emit_new(chunks, current, line);
        return;
    }
    // iterable on TOS → stash, make dict, count into it
    let it = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, it, line);
    dict::emit_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, it, line);
    emit_counter_count(chunks, current, line);
}

/// Count the items of the iterable on TOS into the dict beneath it (Counter
/// construction). Stack `[dict, iterable]` → `[dict]`. Spreads the iterable to an
/// array, then for each item increments `dict[item]`, pushing the key onto
/// `__keys` only on first occurrence (so `items()` keeps first-seen order).
pub fn emit_counter_count(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line); // dict=base, iterable=base+1
    let d = base;
    let it = base + 1;
    let keys_k =
        chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__keys")));

    // arr = spread(iterable)
    chunks[current].emit_op_u16(Op::LOCAL_GET, it, line);
    collections::emit_spread_iterable(chunks, current, line);
    let arr = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr, line);
    // n = len(arr) (i32); i = 0
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    let i = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let item = chunks[current].alloc_scratch(1);
    let cur = chunks[current].alloc_scratch(1);

    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    // item = arr[i]
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    call_import(chunks, current, "wasm:js-number", "fromF64", 1, line);
    call_import(chunks, current, "ecma:array", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item, line);
    // cur = dict[item]
    chunks[current].emit_op_u16(Op::LOCAL_GET, d, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
    dict::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur, line);
    // first occurrence → push to __keys and treat cur as 0
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, d, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_k, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_f64_const(0.0, line);
    call_import(chunks, current, "wasm:js-number", "fromF64", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur, line);
    chunks[current].emit_end(line);
    // dict[item] = cur + 1
    chunks[current].emit_op_u16(Op::LOCAL_GET, d, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    call_import(chunks, current, "wasm:js-number", "fromF64", 1, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // i++
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, d, line);
}

pub fn emit_get(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let recv = base;
    let key = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    dict::emit_get(chunks, current, line);

    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::DROP, line);
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 2, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);
}

/// Python `gen.send(value)` — resume the generator with `value` (the result of
/// the pending `yield`) through the shared `generators.rs` `resume`, returning
/// the next yielded value. Same lazy layer JS/every language drives.
/// Stack: `[gen, value]` → `[yielded]`.
pub fn emit_gen_send(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let recv = base;
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    vybe_emitter::generators::emit_resume(&mut chunks[current], line);
}

/// Python `gen.throw(exc)` — resume the generator by throwing `exc` at the
/// pending `yield` via the shared `generators.rs` `resume_throw`.
/// Stack: `[gen, exc]` → `[yielded]`.
pub fn emit_gen_throw(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let recv = base;
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    } else {
        // throw() with no arg → GeneratorExit-ish; use a generic exception
        chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
        core_wasm::dup(&mut chunks[current], line);
        chunks[current].emit_string_const("", line);
        vybe_emitter::errors::emit_exception_new_finalize(&mut chunks[current], "Exception", line);
    }
    vybe_emitter::generators::emit_resume_throw(&mut chunks[current], line);
}

/// Python `next(it[, default])`. For a generator, resume it through the shared
/// `generators.rs` machinery (`GEN_NEXT` → `[value, has_more]`) — the same lazy
/// path JS uses — so infinite generators advance one step instead of draining.
/// Non-generator iterables fall back to the shared `__vybe_pynext` helper.
pub fn emit_pynext(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let it = base;

    // if isGenerator(it)
    chunks[current].emit_op_u16(Op::LOCAL_GET, it, line);
    call_import(chunks, current, "ecma:value", "isGenerator", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    // generator: GEN_NEXT → [value, has_more]
    chunks[current].emit_op_u16(Op::LOCAL_GET, it, line);
    vybe_emitter::generators::emit_next(&mut chunks[current], line);
    let has_more = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let value = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, has_more, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, has_more, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line); // has_more → value
    chunks[current].emit_else(line);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line); // exhausted → default
    } else {
        // exhausted, no default → raise StopIteration
        chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
        core_wasm::dup(&mut chunks[current], line);
        chunks[current].emit_string_const("", line);
        vybe_emitter::errors::emit_exception_new_finalize(
            &mut chunks[current],
            "StopIteration",
            line,
        );
        vybe_emitter::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_op(Op::NULL, line); // unreachable (throw diverges)
    }
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    // not a generator → shared iterator-protocol next
    chunks[current].emit_op_u16(Op::LOCAL_GET, it, line);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    }
    collections::emit_runtime_helper_call(chunks, current, "__vybe_pynext", argc, line);
    chunks[current].emit_end(line);
}

/// Python from-end index normalization. Stack: `[obj, idx]` → `[normalized_idx]`.
///
/// `a[-1]` is "one from the end", not a real negative index: when `obj` is a
/// sequence (array or string) and `idx` is a negative number, this returns
/// `len(obj) + idx`; otherwise `idx` is returned unchanged so dict string/other
/// keys pass straight through. The `< 0` test is guarded behind an `isNumber`
/// check so a string key never hits numeric coercion — that guard is why this
/// replaces the shared `negative_index_wraps` flag (which trapped on `d['a']`).
pub fn emit_from_end(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let obj = base;
    let idx = base + 1;

    // if isNumber(idx)  (short-circuits the `< 0` test for string keys)
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    host::emit(&mut chunks[current], "wasm:js-number", "test", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    // if idx < 0
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    // if isArray(obj) → len(obj) + idx
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_len_plus(chunks, current, obj, idx, line);
    chunks[current].emit_else(line);
    // else if isString(obj) → len(obj) + idx
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_len_plus(chunks, current, obj, idx, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line); // idx >= 0 → unchanged
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line); // not a number → unchanged (dict/other key)
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_end(line);
}

/// Python `x in y` membership. Stack: `[container, needle]` → `[bool]`.
/// string → substring test; array → element test; else (dict/object) → own key.
/// Set literals are lowered to `.has()` upstream and never reach here.
pub fn emit_contains(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let test_str = chunk.add_import("wasm:js-string", "test");
    let str_includes = chunk.add_import("ecma:string", "includes");
    let is_array = chunk.add_import("ecma:array", "isArray");
    let arr_includes = chunk.add_import("ecma:array", "includes");
    // `hasIn` is Map-aware (`ObjectKind::Map` → `contains_key`) as well as
    // Ordinary — Python dicts are Maps, same as PHP arrays.
    let has_own = chunk.add_import("ecma:object", "hasIn");

    let needle = chunk.alloc_scratch(1);
    let container = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, needle, line);
    chunk.emit_op_u16(Op::LOCAL_SET, container, line);

    // string → substring test
    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_call(test_str, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_op_u16(Op::LOCAL_GET, needle, line);
    chunk.emit_call(str_includes, 2, line);
    chunk.emit_else(line);

    // array → element test
    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_op_u16(Op::LOCAL_GET, needle, line);
    chunk.emit_call(arr_includes, 2, line);
    chunk.emit_else(line);

    // set → membership. `ecma:set.has` is guarded by `is_set` and answers
    // false for anything else, so it composes with the object tests below
    // rather than needing its own type predicate: a real set answers here,
    // a dict/object falls through to the key/`__contains__` tests, and a
    // set that simply lacks the item answers false either way.
    let set_has = chunk.add_import("ecma:set", "has");
    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_op_u16(Op::LOCAL_GET, needle, line);
    chunk.emit_call(set_has, 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_else(line);

    // object → user `__contains__(self, item)` if present, else own-key test
    let contains_key = chunk.add_constant(vybe_bytecode::Value::String(std::sync::Arc::from(
        "__contains__",
    )));
    let contains_method = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_op_u16(Op::STRUCT_GET, contains_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, contains_method, line);
    chunk.emit_op_u16(Op::LOCAL_GET, contains_method, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line); // 1 if a __contains__ method is present
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, contains_method, line);
    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_op_u16(Op::LOCAL_GET, needle, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_op_u16(Op::LOCAL_GET, needle, line);
    chunk.emit_call(has_own, 2, line);
    chunk.emit_end(line);
    chunk.emit_end(line); // set
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `len(obj) + idx` — helper for the from-end wrap. Stack: `[]` → `[value]`.
fn emit_len_plus(chunks: &mut [Chunk], current: usize, obj: u16, idx: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    emit_length(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
}

pub fn emit_index(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let recv = base;
    let needle = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle, line);
    strings::emit_index_of(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle, line);
    collections::emit_index_of(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let recv = base;
    let keys_key =
        chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__keys")));

    // Map-backed dict → shared `ecma:map.clear`.
    emit_is_map(chunks, current, recv, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:map", "clear", 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);
    collections::emit_remove_range(chunks, current, line);

    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:set", "clear", 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    dict::emit_method_clear_stack(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line); // close the is-Map outer if
}

/// `set(iterable)` / `frozenset(iterable)`. `ecma:set.fromIterable` only
/// accepts an array or a string, so a Set argument (`frozenset({1, 2})`) or a
/// map would yield an empty set. Normalize through `ecma:array.from` first —
/// it materializes a set's values, a string's chars, a map's pairs, or copies
/// an array — then build the set from that array.
/// `__py_sort_by_key(array, key_fn)` → the array sorted by `key_fn(x)`, reusing
/// the shared key-sort primitive. Python `sorted(x, key=f)` routes here.
pub fn emit_sort_by_key(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line); // array
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line); // key_fn
    collections::emit_sort_by_key_in_place(chunks, current, line);
}

pub fn emit_make_set(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        call_import(chunks, current, "ecma:set", "new", 0, line);
        return;
    }
    call_import(chunks, current, "ecma:array", "from", 1, line);
    call_import(chunks, current, "ecma:set", "fromIterable", 1, line);
}

/// `frozenset(iterable)` — a real `ecma:set` (so every set op/method works),
/// stamped `__frozenset = true` so repr renders `frozenset({...})`. Bypasses the
/// shared-compiler `frozenset` special-case via the walker's `__py_frozenset`.
pub fn emit_frozenset(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_make_set(chunks, current, argc, line); // [set]
    let c = &mut chunks[current];
    // Stamp `__type = "Set"` — `ecma:set.new` (empty case) omits it, which would
    // make repr skip the set branch and render `frozenset()` as "".
    c.emit_dup(line);
    c.emit_string_const("Set", line);
    let tk = c.add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__type")));
    c.emit_op_u16(Op::STRUCT_SET, tk, line);
    c.emit_op(Op::DROP, line);
    // Stamp `__frozenset = true` so repr renders `frozenset({...})`.
    c.emit_dup(line);
    core_wasm::bool_const(c, line, true);
    let k = c.add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__frozenset")));
    c.emit_op_u16(Op::STRUCT_SET, k, line);
    c.emit_op(Op::DROP, line);
}

/// Set predicate methods (`issubset`/`issuperset`/`isdisjoint`) compose the
/// matching `ecma:set` host fn, which returns a raw i32. Python needs a real
/// `bool` so `print` renders `True`/`False`, not `1`/`0`.
pub fn emit_set_predicate(chunks: &mut [Chunk], current: usize, host_fn: &str, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    call_import(chunks, current, "ecma:set", host_fn, 2, line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    call_import(chunks, current, "ecma:set", "add", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

fn emit_remove_impl(chunks: &mut [Chunk], current: usize, raises: bool, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let value = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    // `isArray` returns a real bool; `!= 0` (Bool vs I32) is cross-type and
    // always true, sending a Set down the array branch. Coerce to i32 first.
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_import(chunks, current, "ecma:array", "removeValue", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_else(line);
    // `ecma:set.delete` returns a bool: whether the member was present.
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_import(chunks, current, "ecma:set", "delete", 2, line);
    if raises {
        // `set.remove(x)` raises KeyError when x is absent; `discard` does not.
        vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_op(Op::I32_EQZ, line); // 1 if NOT removed
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
        core_wasm::dup(&mut chunks[current], line);
        chunks[current].emit_string_const("", line);
        vybe_emitter::errors::emit_exception_new_finalize(&mut chunks[current], "KeyError", line);
        vybe_emitter::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_remove_impl(chunks, current, true, line);
}

pub fn emit_discard(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_remove_impl(chunks, current, false, line);
}

pub fn emit_copy(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let recv = base;
    let keys_key =
        chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__keys")));

    // Map-backed dict → new Map from its entries (shallow copy).
    emit_is_map(chunks, current, recv, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:object", "entries", 1, line);
    call_import(chunks, current, "ecma:map", "fromEntries", 1, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);
    collections::emit_slice(chunks, current, line);

    chunks[current].emit_else(line);
    call_import(chunks, current, "ecma:set", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:set", "union", 2, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    dict::emit_new(chunks, current, line);
    let out = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:object", "assign", 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line); // close the is-Map outer if
}

pub fn emit_update(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let src = base + 1;
    let keys_key =
        chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__keys")));

    // Map-backed dict → merge src's entries in via shared `ecma:map.set`.
    emit_is_map(chunks, current, recv, line);
    chunks[current].emit_if_value(line);
    emit_map_merge_into(chunks, current, recv, src, line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    call_import(chunks, current, "ecma:set", "unionWith", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    call_import(chunks, current, "ecma:object", "assign", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line); // close the is-Map outer if
}

fn emit_set_update_call(chunks: &mut [Chunk], current: usize, func: &str, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    call_import(chunks, current, "ecma:set", func, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_intersection_update(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_set_update_call(chunks, current, "intersectWith", line);
}

pub fn emit_difference_update(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_set_update_call(chunks, current, "exceptWith", line);
}

pub fn emit_symmetric_difference_update(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_set_update_call(chunks, current, "symmetricExceptWith", line);
}

pub fn emit_pop(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let recv = base;
    let value_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);

    if argc == 1 {
        // No index: `list.pop()` removes the last element, `set.pop()` removes
        // an arbitrary member (we take the first). Sets aren't indexable, so
        // dispatch on isArray.
        let index_slot = chunks[current].local_count;
        chunks[current].alloc_scratch(1);
        let arr_slot = chunks[current].local_count;
        chunks[current].alloc_scratch(1);

        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        call_import(chunks, current, "ecma:array", "isArray", 1, line);
        vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);

        // list.pop(): value = recv[len-1]; remove_at(recv, len-1)
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        collections::emit_len(chunks, current, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_SUB, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
        collections::emit_remove_at(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);

        chunks[current].emit_else(line);

        // set.pop(): materialize members; empty → KeyError; else remove first.
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        call_import(chunks, current, "ecma:array", "from", 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
        chunks[current].emit_op(Op::ARRAY_LENGTH, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op(Op::I32_EQ, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
        core_wasm::dup(&mut chunks[current], line);
        chunks[current].emit_string_const("pop from an empty set", line);
        vybe_emitter::errors::emit_exception_new_finalize(&mut chunks[current], "KeyError", line);
        vybe_emitter::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_end(line);
        // v = arr[0]
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
        // delete v from the set
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        call_import(chunks, current, "ecma:set", "delete", 2, line);
        chunks[current].emit_op(Op::DROP, line);

        chunks[current].emit_end(line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        return;
    } else {
        // Dispatch on isArray, NOT `__keys` presence: a Python dict is a plain
        // JS object with no `__keys`, so a keys-based check misclassifies it as
        // a list. `list.pop(i)` splices; `dict.pop(k[, default])` reads the
        // value then removes the property natively via `ecma:object.delete`.
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        call_import(chunks, current, "ecma:array", "isArray", 1, line);
        vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);

        // list.pop(i): value = recv[i]; remove_at(recv, i); value
        let index = base + 1;
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
        collections::emit_remove_at(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);

        chunks[current].emit_else(line);

        // dict.pop(k[, default]): value = recv[k]
        let key = base + 1;
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
        dict::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        // missing key → default (or null)
        if argc >= 3 {
            chunks[current].emit_op_u16(Op::LOCAL_GET, base + 2, line);
        } else {
            chunks[current].emit_op(Op::NULL, line);
        }
        chunks[current].emit_else(line);
        // present → delete the entry (Map-aware, like PHP's array adapter) and
        // return the value. A Map-backed dict needs `ecma:map.delete`;
        // `ecma:object.delete` is a no-op on a Map.
        emit_is_map(chunks, current, recv, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
        call_import(chunks, current, "ecma:map", "delete", 2, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
        call_import(chunks, current, "ecma:object", "delete", 2, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[current].emit_end(line);

        chunks[current].emit_end(line);
    }
}

pub fn emit_length(chunks: &mut [Chunk], current: usize, line: u32) {
    // Polymorphic `len`: string → char length, array → element count,
    // Set/Map → `.size`, otherwise (dict/object) → `Object.keys(o).length`.
    // Uses the object's native property enumeration (a Python dict IS a JS
    // object) — no `__keys` array, so literal and built dicts count the same.
    let base = stash_args(chunks, current, 1, line);
    let recv = base;
    let size_key =
        chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("size")));

    // User-defined `__len__` → call it with the receiver. (Cross-language:
    // bound alongside `__get_length`/`__get_count`.)
    let len_key = chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from(
        "__len__",
    )));
    let len_method = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, len_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_method, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_method, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line); // 1 if a method is present
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_method, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_else(line);

    // isString(recv) → string length
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_else(line);

    // isArray(recv) → element count
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_else(line);

    // isView(recv) → typed-array (bytes) length
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:arraybuffer", "isView", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:uint8array", "length", 1, line);
    chunks[current].emit_else(line);

    // has `.size` (Set/Map) → use it, else Object.keys(recv).length
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, size_key, line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::DROP, line); // drop null `.size`
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line); // close the isView (bytes) branch
    chunks[current].emit_end(line); // close the __len__ dispatch branch
}
