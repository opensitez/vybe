//! Ruby runtime-surface emitters routed via `common:ruby.*`.
//!
//! Ruby is over wasm/js — these compose `ecma:*` host calls directly rather
//! than pulling `__vybe_*` stdlib bundle chunks. All value-method ops are now
//! chunk-free (no `__vybe_*` fallback remains).

use crate::emitter::collections;
use crate::emitter::instructions::core_wasm;
use crate::emitter::ops;
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

/// Emit `<module>.<name>(argc args)` — receiver/args already on the stack.
fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    chunks[current].emit_call(idx, argc, line);
}

pub fn emit_helper(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    match name {
        // `arr.uniq` — order-preserving dedup = `Array.from(new Set(arr))`.
        "ruby.uniq" => {
            call_import(chunks, current, "ecma:set", "fromIterable", 1, line);
            call_import(chunks, current, "ecma:array", "from", 1, line);
        }
        // `x.to_s` — string coercion `x + ""` (dyn_add stringifies any type).
        "ruby.tostring" => {
            chunks[current].emit_string_const("", line);
            ops::emit_dyn_add(&mut chunks[current], line);
        }
        // `x.empty?` — polymorphic length == 0.
        "ruby.isempty" => {
            collections::emit_len(chunks, current, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            ops::emit_dyn_eq(&mut chunks[current], line);
        }
        // `s.encoding` — receiver ignored, constant "UTF-8".
        "ruby.encoding" => {
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_string_const("UTF-8", line);
        }
        // `x.hash` / `x.object_id` — toString-stable stand-in: `String(x).length`.
        "ruby.hash" | "ruby.id" => {
            call_import(chunks, current, "ecma:string", "String", 1, line);
            call_import(chunks, current, "ecma:string", "length", 1, line);
        }
        // `s.bytes` — `TextEncoder().encode(s)` (web:encoding host surface).
        "ruby.bytes" => {
            call_import(chunks, current, "web:encoding", "encoderNew", 0, line);
            call_import(chunks, current, "web:encoding", "encode", 2, line);
        }
        // `arr.minmax` → `[arr.min, arr.max]` via ecma:math:minOf/maxOf (both
        // flatten a single array arg). Stash arr (consumed twice), build [min,max].
        "ruby.minmax" => {
            let base = chunks[current].alloc_scratch(3);
            let (arr_s, min_s, max_s) = (base, base + 1, base + 2);
            chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            call_import(chunks, current, "ecma:math", "minOf", 1, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, min_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            call_import(chunks, current, "ecma:math", "maxOf", 1, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, max_s, line);
            collections::emit_array_new(chunks, current, 0, line);
            core_wasm::dup(&mut chunks[current], line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, min_s, line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            core_wasm::dup(&mut chunks[current], line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, max_s, line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
        }
        // `x.include?(v)` / `x.member?(v)` — polymorphic membership, stack
        // [container, needle] → bool. string → substring; array (incl.
        // materialized ranges) → element; else (hash/object) → own key.
        "ruby.includes" => {
            let needle = chunks[current].alloc_scratch(1);
            let container = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, needle, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, container, line);

            // string → substring test
            chunks[current].emit_op_u16(Op::LOCAL_GET, container, line);
            let test_str = chunks[current].add_import("wasm:js-string", "test");
            chunks[current].emit_call(test_str, 1, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, container, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, needle, line);
            call_import(chunks, current, "ecma:string", "includes", 2, line);
            chunks[current].emit_else(line);

            // array → element test
            chunks[current].emit_op_u16(Op::LOCAL_GET, container, line);
            let is_array = chunks[current].add_import("ecma:array", "isArray");
            chunks[current].emit_call(is_array, 1, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, container, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, needle, line);
            call_import(chunks, current, "ecma:array", "includes", 2, line);
            chunks[current].emit_else(line);

            // hash / object → own-key test
            chunks[current].emit_op_u16(Op::LOCAL_GET, container, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, needle, line);
            call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
        }
        // `arr.compact` — new array with nil (null) elements removed. Inline
        // loop over `ecma:array` primitives (no `__vybe_compact` chunk).
        // Stack: [arr] → [result].
        "ruby.compact" => {
            let base = chunks[current].local_count;
            chunks[current].alloc_scratch(5);
            let (arr_s, result_s, i_s, len_s, elem_s) =
                (base, base + 1, base + 2, base + 3, base + 4);
            chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
            // result = []
            collections::emit_array_new(chunks, current, 0, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
            // len = arr.length; i = 0
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            collections::emit_len(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);

            let block_p = chunks[current].emit_block(line);
            let (loop_p, _) = chunks[current].emit_loop_s(line);
            // cond: break when !(i < len)
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            ops::emit_dyn_lt(&mut chunks[current], line);
            ops::emit_dyn_not(&mut chunks[current], line);
            chunks[current].emit_br_if(1, line);
            // elem = arr[i]
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            collections::emit_get(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
            // if elem != nil → result.push(elem)
            chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
            chunks[current].emit_op(Op::REF_IS_NULL, line);
            chunks[current].emit_op(Op::I32_EQZ, line);
            chunks[current].emit_if(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_end(line);
            // i += 1
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].patch_loop(loop_p);
            chunks[current].emit_end(line);
            chunks[current].patch_block(block_p);
            chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        }
        // `s.hex` — parse a leading hex string → int, invalid → 0. Direct
        // `Number.parseInt(s, 16)` (handles `0x` prefix, sign, partial parse);
        // NaN (no valid digits) → 0. Stack: [s] → [int].
        "ruby.hex" => {
            let r = chunks[current].alloc_scratch(1);
            core_wasm::i32_const(&mut chunks[current], line, 16);
            call_import(chunks, current, "ecma:number", "parseInt", 2, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, r, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
            call_import(chunks, current, "ecma:number", "isNaN", 1, line);
            chunks[current].emit_if_value(line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
            chunks[current].emit_end(line);
        }
        // `a.zip(b, …)` → array of tuples. Shared `vybe_emitter` op (variadic;
        // Ruby/PHP/Python can all route here). `argc` = total arrays on stack.
        "ruby.zip" => {
            collections::emit_zip(chunks, current, argc, collections::ZipLen::First, line);
        }
        // `a.rotate(n=1)` → `a.slice(k, len).concat(a.slice(0, k))` where
        // `k = ((n % len) + len) % len` (left rotate; negative rotates right).
        // Composed from `ecma:array` slice+concat. Stack: [a] or [a, n] → [result].
        "ruby.rotate" => {
            let base = chunks[current].local_count;
            chunks[current].alloc_scratch(4);
            let (arr_s, n_s, len_s, nnorm_s) = (base, base + 1, base + 2, base + 3);
            if argc >= 2 {
                chunks[current].emit_op_u16(Op::LOCAL_SET, n_s, line); // top = n
                chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
            } else {
                chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
                core_wasm::i32_const(&mut chunks[current], line, 1);
                chunks[current].emit_op_u16(Op::LOCAL_SET, n_s, line);
            }
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            collections::emit_len(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
            // if len <= 0 → return arr unchanged (also guards `% 0`)
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            ops::emit_dyn_le(&mut chunks[current], line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            chunks[current].emit_else(line);
            // n_norm = ((n % len) + len) % len
            chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            chunks[current].emit_op(Op::I32_REM_S, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            chunks[current].emit_op(Op::I32_REM_S, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, nnorm_s, line);
            // arr.slice(n_norm, len).concat(arr.slice(0, n_norm))
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, nnorm_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            call_import(chunks, current, "ecma:array", "slice", 3, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_GET, nnorm_s, line);
            call_import(chunks, current, "ecma:array", "slice", 3, line);
            call_import(chunks, current, "ecma:array", "concat", 2, line);
            chunks[current].emit_end(line);
        }
        // `srand(n)` — record n as the seed, seed the global PRNG (reproducible
        // streams), and return the PREVIOUS seed (Ruby semantics). Stack:
        // [n] → [old_seed].
        "ruby.srand" => {
            let n_s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, n_s, line);
            let seed_g = chunks[current].add_constant(vybe_bytecode::Value::String(
                std::sync::Arc::from("__vybe_rng_seed"),
            ));
            chunks[current].emit_op_u16(Op::GLOBAL_GET, seed_g, line); // old seed (null if unset)
            chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
            chunks[current].emit_op_u16(Op::GLOBAL_SET, seed_g, line); // seed = n
            chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
            crate::emitter::random::emit_seed(chunks, current, line); // set PRNG state, pops n
        }
        // `rand` → float [0,1); `rand(n)` → int [0,n). Rides the seedable PRNG.
        "ruby.rand" => {
            if argc >= 1 {
                crate::emitter::random::emit_rand_below(chunks, current, line);
            } else {
                crate::emitter::random::emit_next_unit(chunks, current, line);
            }
        }
        // `a.sample` → one uniformly-random element (null if empty). Shared
        // `vybe_emitter::random` op (Ruby/Python).
        "ruby.sample" => {
            crate::emitter::random::emit_sample(chunks, current, argc, line);
        }
        // `a.shuffle`/`shuffle!` → in-place Fisher-Yates. Shared op (Ruby/Python).
        "ruby.shuffle" => {
            crate::emitter::random::emit_shuffle(chunks, current, argc, line);
        }
        // `h.value?(v)` / `h.has_value?(v)` — `Object.values(h).includes(v)`,
        // direct `ecma:object` (no chunk). Stack: [hash, v] → [bool].
        "ruby.has_value" => {
            let v_s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, v_s, line); // stash v → [hash]
            let values = chunks[current].add_import("ecma:object", "values");
            chunks[current].emit_call(values, 1, line); // [values]
            chunks[current].emit_op_u16(Op::LOCAL_GET, v_s, line); // [values, v]
            collections::emit_contains(chunks, current, line); // [bool]
        }
        // `h.invert` — swap keys/values: `Object.fromEntries(entries.map([k,v]→[v,k]))`.
        // Direct `ecma:object` entries/fromEntries (no chunk). Stack: [hash] → [hash].
        "ruby.invert" => {
            let base = chunks[current].local_count;
            chunks[current].alloc_scratch(5);
            let (entries_s, swapped_s, i_s, len_s, pair_s) =
                (base, base + 1, base + 2, base + 3, base + 4);
            let entries = chunks[current].add_import("ecma:object", "entries");
            chunks[current].emit_call(entries, 1, line); // [entries]
            chunks[current].emit_op_u16(Op::LOCAL_SET, entries_s, line);
            collections::emit_array_new(chunks, current, 0, line); // swapped = []
            chunks[current].emit_op_u16(Op::LOCAL_SET, swapped_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, entries_s, line);
            collections::emit_len(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);

            let block_p = chunks[current].emit_block(line);
            let (loop_p, _) = chunks[current].emit_loop_s(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            ops::emit_dyn_lt(&mut chunks[current], line);
            ops::emit_dyn_not(&mut chunks[current], line);
            chunks[current].emit_br_if(1, line);
            // pair = entries[i]
            chunks[current].emit_op_u16(Op::LOCAL_GET, entries_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            collections::emit_get(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, pair_s, line);
            // swapped.push([pair[1], pair[0]])
            chunks[current].emit_op_u16(Op::LOCAL_GET, swapped_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, pair_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            collections::emit_get(chunks, current, line); // pair[1] (new key)
            chunks[current].emit_op_u16(Op::LOCAL_GET, pair_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            collections::emit_get(chunks, current, line); // pair[0] (new value)
            collections::emit_array_pair(chunks, current, line); // [v, k]
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            // i += 1
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].patch_loop(loop_p);
            chunks[current].emit_end(line);
            chunks[current].patch_block(block_p);
            // result = Object.fromEntries(swapped)
            chunks[current].emit_op_u16(Op::LOCAL_GET, swapped_s, line);
            let from_entries = chunks[current].add_import("ecma:object", "fromEntries");
            chunks[current].emit_call(from_entries, 1, line);
        }
        // `h.transform_values { |v| … }` / `h.transform_keys { |k| … }` →
        // `Object.fromEntries(entries.map([k,v] → [k, blk(v)] | [blk(k), v]))`.
        // Direct `ecma:object` + `CALL_REF` on the block (no chunk).
        // Stack: [hash, block] → [hash].
        "ruby.transform_values" | "ruby.transform_keys" => {
            let on_keys = name == "ruby.transform_keys";
            let base = chunks[current].local_count;
            chunks[current].alloc_scratch(6);
            let (fn_s, entries_s, out_s, i_s, len_s, pair_s) =
                (base, base + 1, base + 2, base + 3, base + 4, base + 5);
            chunks[current].emit_op_u16(Op::LOCAL_SET, fn_s, line); // stash block → [hash]
            let entries = chunks[current].add_import("ecma:object", "entries");
            chunks[current].emit_call(entries, 1, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, entries_s, line);
            collections::emit_array_new(chunks, current, 0, line); // out = []
            chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, entries_s, line);
            collections::emit_len(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);

            let block_p = chunks[current].emit_block(line);
            let (loop_p, _) = chunks[current].emit_loop_s(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            ops::emit_dyn_lt(&mut chunks[current], line);
            ops::emit_dyn_not(&mut chunks[current], line);
            chunks[current].emit_br_if(1, line);
            // pair = entries[i]
            chunks[current].emit_op_u16(Op::LOCAL_GET, entries_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            collections::emit_get(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, pair_s, line);
            // out.push( transform_keys ? [blk(k), v] : [k, blk(v)] )
            chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
            // slot indices: transform the key (0) or the value (1)
            let (transform_idx, keep_idx) = if on_keys { (0, 1) } else { (1, 0) };
            // first element of the new pair
            if on_keys {
                // blk(k)
                chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, pair_s, line);
                core_wasm::i32_const(&mut chunks[current], line, transform_idx);
                collections::emit_get(chunks, current, line);
                chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
                // v (kept)
                chunks[current].emit_op_u16(Op::LOCAL_GET, pair_s, line);
                core_wasm::i32_const(&mut chunks[current], line, keep_idx);
                collections::emit_get(chunks, current, line);
            } else {
                // k (kept)
                chunks[current].emit_op_u16(Op::LOCAL_GET, pair_s, line);
                core_wasm::i32_const(&mut chunks[current], line, keep_idx);
                collections::emit_get(chunks, current, line);
                // blk(v)
                chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, pair_s, line);
                core_wasm::i32_const(&mut chunks[current], line, transform_idx);
                collections::emit_get(chunks, current, line);
                chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
            }
            collections::emit_array_pair(chunks, current, line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].patch_loop(loop_p);
            chunks[current].emit_end(line);
            chunks[current].patch_block(block_p);
            chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
            let from_entries = chunks[current].add_import("ecma:object", "fromEntries");
            chunks[current].emit_call(from_entries, 1, line);
        }
        _ => return false,
    }
    true
}
