//! Centralized pseudo-random emitter — one seedable PRNG plus the derived ops
//! shared by every language's random surface (Ruby `rand`/`srand`/`sample`/
//! `shuffle`, Python `random.*`, PHP `rand`/`mt_rand`/`shuffle`, …).
//!
//! Design:
//! - A single global cell `__vybe_rng` holds an `i32` xorshift32 state. While
//!   unseeded it lazily seeds from `ecma:math:random()` (non-deterministic per
//!   run); [`emit_seed`] makes the stream deterministic/reproducible.
//! - [`emit_next_unit`] is the ONE entropy tap (→ `f64` in `[0, 1)`); every
//!   derived op (int/float ranges, `sample`, `shuffle`) rides on it, so a seed
//!   deterministically drives all of them.
//! - Crypto-strength randomness (secure ints/bytes/UUID) is intentionally NOT
//!   here — those map straight to `wasi:random` at the language layer.
//!
//! The PRNG itself is plain WASM integer ops (xorshift32); range/select math
//! leans on the existing `collections`/`ops`/`ecma:math` surfaces rather than
//! reinventing them.

use crate::primitives::instructions::core_wasm;
use std::sync::Arc;
use vybe_runtime::Chunk;
use vybe_runtime::Value;
use vybe_runtime::opcode::Op;

const RNG_GLOBAL: &str = "__vybe_rng";

fn rng_global(chunk: &mut Chunk) -> u16 {
    chunk.add_constant(Value::String(Arc::from(RNG_GLOBAL)))
}

/// `seed(n)` / `srand(n)` — set the global PRNG state deterministically from
/// `n` (mixed and forced non-zero so xorshift never sticks at 0). Same seed →
/// same subsequent stream. Stack: `[n]` → `[]`.
pub fn emit_seed(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let s = base;
    chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
    // state = (n ^ (n << 13)) | 1
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    core_wasm::i32_const(&mut chunks[current], line, 13);
    chunks[current].emit_op(Op::I32_SHL, line);
    chunks[current].emit_op(Op::I32_XOR, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_OR, line);
    let g = rng_global(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::GLOBAL_SET, g, line);
}

/// One PRNG step → uniform `f64` in `[0, 1)`. Lazily seeds the global cell from
/// `ecma:math:random` when still null (unseeded). Stack: `[]` → `[f64]`.
pub fn emit_next_unit(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let s = base;
    // s = global
    let g = rng_global(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::GLOBAL_GET, g, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
    // if null → seed from ecma:math:random: s = floor(random() * 2^30) | 1
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    let r = chunks[current].add_import("ecma:math", "random");
    chunks[current].emit_call(r, 0, line);
    core_wasm::f64_const(&mut chunks[current], line, 1073741824.0); // 2^30
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::I32_FROM_F64, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
    chunks[current].emit_end(line);
    // xorshift32: s ^= s << 13; s ^= s >>> 17; s ^= s << 5
    for (shift, op) in [(13i32, Op::I32_SHL), (17, Op::I32_SHR_U), (5, Op::I32_SHL)] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        core_wasm::i32_const(&mut chunks[current], line, shift);
        chunks[current].emit_op(op, line);
        chunks[current].emit_op(Op::I32_XOR, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
    }
    // store back
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    let g2 = rng_global(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::GLOBAL_SET, g2, line);
    // (s & 0x7FFFFF) / 2^23  → [0, 1) with 23 bits of entropy
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0x7FFFFF);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    core_wasm::f64_const(&mut chunks[current], line, 8388608.0); // 2^23
    chunks[current].emit_op(Op::F64_DIV, line);
}

/// `rand(n)` / `randrange(n)` → uniform int in `[0, n)`, returned as a boxed
/// number (`ecma:math:floor` boxes it so `==`/printing behave). Stack: `[n]` → `[int]`.
pub fn emit_rand_below(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let n = base;
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    emit_next_unit(chunks, current, line); // u (f64)
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    chunks[current].emit_op(Op::F64_MUL, line); // u * n
    let floor = chunks[current].add_import("ecma:math", "floor");
    chunks[current].emit_call(floor, 1, line); // → boxed int
}

/// `randint(lo, hi)` / `rand(lo..hi)` → uniform int in `[lo, hi]` (inclusive),
/// boxed. Stack: `[lo, hi]` → `[int]`.
pub fn emit_rand_int_inclusive(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].local_count;
    chunks[current].alloc_scratch(2);
    let (lo, hi) = (base, base + 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, hi, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lo, line);
    emit_next_unit(chunks, current, line); // u
    // range = hi - lo + 1
    chunks[current].emit_op_u16(Op::LOCAL_GET, hi, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lo, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    chunks[current].emit_op(Op::F64_MUL, line); // u * range
    let floor = chunks[current].add_import("ecma:math", "floor");
    chunks[current].emit_call(floor, 1, line); // boxed offset in [0, range)
    chunks[current].emit_op_u16(Op::LOCAL_GET, lo, line);
    crate::primitives::ops::emit_dyn_add(&mut chunks[current], line); // + lo (boxed)
}

/// `arr.sample` (Ruby) / `random.choice(arr)` (Python) → one uniformly-random
/// element, or null if empty. `argc` = total operands; args past the receiver
/// (e.g. an ignored count/seed) are dropped. Stack: `[arr, extra…]` → `[element]`.
pub fn emit_sample(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 1..argc.max(1) {
        chunks[current].emit_op(Op::DROP, line);
    }
    let base = chunks[current].local_count;
    chunks[current].alloc_scratch(3);
    let (arr_s, len_s, idx_s) = (base, base + 1, base + 2);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    crate::primitives::collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    // empty → null
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    crate::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);
    // idx = floor(next_unit() * len); arr[idx]
    emit_next_unit(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::I32_FROM_F64, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
}

/// `arr.shuffle`/`shuffle!` (Ruby) / `random.shuffle(arr)` (Python) → in-place
/// Fisher-Yates, returning the (same, now-shuffled) array so `shuffle!`
/// preserves identity. `argc` = total operands; args past the receiver (e.g. a
/// `random:` seed) are dropped. Stack: `[arr, extra…]` → `[arr]`.
pub fn emit_shuffle(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 1..argc.max(1) {
        chunks[current].emit_op(Op::DROP, line);
    }
    let base = chunks[current].local_count;
    chunks[current].alloc_scratch(5);
    let (arr_s, i_s, j_s, tmp_s, len_s) = (base, base + 1, base + 2, base + 3, base + 4);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    crate::primitives::collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    // i = len - 1
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);

    let block_p = chunks[current].emit_block(line);
    let (loop_p, _) = chunks[current].emit_loop_s(line);
    // exit when i <= 0
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    crate::primitives::ops::emit_dyn_le(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    // j = floor(next_unit() * (i + 1))
    emit_next_unit(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::I32_FROM_F64, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j_s, line);
    // tmp = arr[i]
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, tmp_s, line);
    // arr[i] = arr[j]
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j_s, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    crate::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // arr[j] = tmp
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tmp_s, line);
    crate::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // i--
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_p);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_p);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
}

/// `random.sample(population, k)` → a NEW list of `k` unique elements (no
/// repeats), via a partial Fisher-Yates over a copy so the population is not
/// mutated. Rides the seedable PRNG. Stack: `[population, k]` → `[result]`.
pub fn emit_sample_k(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].local_count;
    chunks[current].alloc_scratch(8);
    let (pop_s, k_s, copy_s, n_s, i_s, j_s, tmp_s, result_s) = (
        base,
        base + 1,
        base + 2,
        base + 3,
        base + 4,
        base + 5,
        base + 6,
        base + 7,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, k_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pop_s, line);
    // n = pop.length
    chunks[current].emit_op_u16(Op::LOCAL_GET, pop_s, line);
    crate::primitives::collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n_s, line);
    // copy = pop.slice(0, n)  — don't mutate the population
    chunks[current].emit_op_u16(Op::LOCAL_GET, pop_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
    crate::primitives::collections::emit_slice(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, copy_s, line);
    // result = []; i = 0
    crate::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);

    let block_p = chunks[current].emit_block(line);
    let (loop_p, _) = chunks[current].emit_loop_s(line);
    // exit when !(i < k)
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, k_s, line);
    crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    // j = i + floor(next_unit() * (n - i))
    emit_next_unit(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::I32_FROM_F64, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j_s, line);
    // swap copy[i], copy[j]
    chunks[current].emit_op_u16(Op::LOCAL_GET, copy_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, tmp_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, copy_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, copy_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j_s, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    crate::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, copy_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tmp_s, line);
    crate::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // result.push(copy[i])
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, copy_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    crate::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // i++
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
