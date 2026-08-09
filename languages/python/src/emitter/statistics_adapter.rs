//! Python `statistics` adapter — bytecode-only.
//!
//! Composes `ecma:math.*` / `ecma:array.*` into the `statistics` surface.
//! No new host fns: these are ordinary arithmetic over a list, so they are
//! emitted here rather than invented as host builtins (there is no
//! `Math.mean` in ECMA to route to).
//!
//! Every result is an f64. Python's own display rules then apply for free:
//! `mean([42])` is `42`, `mean([1,2,3,4])` is `2.5`, `pvariance([1,2,3])` is
//! `0.6666666666666666` — no float-repr wrapping needed.

use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// `sum(data)`. Stack: `[data]` → `[num]`.
fn emit_sum(chunk: &mut Chunk, data: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, data, line);
    let sum = chunk.add_import("ecma:math", "sumPrecise");
    chunk.emit_call(sum, 1, line);
}

/// `len(data)` as f64. Stack: `[]` → `[num]`.
fn emit_len(chunks: &mut [Chunk], current: usize, data: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
}

/// Stash the single list argument into a local. Stack: `[data]` → `[]`.
fn stash_data(chunk: &mut Chunk, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

/// `statistics.mean(data)` / `fmean(data)`. Stack: `[data]` → `[num]`.
pub fn emit_mean(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let data = stash_data(&mut chunks[current], line);
    emit_sum(&mut chunks[current], data, line);
    emit_len(chunks, current, data, line);
    chunks[current].emit_op(Op::F64_DIV, line);
}

/// The data sorted ascending, as a new list. Stack: `[]` → `[array]`.
fn emit_sorted(chunk: &mut Chunk, data: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, data, line);
    let sorted = chunk.add_import("ecma:array", "toSorted");
    chunk.emit_call(sorted, 1, line);
}

/// `s[i]` where `i` is an f64-valued local. Stack: `[]` → `[value]`.
fn emit_at(chunk: &mut Chunk, arr: u16, idx: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, arr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

/// `statistics.median(data)` — the middle of the sorted data, or the mean of
/// the middle two when the count is even. Stack: `[data]` → `[num]`.
pub fn emit_median(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let data = stash_data(&mut chunks[current], line);
    let s = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let mid = chunks[current].alloc_scratch(1);

    emit_sorted(&mut chunks[current], data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
    emit_len(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    let chunk = &mut chunks[current];
    // mid = n / 2  (integer division)
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    core_wasm::i32_const(chunk, line, 2);
    chunk.emit_op(Op::I32_DIV_S, line);
    chunk.emit_op_u16(Op::LOCAL_SET, mid, line);

    // n % 2 == 1 → the exact middle; else the mean of the two straddling it.
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    core_wasm::i32_const(chunk, line, 2);
    chunk.emit_op(Op::I32_REM_S, line);
    chunk.emit_if_value(line);
    emit_at(chunk, s, mid, line);
    chunk.emit_else(line);
    emit_at(chunk, s, mid, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_op_u16(Op::LOCAL_GET, mid, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::F64_ADD, line);
    core_wasm::f64_const(chunk, line, 2.0);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_end(line);
}

/// `median_low` / `median_high` — for an even count Python takes the lower or
/// upper of the two middle values rather than averaging them.
fn emit_median_side(chunks: &mut [Chunk], current: usize, high: bool, line: u32) {
    let data = stash_data(&mut chunks[current], line);
    let s = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);

    emit_sorted(&mut chunks[current], data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
    emit_len(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    let chunk = &mut chunks[current];
    // low → (n - 1) / 2 ; high → n / 2. They coincide for an odd count.
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    if !high {
        core_wasm::i32_const(chunk, line, 1);
        chunk.emit_op(Op::I32_SUB, line);
    }
    core_wasm::i32_const(chunk, line, 2);
    chunk.emit_op(Op::I32_DIV_S, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx, line);
    emit_at(chunk, s, idx, line);
}

pub fn emit_median_low(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_median_side(chunks, current, false, line);
}

pub fn emit_median_high(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_median_side(chunks, current, true, line);
}

/// Sum of `(x - mean)**2` over the data — the numerator both variances share.
/// Stack: `[]` → `[num]`, with `data`/`mean` already in locals.
fn emit_sq_dev_sum(chunks: &mut [Chunk], current: usize, data: u16, mean: u16, line: u32) {
    let acc = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);

    core_wasm::f64_const(&mut chunks[current], line, 0.0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    emit_len(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(std::slice::from_mut(chunk), 0, line);

    // acc += (data[i] - mean) ** 2
    chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
    chunk.emit_op_u16(Op::LOCAL_GET, data, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_GET, mean, line);
    chunk.emit_op(Op::F64_SUB, line);
    let d = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, d, line);
    chunk.emit_op_u16(Op::LOCAL_GET, d, line);
    chunk.emit_op_u16(Op::LOCAL_GET, d, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc, line);

    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
}

/// Shared variance body. `sample` selects Bessel's correction (`n - 1`),
/// which is what separates `variance` from `pvariance`.
fn emit_variance_inner(chunks: &mut [Chunk], current: usize, data: u16, sample: bool, line: u32) {
    let mean = chunks[current].alloc_scratch(1);
    emit_sum(&mut chunks[current], data, line);
    emit_len(chunks, current, data, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mean, line);

    emit_sq_dev_sum(chunks, current, data, mean, line);
    emit_len(chunks, current, data, line);
    if sample {
        core_wasm::f64_const(&mut chunks[current], line, 1.0);
        chunks[current].emit_op(Op::F64_SUB, line);
    }
    chunks[current].emit_op(Op::F64_DIV, line);
}

pub fn emit_variance(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let data = stash_data(&mut chunks[current], line);
    emit_variance_inner(chunks, current, data, true, line);
}

pub fn emit_pvariance(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let data = stash_data(&mut chunks[current], line);
    emit_variance_inner(chunks, current, data, false, line);
}

fn emit_sqrt(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_SQRT, line);
}

pub fn emit_stdev(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let data = stash_data(&mut chunks[current], line);
    emit_variance_inner(chunks, current, data, true, line);
    emit_sqrt(&mut chunks[current], line);
}

pub fn emit_pstdev(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let data = stash_data(&mut chunks[current], line);
    emit_variance_inner(chunks, current, data, false, line);
    emit_sqrt(&mut chunks[current], line);
}

/// How many times `data[j]` equals `data[i]`, for the mode scan.
/// Stack: `[]` → `[count]`, leaves `j` clobbered.
fn emit_count_of(chunks: &mut [Chunk], current: usize, data: u16, i: u16, n: u16, line: u32) {
    let count = chunks[current].alloc_scratch(1);
    let j = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, j, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(std::slice::from_mut(chunk), 0, line);

    chunk.emit_op_u16(Op::LOCAL_GET, data, line);
    chunk.emit_op_u16(Op::LOCAL_GET, j, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_GET, data, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, count, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, count, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, j, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, j, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, count, line);
}

/// `statistics.mode(data)` — the most common value; ties go to the one seen
/// first, which is what Python guarantees since 3.8. Stack: `[data]` → `[value]`.
pub fn emit_mode(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let data = stash_data(&mut chunks[current], line);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let best = chunks[current].alloc_scratch(1);
    let best_count = chunks[current].alloc_scratch(1);

    emit_len(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best_count, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    emit_count_of(chunks, current, data, i, n, line);
    let c = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, c, line);

    // Strictly greater — so an equal count keeps the earlier value.
    chunk.emit_op_u16(Op::LOCAL_GET, c, line);
    chunk.emit_op_u16(Op::LOCAL_GET, best_count, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, c, line);
    chunk.emit_op_u16(Op::LOCAL_SET, best_count, line);
    chunk.emit_op_u16(Op::LOCAL_GET, data, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, best, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, best, line);
}

/// `statistics.multimode(data)` — every value tying the top count, in first-seen
/// order. Stack: `[data]` → `[array]`.
pub fn emit_multimode(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let data = stash_data(&mut chunks[current], line);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let best_count = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);

    emit_len(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best_count, line);

    // Pass 1 — the winning count.
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let s1 = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    emit_count_of(chunks, current, data, i, n, line);
    let c1 = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, c1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, c1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, best_count, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, c1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, best_count, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, s1, line);

    // Pass 2 — collect the winners, skipping duplicates.
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let s2 = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    emit_count_of(chunks, current, data, i, n, line);
    let c2 = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, c2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, c2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, best_count, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    // `x in out` guards the duplicate — each winner is listed once.
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    let includes = chunks[current].add_import("ecma:array", "includes");
    chunks[current].emit_call(includes, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, s2, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `statistics.quantiles(data, n=4)` — CPython's default "exclusive" method:
/// `n - 1` cut points interpolated over the sorted data. Stack: `[data]` → `[array]`.
pub fn emit_quantiles(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk0 = &mut chunks[current];
    // `n` defaults to 4 (quartiles).
    let nq = chunk0.alloc_scratch(1);
    if argc >= 2 {
        chunk0.emit_op_u16(Op::LOCAL_SET, nq, line);
    } else {
        core_wasm::i32_const(chunk0, line, 4);
        chunk0.emit_op_u16(Op::LOCAL_SET, nq, line);
    }
    let data = stash_data(&mut chunks[current], line);
    let s = chunks[current].alloc_scratch(1);
    let ld = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let j = chunks[current].alloc_scratch(1);
    let delta = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);

    emit_sorted(&mut chunks[current], data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
    emit_len(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ld, line);
    // m = ld + 1
    chunks[current].emit_op_u16(Op::LOCAL_GET, ld, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, m, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, nq, line);
    chunk.emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(std::slice::from_mut(chunk), 0, line);

    // j = clamp(i * m / nq, 1, ld - 1)
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, m, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, nq, line);
    chunk.emit_op(Op::I32_DIV_S, line);
    chunk.emit_op_u16(Op::LOCAL_SET, j, line);
    chunk.emit_op_u16(Op::LOCAL_GET, j, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_if(line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op_u16(Op::LOCAL_SET, j, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ld, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_GET, j, line);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ld, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, j, line);
    chunk.emit_end(line);

    // delta = i*m - j*nq
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, m, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, j, line);
    chunk.emit_op_u16(Op::LOCAL_GET, nq, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, delta, line);

    // out.push((s[j-1] * (nq - delta) + s[j] * delta) / nq)
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_op_u16(Op::LOCAL_GET, j, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_GET, nq, line);
    chunk.emit_op_u16(Op::LOCAL_GET, delta, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_op_u16(Op::LOCAL_GET, j, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_GET, delta, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, nq, line);
    chunk.emit_op(Op::F64_DIV, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `statistics.median_grouped(data, interval=1)` — the median of continuous
/// data, interpolated within the interval the midpoint falls in:
/// `L + interval * (n/2 - cf) / f`. Stack: `[data]` → `[num]`.
pub fn emit_median_grouped(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let data = stash_data(&mut chunks[current], line);
    let s = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let x = chunks[current].alloc_scratch(1);
    let cf = chunks[current].alloc_scratch(1);
    let f = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);

    emit_sorted(&mut chunks[current], data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
    emit_len(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    // x = s[n // 2] — the value the median falls on.
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    core_wasm::i32_const(chunk, line, 2);
    chunk.emit_op(Op::I32_DIV_S, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, x, line);

    // cf = how many values precede x; f = how many equal it. One pass over
    // the sorted data gives both.
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, cf, line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, f, line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(std::slice::from_mut(chunk), 0, line);

    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, cf, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cf, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, f, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, f, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    // (x - 0.5) + (n/2 - cf) / f      [interval = 1]
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    core_wasm::f64_const(chunk, line, 0.5);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    core_wasm::f64_const(chunk, line, 2.0);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cf, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_GET, f, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_ADD, line);
}

/// `harmonic_mean(data)` = `n / sum(1/x)`. Stack: `[data]` → `[num]`.
pub fn emit_harmonic_mean(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let data = stash_data(&mut chunks[current], line);
    let acc = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);

    core_wasm::f64_const(&mut chunks[current], line, 0.0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    emit_len(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(std::slice::from_mut(chunk), 0, line);

    chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
    core_wasm::f64_const(chunk, line, 1.0);
    chunk.emit_op_u16(Op::LOCAL_GET, data, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc, line);

    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    emit_len(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
    chunks[current].emit_op(Op::F64_DIV, line);
}

/// `geometric_mean(data)` = `exp(mean(ln x))`. Summing logs rather than
/// multiplying keeps a long series from overflowing.
pub fn emit_geometric_mean(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let data = stash_data(&mut chunks[current], line);
    let acc = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);

    core_wasm::f64_const(&mut chunks[current], line, 0.0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    emit_len(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(std::slice::from_mut(chunk), 0, line);

    chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
    chunk.emit_op_u16(Op::LOCAL_GET, data, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let log = chunk.add_import("ecma:math", "log");
    chunk.emit_call(log, 1, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc, line);

    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
    emit_len(chunks, current, data, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    let exp = chunks[current].add_import("ecma:math", "exp");
    chunks[current].emit_call(exp, 1, line);
}
