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

use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ValueSource,
};
use vybe_compiler::primitives::functions::create_function_chunk;
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// `len(data)` as f64. Stack: `[]` → `[num]`.
fn emit_len(chunks: &mut [Chunk], current: usize, data: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
}

fn protocol_key(dunder: &str) -> String {
    match crate::protocol::canonical_method(dunder).1 {
        Some(slot) => vybe_ast::protocol_slot_key(slot),
        None => dunder.to_string(),
    }
}

fn emit_slot_as_f64(chunk: &mut Chunk, slot: u16, line: u32) {
    let typeof_fn = chunk.add_import("ecma:value", "typeof");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    let float_key = class_slots::resolve_interned(
        chunk,
        &ClassSlot::internal(protocol_key("__float__")),
        &PlainNames,
    );
    let method = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_call(typeof_fn, 1, line);
    chunk.emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &float_key, Dest::Stack, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method, line);
    chunk.emit_op_u16(Op::LOCAL_GET, method, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, method, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, 1, line);
    chunk.emit_call(to_f64, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_call(to_f64, 1, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_call(to_f64, 1, line);
    chunk.emit_end(line);
}

/// `sum(data)`. Stack: `[]` → `[num]`.
fn emit_sum(chunks: &mut [Chunk], current: usize, data: u16, line: u32) {
    let acc = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let item = chunks[current].alloc_scratch(1);

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

    chunk.emit_op_u16(Op::LOCAL_GET, data, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, item, line);
    chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
    emit_slot_as_f64(chunk, item, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc, line);

    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
}

/// Stash the single list argument into a local. Stack: `[data]` → `[]`.
fn stash_data(chunk: &mut Chunk, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn stash_args(chunk: &mut Chunk, argc: u8, line: u32) -> u16 {
    let base = chunk.alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

fn emit_statistics_error(chunks: &mut [Chunk], current: usize, message: &str, line: u32) {
    chunks[current].emit_string_const(message, line);
    crate::emitter::runtime_adapter::emit_py_raise(
        chunks,
        current,
        1,
        "statistics.StatisticsError",
        line,
    );
}

fn emit_empty_data_guard(chunks: &mut [Chunk], current: usize, data: u16, line: u32) {
    emit_len(chunks, current, data, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_statistics_error(chunks, current, "no data points", line);
    chunks[current].emit_end(line);
}

/// `statistics.mean(data)` / `fmean(data)`. Stack: `[data]` → `[num]`.
pub fn emit_mean(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let data = stash_data(&mut chunks[current], line);
    emit_empty_data_guard(chunks, current, data, line);
    emit_sum(chunks, current, data, line);
    emit_len(chunks, current, data, line);
    chunks[current].emit_op(Op::F64_DIV, line);
}

fn ensure_numeric_cmp_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    const NAME: &str = "__py_statistics_numeric_cmp";
    if let Some(idx) = chunks.iter().position(|chunk| chunk.name == NAME) {
        return idx;
    }
    let idx = chunks.len();
    let mut c = create_function_chunk(NAME, 2);
    c.alloc_scratch(2);
    emit_slot_as_f64(&mut c, 0, line);
    emit_slot_as_f64(&mut c, 1, line);
    c.emit_op(Op::F64_SUB, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    idx
}

/// The data sorted numerically ascending, as a new list. Stack: `[]` → `[array]`.
fn emit_sorted(chunks: &mut Vec<Chunk>, current: usize, data: u16, line: u32) {
    let cmp = ensure_numeric_cmp_chunk(chunks, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, data, line);
    let sorted = chunk.add_import("ecma:array", "toSorted");
    chunk.emit_call(sorted, 1, line);
    chunk.emit_op_u16(Op::REF_FUNC, cmp as u16, line);
    chunk.emit(0, line);
    let _ = chunk;
    vybe_compiler::primitives::collections::emit_sort_with_comparator(chunks, current, line);
}

/// `s[i]` where `i` is an f64-valued local. Stack: `[]` → `[value]`.
fn emit_at(chunk: &mut Chunk, arr: u16, idx: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, arr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

/// `statistics.median(data)` — the middle of the sorted data, or the mean of
/// the middle two when the count is even. Stack: `[data]` → `[num]`.
pub fn emit_median(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let data = stash_data(&mut chunks[current], line);
    let s = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let mid = chunks[current].alloc_scratch(1);

    emit_sorted(chunks, current, data, line);
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
fn emit_median_side(chunks: &mut Vec<Chunk>, current: usize, high: bool, line: u32) {
    let data = stash_data(&mut chunks[current], line);
    let s = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);

    emit_sorted(chunks, current, data, line);
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

pub fn emit_median_low(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_median_side(chunks, current, false, line);
}

pub fn emit_median_high(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
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
    emit_sum(chunks, current, data, line);
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
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_statistics_error(chunks, current, "no mode for empty data", line);
    chunks[current].emit_end(line);
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
pub fn emit_quantiles(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk0 = &mut chunks[current];
    // `n` defaults to 4 (quartiles).
    let nq = chunk0.alloc_scratch(1);
    if argc >= 2 {
        chunk0.emit_op_u16(Op::LOCAL_SET, nq, line);
    } else {
        core_wasm::i32_const(chunk0, line, 4);
        chunk0.emit_op_u16(Op::LOCAL_SET, nq, line);
    }
    chunk0.emit_op_u16(Op::LOCAL_GET, nq, line);
    core_wasm::i32_const(chunk0, line, 1);
    chunk0.emit_op(Op::I32_LT_S, line);
    chunk0.emit_if(line);
    emit_statistics_error(chunks, current, "n must be at least 1", line);
    chunks[current].emit_end(line);
    let data = stash_data(&mut chunks[current], line);
    let s = chunks[current].alloc_scratch(1);
    let ld = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let j = chunks[current].alloc_scratch(1);
    let delta = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);

    emit_sorted(chunks, current, data, line);
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
pub fn emit_median_grouped(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_args(&mut chunks[current], argc, line);
    let data = base;
    let interval = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        let chunk = &mut chunks[current];
        emit_slot_as_f64(chunk, base + 1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, interval, line);
    } else {
        core_wasm::f64_const(&mut chunks[current], line, 1.0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, interval, line);
    }
    let s = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let x = chunks[current].alloc_scratch(1);
    let cf = chunks[current].alloc_scratch(1);
    let f = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);

    emit_sorted(chunks, current, data, line);
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

    // (x - interval/2) + interval * (n/2 - cf) / f
    let chunk = &mut chunks[current];
    emit_slot_as_f64(chunk, x, line);
    chunk.emit_op_u16(Op::LOCAL_GET, interval, line);
    core_wasm::f64_const(chunk, line, 2.0);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_GET, interval, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_op(Op::F64_CONVERT_I32_U, line);
    core_wasm::f64_const(chunk, line, 2.0);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cf, line);
    chunk.emit_op(Op::F64_CONVERT_I32_U, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_GET, f, line);
    chunk.emit_op(Op::F64_CONVERT_I32_U, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_MUL, line);
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

fn emit_pair_moments(
    chunks: &mut [Chunk],
    current: usize,
    line: u32,
) -> (u16, u16, u16, u16, u16, u16) {
    let base = stash_args(&mut chunks[current], 2, line);
    let xs = base;
    let ys = base + 1;
    let n = chunks[current].alloc_scratch(1);
    let mean_x = chunks[current].alloc_scratch(1);
    let mean_y = chunks[current].alloc_scratch(1);
    let sxx = chunks[current].alloc_scratch(1);
    let syy = chunks[current].alloc_scratch(1);
    let sxy = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);

    emit_len(chunks, current, xs, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    emit_sum(chunks, current, xs, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mean_x, line);
    emit_sum(chunks, current, ys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mean_y, line);

    for slot in [sxx, syy, sxy] {
        core_wasm::f64_const(&mut chunks[current], line, 0.0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let dx = chunks[current].alloc_scratch(1);
    let dy = chunks[current].alloc_scratch(1);
    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_op_u16(Op::LOCAL_GET, n, line);
        chunk.emit_op(Op::I32_LT_S, line);
    }
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, xs, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op_u16(Op::LOCAL_GET, mean_x, line);
        chunk.emit_op(Op::F64_SUB, line);
        chunk.emit_op_u16(Op::LOCAL_SET, dx, line);

        chunk.emit_op_u16(Op::LOCAL_GET, ys, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op_u16(Op::LOCAL_GET, mean_y, line);
        chunk.emit_op(Op::F64_SUB, line);
        chunk.emit_op_u16(Op::LOCAL_SET, dy, line);

        chunk.emit_op_u16(Op::LOCAL_GET, sxx, line);
        chunk.emit_op_u16(Op::LOCAL_GET, dx, line);
        chunk.emit_op_u16(Op::LOCAL_GET, dx, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_SET, sxx, line);

        chunk.emit_op_u16(Op::LOCAL_GET, syy, line);
        chunk.emit_op_u16(Op::LOCAL_GET, dy, line);
        chunk.emit_op_u16(Op::LOCAL_GET, dy, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_SET, syy, line);

        chunk.emit_op_u16(Op::LOCAL_GET, sxy, line);
        chunk.emit_op_u16(Op::LOCAL_GET, dx, line);
        chunk.emit_op_u16(Op::LOCAL_GET, dy, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_SET, sxy, line);

        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        core_wasm::i32_const(chunk, line, 1);
        chunk.emit_op(Op::I32_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    }
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    (n, mean_x, mean_y, sxx, syy, sxy)
}

pub fn emit_covariance(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let (n, _, _, _, _, sxy) = emit_pair_moments(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sxy, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op(Op::F64_DIV, line);
}

pub fn emit_correlation(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let (_, _, _, sxx, syy, sxy) = emit_pair_moments(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sxy, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sxx, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, syy, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SQRT, line);
    chunks[current].emit_op(Op::F64_DIV, line);
}

pub fn emit_linear_regression(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let (_, mean_x, mean_y, sxx, _, sxy) = emit_pair_moments(chunks, current, line);
    let slope = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sxy, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sxx, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slope, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slope, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, mean_y, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slope, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, mean_x, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    vybe_compiler::primitives::tuples::emit_tuple(chunks, current, 2, line);
}

pub fn emit_normal_dist(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(&mut chunks[current], argc, line);
    let mu = chunks[current].alloc_scratch(1);
    let sigma = chunks[current].alloc_scratch(1);

    if argc >= 1 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    } else {
        core_wasm::f64_const(&mut chunks[current], line, 0.0);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, mu, line);

    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    } else {
        core_wasm::f64_const(&mut chunks[current], line, 1.0);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, sigma, line);

    class_slots::emit_class_alloc(&mut chunks[current], line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("NormalDist", line);
    let type_slot = class_slots::resolve(&ClassSlot::TypeIdentity, &PlainNames);
    class_slots::emit_class_set(
        &mut chunks[current],
        ObjSource::Stack,
        &type_slot,
        ValueSource::Stack,
        line,
    );
    for (field, slot) in [("mean", mu), ("stdev", sigma)] {
        chunks[current].emit_dup(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        let key = class_slots::resolve_interned(
            &mut chunks[current],
            &ClassSlot::internal(field),
            &PlainNames,
        );
        class_slots::emit_class_set(
            &mut chunks[current],
            ObjSource::Stack,
            &key,
            ValueSource::Stack,
            line,
        );
    }
}
