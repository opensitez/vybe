//! Python `bisect` adapter — bytecode-only.
//!
//! Binary search over a sorted list. There is no ECMA equivalent to route to,
//! so the search is emitted here from ordinary opcodes rather than invented as
//! a host fn. `insort` composes the same search with `ecma:array.splice`.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use vybe_compiler::primitives::instructions::core_wasm;

/// Which end of a run of equal values the insertion point lands on.
#[derive(Clone, Copy, PartialEq)]
pub enum Side {
    Left,
    Right,
}

/// The insertion point for `x` in the sorted `a[lo..hi]`.
///
/// Standard binary search: `Left` stops before an equal run, `Right` after it —
/// the single comparison below is the only difference between them.
/// Stack: `[]` → `[index]`, reading `a`/`x`/`lo`/`hi` from locals.
fn emit_search(
    chunks: &mut [Chunk],
    current: usize,
    a: u16,
    x: u16,
    lo: u16,
    hi: u16,
    side: Side,
    line: u32,
) {
    let mid = chunks[current].alloc_scratch(1);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, lo, line);
    chunk.emit_op_u16(Op::LOCAL_GET, hi, line);
    chunk.emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(std::slice::from_mut(chunk), 0, line);

    // mid = (lo + hi) / 2
    chunk.emit_op_u16(Op::LOCAL_GET, lo, line);
    chunk.emit_op_u16(Op::LOCAL_GET, hi, line);
    chunk.emit_op(Op::I32_ADD, line);
    core_wasm::i32_const(chunk, line, 2);
    chunk.emit_op(Op::I32_DIV_S, line);
    chunk.emit_op_u16(Op::LOCAL_SET, mid, line);

    // left:  a[mid] <  x  → search right half
    // right: a[mid] <= x  → search right half (so equals end up after)
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, mid, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    match side {
        Side::Left => vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line),
        Side::Right => vybe_compiler::primitives::ops::emit_dyn_le(chunk, line),
    }
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, mid, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, lo, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, mid, line);
    chunk.emit_op_u16(Op::LOCAL_SET, hi, line);
    chunk.emit_end(line);

    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lo, line);
}

/// Pop the call's arguments into locals: `(a, x[, lo[, hi]])`. `lo` defaults to
/// 0 and `hi` to `len(a)`.
fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> (u16, u16, u16, u16) {
    let a = chunks[current].alloc_scratch(1);
    let x = chunks[current].alloc_scratch(1);
    let lo = chunks[current].alloc_scratch(1);
    let hi = chunks[current].alloc_scratch(1);

    // Unwind back-to-front — arguments were pushed left to right.
    if argc >= 4 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, hi, line);
    }
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, lo, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, x, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a, line);

    if argc < 3 {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, lo, line);
    }
    if argc < 4 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, a, line);
        vybe_compiler::primitives::collections::emit_len(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, hi, line);
    }
    (a, x, lo, hi)
}

/// `bisect.bisect_left(a, x[, lo[, hi]])`. Stack: `[args…]` → `[index]`.
pub fn emit_bisect_left(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (a, x, lo, hi) = stash_args(chunks, current, argc, line);
    emit_search(chunks, current, a, x, lo, hi, Side::Left, line);
}

/// `bisect.bisect_right(a, x[, lo[, hi]])` / `bisect.bisect`.
pub fn emit_bisect_right(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (a, x, lo, hi) = stash_args(chunks, current, argc, line);
    emit_search(chunks, current, a, x, lo, hi, Side::Right, line);
}

/// `insort_left` / `insort_right` — insert `x` at its search point, keeping `a`
/// sorted. In-place, and returns None like Python does.
fn emit_insort(chunks: &mut [Chunk], current: usize, argc: u8, side: Side, line: u32) {
    let (a, x, lo, hi) = stash_args(chunks, current, argc, line);
    emit_search(chunks, current, a, x, lo, hi, side, line);
    let idx = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);

    // a.splice(idx, 0, x)
    chunks[current].emit_op_u16(Op::LOCAL_GET, a, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    let splice = chunks[current].add_import("ecma:array", "splice");
    chunks[current].emit_call(splice, 4, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_insort_left(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_insort(chunks, current, argc, Side::Left, line);
}

/// `bisect.insort` is `insort_right`.
pub fn emit_insort_right(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_insort(chunks, current, argc, Side::Right, line);
}
