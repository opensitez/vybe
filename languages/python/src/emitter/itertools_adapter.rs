//! Python `itertools` / `functools` / `operator` adapter — bytecode-only.
//!
//! These are list transforms with no ECMA counterpart to route to
//! (`Array.prototype` has no `product`/`permutations`), so they are emitted
//! here from `ecma:array.*` plus explicit loops. Eager: each returns a list,
//! which `list(...)`/`next(...)`/iteration all accept. `count`/`cycle` are the
//! exception — they are infinite, so they materialise a bounded prefix.
//!
//! No new host fns.

use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, ObjSource, PlainNames, ValueSource,
};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// How many items an infinite generator (`count`, `cycle`) materialises.
/// Bounded because the list is eager; callers take a prefix via `next`/`islice`.
const INFINITE_PREFIX: i32 = 1000;
const FLOAT_ITEMS_TAG: &str = "__py_float_items";

fn push(chunk: &mut Chunk, line: u32) {
    let p = chunk.add_import("ecma:array", "push");
    chunk.emit_call(p, 2, line);
    chunk.emit_op(Op::DROP, line);
}

fn len_of(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
}

fn get_index(chunks: &mut [Chunk], current: usize, slot: u16, index: i32, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    core_wasm::i32_const(&mut chunks[current], line, index);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
}

fn get_index_slot(chunks: &mut [Chunk], current: usize, slot: u16, index: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
}

fn set_index_from_stack(chunks: &mut [Chunk], current: usize, slot: u16, index: u16, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn push_empty_tuple(chunks: &mut [Chunk], current: usize, out: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    vybe_compiler::primitives::tuples::emit_tuple(chunks, current, 0, line);
    push(&mut chunks[current], line);
}

fn push_tuple_from_indices(
    chunks: &mut [Chunk],
    current: usize,
    out: u16,
    data: u16,
    indices: u16,
    r: u16,
    line: u32,
) {
    let tuple = chunks[current].alloc_scratch(1);
    let j = chunks[current].alloc_scratch(1);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, tuple, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);

    let tuple_loop = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tuple, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, indices, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    push(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, tuple_loop, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tuple, line);
    vybe_compiler::primitives::tuples::emit_tag(chunks, current, line);
    push(&mut chunks[current], line);
}

fn array_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let from = chunks[current].add_import("ecma:array", "from");
    chunks[current].emit_call(from, 1, line);
}

// ── operator module ────────────────────────────────────────────────────────
//
// The `operator.*` functions ARE the operators, so each reuses the emit its
// `__py*__` lowering already routes to. The predicates below need one extra
// step: the comparison emits yield an i32, and Python's `bool` is a real
// value, so they lift it the way `materialize_bool_results` does.

/// `operator.truth(x)` → a real `bool`. Stack: `[x]` → `[bool]`.
pub fn emit_op_truth(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}

/// `operator.not_(x)`. Stack: `[x]` → `[bool]`.
pub fn emit_op_not(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}

/// `operator.eq(a, b)` / `ne`. Stack: `[a, b]` → `[bool]`.
pub fn emit_op_eq(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}

pub fn emit_op_ne(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_ne(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}

/// `operator.pos(x)` — unary plus is identity for a number.
pub fn emit_op_pos(_chunks: &mut [Chunk], _current: usize, _argc: u8, _line: u32) {}

/// `operator.abs(x)`. Stack: `[x]` → `[num]`.
pub fn emit_op_abs(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    chunks[current].emit_op(Op::F64_ABS, line);
}

/// `operator.inv(x)` / `invert` — bitwise NOT. Stack: `[x]` → `[num]`.
pub fn emit_op_inv(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let line2 = line;
    vybe_compiler::primitives::expressions::emit_i32_not(&mut chunks[current], line2);
}

/// The bitwise/shift pairs — each is its plain i32 opcode.
fn emit_bin_i32(chunks: &mut [Chunk], current: usize, op: Op, line: u32) {
    chunks[current].emit_op(op, line);
}

pub fn emit_op_and(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_bin_i32(chunks, current, Op::I32_AND, line);
}
pub fn emit_op_or(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_bin_i32(chunks, current, Op::I32_OR, line);
}
pub fn emit_op_xor(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_bin_i32(chunks, current, Op::I32_XOR, line);
}
pub fn emit_op_lshift(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_bin_i32(chunks, current, Op::I32_SHL, line);
}
pub fn emit_op_rshift(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_bin_i32(chunks, current, Op::I32_SHR_S, line);
}

/// `operator.getitem(a, k)` — `a[k]`. Stack: `[a, k]` → `[value]`.
pub fn emit_op_getitem(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
}

/// `operator.setitem(a, k, v)` — `a[k] = v`, returning None.
/// Stack: `[a, k, v]` → `[null]`.
pub fn emit_op_setitem(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `operator.concat(a, b)` — sequence concatenation, so it has to serve both
/// strings and lists. Stack: `[a, b]` → `[value]`.
pub fn emit_op_concat(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let b = chunk.alloc_scratch(1);
    let a = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, b, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a, line);

    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    let is_array = chunk.add_import("ecma:array", "isArray");
    chunk.emit_call(is_array, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    let concat = chunk.add_import("ecma:array", "concat");
    chunk.emit_call(concat, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    let scat = chunk.add_import("wasm:js-string", "concat");
    chunk.emit_call(scat, 2, line);
    chunk.emit_end(line);
}

/// `functools.reduce(f, xs[, init])` IS `Array.prototype.reduce` — but Python
/// takes the function first and the sequence second, so the two arguments
/// swap before the host call. Stack: `[f, xs, init?]` → `[value]`.
pub fn emit_reduce(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let init = chunk.alloc_scratch(1);
    let xs = chunk.alloc_scratch(1);
    let f = chunk.alloc_scratch(1);
    if argc >= 3 {
        chunk.emit_op_u16(Op::LOCAL_SET, init, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, xs, line);
    chunk.emit_op_u16(Op::LOCAL_SET, f, line);

    chunk.emit_op_u16(Op::LOCAL_GET, xs, line);
    chunk.emit_op_u16(Op::LOCAL_GET, f, line);
    let reduce = chunk.add_import("ecma:array", "reduce");
    if argc >= 3 {
        chunk.emit_op_u16(Op::LOCAL_GET, init, line);
        chunk.emit_call(reduce, 3, line);
    } else {
        chunk.emit_call(reduce, 2, line);
    }
}

/// The predicate-driven filters share a shape: walk the list, call `f(x)`, and
/// decide per element. `keep_when` is whether a true predicate keeps the item
/// (`filterfalse` inverts it); `stop_at_first_false` / `skip_leading` select
/// `takewhile` / `dropwhile`.
struct Filter {
    keep_when: bool,
    stop_at_first_false: bool,
    skip_leading: bool,
}

fn emit_pred_filter(chunks: &mut [Chunk], current: usize, spec: Filter, line: u32) {
    let xs = chunks[current].alloc_scratch(1);
    let f = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    // `takewhile` latches off at the first false; `dropwhile` latches on.
    let latch = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, xs, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, f, line);
    len_of(chunks, current, xs, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    core_wasm::i32_const(
        &mut chunks[current],
        line,
        if spec.skip_leading { 1 } else { 0 },
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, latch, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(std::slice::from_mut(chunk), 0, line);

    // p = truthy(f(xs[i])) or truthy(xs[i]) when predicate is None.
    let p = chunk.alloc_scratch(1);
    let item = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, xs, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, item, line);
    chunk.emit_op_u16(Op::LOCAL_GET, f, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, item, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_else(line);
    let recv = vybe_compiler::primitives::callable::push_callback_from_slot(
        std::slice::from_mut(chunk),
        0,
        f,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_GET, item, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, 1 + recv, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, p, line);

    if spec.stop_at_first_false {
        // takewhile: once false, nothing more is taken.
        chunk.emit_op_u16(Op::LOCAL_GET, p, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        core_wasm::i32_const(chunk, line, 1);
        chunk.emit_op_u16(Op::LOCAL_SET, latch, line);
        chunk.emit_end(line);
        chunk.emit_op_u16(Op::LOCAL_GET, latch, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, out, line);
        chunk.emit_op_u16(Op::LOCAL_GET, xs, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        push(chunk, line);
        chunk.emit_end(line);
    } else if spec.skip_leading {
        // dropwhile: drop until the predicate first fails, then keep all.
        chunk.emit_op_u16(Op::LOCAL_GET, latch, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, p, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        core_wasm::i32_const(chunk, line, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, latch, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
        chunk.emit_op_u16(Op::LOCAL_GET, latch, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, out, line);
        chunk.emit_op_u16(Op::LOCAL_GET, xs, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        push(chunk, line);
        chunk.emit_end(line);
    } else {
        chunk.emit_op_u16(Op::LOCAL_GET, p, line);
        if !spec.keep_when {
            chunk.emit_op(Op::I32_EQZ, line);
        }
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, out, line);
        chunk.emit_op_u16(Op::LOCAL_GET, xs, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        push(chunk, line);
        chunk.emit_end(line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `itertools.filterfalse(pred, xs)` — the items the predicate rejects.
pub fn emit_filterfalse(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_pred_filter(
        chunks,
        current,
        Filter {
            keep_when: false,
            stop_at_first_false: false,
            skip_leading: false,
        },
        line,
    );
}

/// `itertools.takewhile(pred, xs)` — the leading run the predicate accepts.
pub fn emit_takewhile(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_pred_filter(
        chunks,
        current,
        Filter {
            keep_when: true,
            stop_at_first_false: true,
            skip_leading: false,
        },
        line,
    );
}

/// `itertools.dropwhile(pred, xs)` — everything from the first rejection on.
pub fn emit_dropwhile(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_pred_filter(
        chunks,
        current,
        Filter {
            keep_when: true,
            stop_at_first_false: false,
            skip_leading: true,
        },
        line,
    );
}

/// `itertools.zip_longest(a, b[, fillvalue])` — pairs padded to the longer input.
/// Stack: `[a, b, fill?]` → `[array]`.
pub fn emit_zip_longest(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let fill = chunks[current].alloc_scratch(1);
    let b = chunks[current].alloc_scratch(1);
    let a = chunks[current].alloc_scratch(1);
    let na = chunks[current].alloc_scratch(1);
    let nb = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, fill, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, fill, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, b, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a, line);
    array_from_slot(chunks, current, a, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a, line);
    array_from_slot(chunks, current, b, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b, line);
    len_of(chunks, current, a, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, na, line);
    len_of(chunks, current, b, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nb, line);

    // n = max(na, nb) — "longest" is the whole point.
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, na, line);
    chunk.emit_op_u16(Op::LOCAL_SET, n, line);
    chunk.emit_op_u16(Op::LOCAL_GET, nb, line);
    chunk.emit_op_u16(Op::LOCAL_GET, na, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, nb, line);
    chunk.emit_op_u16(Op::LOCAL_SET, n, line);
    chunk.emit_end(line);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    for (src, len) in [(a, na), (b, nb)] {
        // Past its end → None, which is what makes this `_longest`.
        chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, fill, line);
        chunks[current].emit_end(line);
    }
    vybe_compiler::primitives::tuples::emit_tuple(chunks, current, 2, line);
    push(&mut chunks[current], line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `itertools.chain(*iterables)` — one list, in order. Stack: `[a, b, …]` → `[array]`.
pub fn emit_chain(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + i, line);
    }
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    let from = chunk.add_import("ecma:array", "from");
    chunk.emit_call(from, 1, line);
    for i in 1..argc as u16 {
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        let concat = chunk.add_import("ecma:array", "concat");
        chunk.emit_call(concat, 2, line);
    }
}

/// `itertools.product(*iterables)` — Cartesian product as tuples.
/// Stack: `[iter0, iter1, …]` → `[array]`.
pub fn emit_product(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let out = chunks[current].alloc_scratch(1);
    let next = chunks[current].alloc_scratch(1);
    let pool = chunks[current].alloc_scratch(1);
    let pool_len = chunks[current].alloc_scratch(1);
    let p = chunks[current].alloc_scratch(1);
    let out_len = chunks[current].alloc_scratch(1);
    let oi = chunks[current].alloc_scratch(1);
    let ii = chunks[current].alloc_scratch(1);
    let prefix = chunks[current].alloc_scratch(1);
    let tuple = chunks[current].alloc_scratch(1);
    let pools = chunks[current].alloc_scratch(1);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, argc as u16, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pools, line);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    vybe_compiler::primitives::tuples::emit_tag(chunks, current, line);
    push(&mut chunks[current], line);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, p, line);
    let pools_loop = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
    core_wasm::i32_const(&mut chunks[current], line, argc as i32);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, pools, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    let from = chunks[current].add_import("ecma:array", "from");
    chunks[current].emit_call(from, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pool, line);
    len_of(chunks, current, pool, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pool_len, line);
    len_of(chunks, current, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_len, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, next, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, oi, line);

    let outer = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, oi, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, oi, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, prefix, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ii, line);

    let inner = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ii, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pool_len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, prefix, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::expressions::emit_undefined(&mut chunks[current], line);
    let slice = chunks[current].add_import("ecma:array", "slice");
    chunks[current].emit_call(slice, 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, tuple, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tuple, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pool, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ii, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    push(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, next, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tuple, line);
    vybe_compiler::primitives::tuples::emit_tag(chunks, current, line);
    push(&mut chunks[current], line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, ii, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ii, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, inner, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, oi, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, oi, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, outer, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, next, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, p, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, pools_loop, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `itertools.combinations(iterable, r)` — lexicographic r-length subsequences.
/// Stack: `[iterable, r]` → `[array-of-tuples]`.
pub fn emit_combinations(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let r = chunks[current].alloc_scratch(1);
    let data = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let indices = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let j = chunks[current].alloc_scratch(1);
    let prev = chunks[current].alloc_scratch(1);
    let found = chunks[current].alloc_scratch(1);
    let found_idx = chunks[current].alloc_scratch(1);
    let done = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, r, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    array_from_slot(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    len_of(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    push_empty_tuple(chunks, current, out, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, indices, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let init = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, indices, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    push(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, init, line);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, done, line);
    let outer = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, done, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    push_tuple_from_indices(chunks, current, out, data, indices, r, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let scan = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_GE_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, found, line);
    chunks[current].emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_else(line);
    get_index_slot(chunks, current, indices, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op(Op::I32_NE, line);
    chunks[current].emit_if(line);
    get_index_slot(chunks, current, indices, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    set_index_from_stack(chunks, current, indices, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found_idx, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, scan, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, found, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, done, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, found_idx, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    let tail = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, prev, line);
    get_index_slot(chunks, current, indices, prev, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    set_index_from_stack(chunks, current, indices, j, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, tail, line);
    chunks[current].emit_end(line);

    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, outer, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `itertools.combinations_with_replacement(iterable, r)`.
/// Stack: `[iterable, r]` → `[array-of-tuples]`.
pub fn emit_combinations_with_replacement(
    chunks: &mut [Chunk],
    current: usize,
    _argc: u8,
    line: u32,
) {
    let r = chunks[current].alloc_scratch(1);
    let data = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let indices = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let j = chunks[current].alloc_scratch(1);
    let found = chunks[current].alloc_scratch(1);
    let found_idx = chunks[current].alloc_scratch(1);
    let done = chunks[current].alloc_scratch(1);
    let next_val = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, r, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    array_from_slot(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    len_of(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    push_empty_tuple(chunks, current, out, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, indices, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let init = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, indices, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    push(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, init, line);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, done, line);
    let outer = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, done, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    push_tuple_from_indices(chunks, current, out, data, indices, r, line);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let scan = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_GE_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, found, line);
    chunks[current].emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_else(line);
    get_index_slot(chunks, current, indices, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op(Op::I32_NE, line);
    chunks[current].emit_if(line);
    get_index_slot(chunks, current, indices, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, next_val, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found_idx, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, scan, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, found, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, done, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, found_idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    let fill = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, next_val, line);
    set_index_from_stack(chunks, current, indices, j, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, fill, line);
    chunks[current].emit_end(line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, outer, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `itertools.permutations(iterable[, r])` — CPython's cycles algorithm.
/// Stack: `[iterable, r?]` → `[array-of-tuples]`.
pub fn emit_permutations(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let r = chunks[current].alloc_scratch(1);
    let data = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let indices = chunks[current].alloc_scratch(1);
    let cycles = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let j = chunks[current].alloc_scratch(1);
    let done = chunks[current].alloc_scratch(1);
    let advanced = chunks[current].alloc_scratch(1);
    let tmp = chunks[current].alloc_scratch(1);
    let swap_idx = chunks[current].alloc_scratch(1);

    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, r, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    array_from_slot(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    len_of(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    if argc < 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, r, line);
    }
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    push_empty_tuple(chunks, current, out, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, indices, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let init_idx = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, indices, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    push(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, init_idx, line);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cycles, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let init_cycles = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cycles, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    push(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, init_cycles, line);

    push_tuple_from_indices(chunks, current, out, data, indices, r, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, done, line);
    let outer = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, done, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, advanced, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let scan = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_GE_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, advanced, line);
    chunks[current].emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_else(line);

    get_index_slot(chunks, current, cycles, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    set_index_from_stack(chunks, current, cycles, i, line);
    get_index_slot(chunks, current, cycles, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);

    get_index_slot(chunks, current, indices, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, tmp, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    let rotate = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, swap_idx, line);
    get_index_slot(chunks, current, indices, swap_idx, line);
    set_index_from_stack(chunks, current, indices, j, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, rotate, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tmp, line);
    set_index_from_stack(chunks, current, indices, j, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    set_index_from_stack(chunks, current, cycles, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    get_index_slot(chunks, current, cycles, i, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, swap_idx, line);
    get_index_slot(chunks, current, indices, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, tmp, line);
    get_index_slot(chunks, current, indices, swap_idx, line);
    set_index_from_stack(chunks, current, indices, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tmp, line);
    set_index_from_stack(chunks, current, indices, swap_idx, line);
    push_tuple_from_indices(chunks, current, out, data, indices, r, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, advanced, line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, scan, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, advanced, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, done, line);
    chunks[current].emit_end(line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, outer, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `itertools.chain.from_iterable(iterables)` — flatten one level.
/// Stack: `[nested]` → `[array]`.
pub fn emit_chain_from_iterable(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let nested = chunks[current].alloc_scratch(1);
    let outer_len = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let part = chunks[current].alloc_scratch(1);
    let inner_len = chunks[current].alloc_scratch(1);
    let j = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, nested, line);
    array_from_slot(chunks, current, nested, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nested, line);
    len_of(chunks, current, nested, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, outer_len, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let outer = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, outer_len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, nested, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, part, line);
    array_from_slot(chunks, current, part, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, part, line);
    len_of(chunks, current, part, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, inner_len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);

    let inner = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, inner_len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, part, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    push(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, inner, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, outer, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `itertools.compress(data, selectors)` — keep items whose selector is true.
/// Stack: `[data, selectors]` → `[array]`.
pub fn emit_compress(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let selectors = chunks[current].alloc_scratch(1);
    let data = chunks[current].alloc_scratch(1);
    let n_data = chunks[current].alloc_scratch(1);
    let n_sel = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, selectors, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    array_from_slot(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    array_from_slot(chunks, current, selectors, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, selectors, line);
    len_of(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n_data, line);
    len_of(chunks, current, selectors, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n_sel, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, n_data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_sel, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_data, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_sel, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    chunks[current].emit_end(line);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, selectors, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    push(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `itertools.starmap(f, iterable)` — call `f(*args)` for every argument row.
/// Stack: `[f, iterable]` → `[array]`.
pub fn emit_starmap(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let data = chunks[current].alloc_scratch(1);
    let func = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let row = chunks[current].alloc_scratch(1);
    let row_len = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, func, line);
    array_from_slot(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    len_of(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, row, line);
    len_of(chunks, current, row, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, row_len, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, row_len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    let recv =
        vybe_compiler::primitives::callable::push_callback_from_slot(chunks, current, func, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(
        &mut chunks[current],
        recv,
        line,
    );
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, row_len, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    let recv =
        vybe_compiler::primitives::callable::push_callback_from_slot(chunks, current, func, line);
    get_index(chunks, current, row, 0, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(
        &mut chunks[current],
        1 + recv,
        line,
    );
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, row_len, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    let recv =
        vybe_compiler::primitives::callable::push_callback_from_slot(chunks, current, func, line);
    get_index(chunks, current, row, 0, line);
    get_index(chunks, current, row, 1, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(
        &mut chunks[current],
        2 + recv,
        line,
    );
    chunks[current].emit_else(line);
    let recv =
        vybe_compiler::primitives::callable::push_callback_from_slot(chunks, current, func, line);
    get_index(chunks, current, row, 0, line);
    get_index(chunks, current, row, 1, line);
    get_index(chunks, current, row, 2, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(
        &mut chunks[current],
        3 + recv,
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    push(&mut chunks[current], line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

fn emit_group_pair(
    chunks: &mut [Chunk],
    current: usize,
    out: u16,
    key: u16,
    group: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, group, line);
    vybe_compiler::primitives::tuples::emit_tuple(chunks, current, 2, line);
    push(&mut chunks[current], line);
}

/// `itertools.groupby(data[, key])` — consecutive groups as `(key, group)` tuples.
/// Stack: `[data, key?]` → `[array]`.
pub fn emit_groupby(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let key_func = chunks[current].alloc_scratch(1);
    let data = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let group = chunks[current].alloc_scratch(1);
    let cur_key = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let item = chunks[current].alloc_scratch(1);
    let started = chunks[current].alloc_scratch(1);

    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, key_func, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, key_func, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    array_from_slot(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    len_of(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, group, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, started, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_func, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
    chunks[current].emit_else(line);
    let recv =
        vybe_compiler::primitives::callable::push_callback_from_slot(chunks, current, key_func, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(
        &mut chunks[current],
        1 + recv,
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, started, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_key, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_group_pair(chunks, current, out, cur_key, group, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, group, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur_key, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur_key, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, started, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, group, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
    push(&mut chunks[current], line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, started, line);
    chunks[current].emit_if(line);
    emit_group_pair(chunks, current, out, cur_key, group, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `itertools.repeat(x, n)`. Stack: `[x, n]` → `[array]`.
pub fn emit_repeat(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let n = chunks[current].alloc_scratch(1);
    let x = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, INFINITE_PREFIX);
        chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, x, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(std::slice::from_mut(chunk), 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    push(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `itertools.count(start=0, step=1)` — infinite, so a bounded prefix.
/// Stack: `[start?, step?]` → `[array]`.
pub fn emit_count(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let step = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let cur = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);

    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, step, line);
    } else {
        core_wasm::f64_const(&mut chunks[current], line, 1.0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, step, line);
    }
    if argc >= 1 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
    } else {
        core_wasm::f64_const(&mut chunks[current], line, 0.0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, INFINITE_PREFIX);
    chunk.emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(std::slice::from_mut(chunk), 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cur, line);
    push(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cur, line);
    chunk.emit_op_u16(Op::LOCAL_GET, step, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cur, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// Float-origin `itertools.count`; same sequence as `count`, with a Python repr
/// tag so integral float values display as `1.0`, matching CPython.
pub fn emit_count_float(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_count(chunks, current, argc, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(true, line);
    let slot = class_slots::resolve_interned(
        &mut chunks[current],
        &ClassSlot::internal(FLOAT_ITEMS_TAG),
        &PlainNames,
    );
    class_slots::emit_class_set(
        &mut chunks[current],
        ObjSource::Stack,
        &slot,
        ValueSource::Stack,
        line,
    );
}

/// `itertools.cycle(iterable)` — infinite, so a bounded prefix of repeats.
/// Stack: `[data]` → `[array]`.
pub fn emit_cycle(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let data = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    len_of(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, INFINITE_PREFIX);
    chunk.emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(std::slice::from_mut(chunk), 0, line);
    // out.push(data[i % n]) — wrapping is what makes it a cycle.
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, data, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_op(Op::I32_REM_S, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    push(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

fn emit_islice_array_loop(
    chunks: &mut [Chunk],
    current: usize,
    data: u16,
    start: u16,
    stop: u16,
    step: u16,
    line: u32,
) {
    let i = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    array_from_slot(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stop, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    push(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, step, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `itertools.islice(iterable, stop)` / `(iterable, start, stop[, step])`.
/// Stack: `[data, …]` → `[array]`.
pub fn emit_islice(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let step = chunks[current].alloc_scratch(1);
    let stop = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    let data = chunks[current].alloc_scratch(1);
    if argc >= 4 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, step, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, step, line);
    }
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, stop, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
    } else {
        // islice(it, stop) — one bound, which is the stop.
        chunks[current].emit_op_u16(Op::LOCAL_SET, stop, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);

    if argc == 2 {
        // Generator continuations are not arrays, and host `Array.slice` cannot
        // resume them. Keep the bounded generator form on the shared generator
        // primitive so `islice(count(), n)` never drains an infinite source.
        chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
        let is_gen = chunks[current].add_import("ecma:value", "isGenerator");
        chunks[current].emit_call(is_gen, 1, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, stop, line);
        vybe_compiler::primitives::generators::emit_take_into_array(chunks, current, line);
        chunks[current].emit_else(line);
        emit_islice_array_loop(chunks, current, data, start, stop, step, line);
        chunks[current].emit_end(line);
    } else {
        emit_islice_array_loop(chunks, current, data, start, stop, step, line);
    }
}

/// Consuming form used for `islice(iter(seq), ...)`. Python iterators advance as
/// they are sliced; regular sequences do not, so the walker only routes known
/// iterator variables here.
pub fn emit_islice_consume(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let step = chunks[current].alloc_scratch(1);
    let stop = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    let data = chunks[current].alloc_scratch(1);
    let prefix = chunks[current].alloc_scratch(1);
    let limit = chunks[current].alloc_scratch(1);

    if argc >= 4 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, step, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, step, line);
    }
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, stop, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_SET, stop, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    let is_gen = chunks[current].add_import("ecma:value", "isGenerator");
    chunks[current].emit_call(is_gen, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stop, line);
    vybe_compiler::primitives::generators::emit_take_into_array(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, prefix, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stop, line);
    let splice = chunks[current].add_import("ecma:array", "splice");
    chunks[current].emit_call(splice, 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, prefix, line);
    chunks[current].emit_end(line);

    len_of(chunks, current, prefix, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, limit, line);
    emit_islice_array_loop(chunks, current, prefix, start, limit, step, line);
}

/// `itertools.accumulate(data[, func[, initial]])`.
/// Stack: `[data, func?, initial?]` → `[array]`.
pub fn emit_accumulate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let initial = chunks[current].alloc_scratch(1);
    let func = chunks[current].alloc_scratch(1);
    let data = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let acc = chunks[current].alloc_scratch(1);
    let item = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, initial, line);
    }
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, func, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, func, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    array_from_slot(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    len_of(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, initial, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, acc, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
        push(&mut chunks[current], line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op(Op::I32_GT_S, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, acc, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
        push(&mut chunks[current], line);
        chunks[current].emit_end(line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    }

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
    chunk.emit_op_u16(Op::LOCAL_GET, func, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
    chunk.emit_op_u16(Op::LOCAL_GET, item, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_else(line);
    let recv = vybe_compiler::primitives::callable::push_callback_from_slot(
        std::slice::from_mut(chunk),
        0,
        func,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
    chunk.emit_op_u16(Op::LOCAL_GET, item, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, 2 + recv, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
    push(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `itertools.pairwise(data)` → `[(d0,d1), (d1,d2), …]`. Stack: `[data]` → `[array]`.
pub fn emit_pairwise(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let data = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    len_of(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    vybe_compiler::primitives::tuples::emit_tuple(chunks, current, 2, line);
    push(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `itertools.batched(data, n)` → lists of up to `n`. Stack: `[data, n]` → `[array]`.
pub fn emit_batched(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let size = chunks[current].alloc_scratch(1);
    let data = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, size, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    len_of(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(std::slice::from_mut(chunk), 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, data, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, size, line);
    chunk.emit_op(Op::I32_ADD, line);
    let slice = chunk.add_import("ecma:array", "slice");
    chunk.emit_call(slice, 3, line);
    push(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, size, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `itertools.tee(data, n)` → `n` independent copies. Stack: `[data, n?]` → `[array]`.
pub fn emit_tee(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let n = chunks[current].alloc_scratch(1);
    let data = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 2);
        chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    let is_gen = chunks[current].add_import("ecma:value", "isGenerator");
    chunks[current].emit_call(is_gen, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    core_wasm::i32_const(&mut chunks[current], line, INFINITE_PREFIX);
    vybe_compiler::primitives::generators::emit_take_into_array(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    chunks[current].emit_else(line);
    array_from_slot(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    chunks[current].emit_end(line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(std::slice::from_mut(chunk), 0, line);
    // Each copy must be independent — `tee`'s whole purpose.
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, data, line);
    core_wasm::i32_const(chunk, line, 0);
    let slice = chunk.add_import("ecma:array", "slice");
    // §23.1.3.27 `end` = undefined means `len`; passed explicitly so the
    // import has ONE arity (a WASM import cannot have an optional argument).
    vybe_compiler::primitives::expressions::emit_undefined(chunk, line);
    chunk.emit_call(slice, 3, line);
    push(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}
