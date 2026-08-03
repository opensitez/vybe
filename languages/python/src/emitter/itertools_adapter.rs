//! Python `itertools` / `functools` / `operator` adapter — bytecode-only.
//!
//! These are list transforms with no ECMA counterpart to route to
//! (`Array.prototype` has no `product`/`permutations`), so they are emitted
//! here from `ecma:array.*` plus explicit loops. Eager: each returns a list,
//! which `list(...)`/`next(...)`/iteration all accept. `count`/`cycle` are the
//! exception — they are infinite, so they materialise a bounded prefix.
//!
//! No new host fns.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_compiler::primitives::instructions::core_wasm;

/// How many items an infinite generator (`count`, `cycle`) materialises.
/// Bounded because the list is eager; callers take a prefix via `next`/`islice`.
const INFINITE_PREFIX: i32 = 1000;

fn push(chunk: &mut Chunk, line: u32) {
    let p = chunk.add_import("ecma:array", "push");
    chunk.emit_call(p, 2, line);
    chunk.emit_op(Op::DROP, line);
}

fn len_of(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
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
    skip_leading: bool }

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

    // p = truthy(f(xs[i]))
    let p = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, f, line);
    chunk.emit_op_u16(Op::LOCAL_GET, xs, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
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
            skip_leading: false },
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
            skip_leading: false },
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
            skip_leading: true },
        line,
    );
}

/// `itertools.zip_longest(a, b)` — pairs padded with None to the longer input.
/// Stack: `[a, b]` → `[array]`.
pub fn emit_zip_longest(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let b = chunks[current].alloc_scratch(1);
    let a = chunks[current].alloc_scratch(1);
    let na = chunks[current].alloc_scratch(1);
    let nb = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a, line);
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
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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

/// `itertools.islice(iterable, stop)` / `(iterable, start, stop[, step])`.
/// Stack: `[data, …]` → `[array]`.
pub fn emit_islice(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let b = chunk.alloc_scratch(1);
    let a = chunk.alloc_scratch(1);
    let data = chunk.alloc_scratch(1);
    if argc >= 3 {
        chunk.emit_op_u16(Op::LOCAL_SET, b, line);
        chunk.emit_op_u16(Op::LOCAL_SET, a, line);
    } else {
        // islice(it, stop) — one bound, which is the stop.
        chunk.emit_op_u16(Op::LOCAL_SET, b, line);
        core_wasm::i32_const(chunk, line, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, a, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, data, line);
    chunk.emit_op_u16(Op::LOCAL_GET, data, line);
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    let slice = chunk.add_import("ecma:array", "slice");
    chunk.emit_call(slice, 3, line);
}

/// `itertools.accumulate(data)` — running sums. Stack: `[data]` → `[array]`.
pub fn emit_accumulate(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let data = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let acc = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    len_of(chunks, current, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    core_wasm::f64_const(&mut chunks[current], line, 0.0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

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
    chunk.emit_op(Op::F64_ADD, line);
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
    chunk.emit_call(slice, 2, line);
    push(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}
