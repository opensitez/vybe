//! Shared slice emission for every Vybe language.
//!
//! Surface slice syntax differs across languages (Python `a[i:j:k]`,
//! C# `a[i..j]`, Fortran `a(i:j:k)`, JS `.slice`), but they all reduce to
//! one canonical operation: given a sequence and `(start, stop, step)`
//! bounds, produce a new sequence. This module is the single home for that
//! lowering so every front-end emits compatible bytecode.
//!
//! Two entry points:
//! - [`emit_contiguous`] — no-step `[obj, start, end] -> result`, delegates
//!   to the polymorphic `ecma:array.slice` host fn (string -> substring,
//!   array -> array, with negative-index wrap + clamp).
//! - [`emit_stepped`] — `[obj, lower|null, upper|null, step|null] -> result`,
//!   the full CPython `PySlice_AdjustIndices` normalization + strided copy,
//!   emitted inline from shared loop/collection/ops helpers.
//!
//! Per-language quirks (negative-index wrap, `step == 0` policy) are captured
//! by [`Options`]; [`Options::for_language`] holds the quirk table.

use crate::compiler::instructions::core_wasm;
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

use crate::compiler::collections;
use crate::compiler::errors;
use crate::compiler::loops;
use crate::compiler::ops;
use crate::compiler::tuples;

/// Behavior when the slice `step` evaluates to zero.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ZeroStep {
    /// Yield an empty sequence (lenient — the retired `__vybe_slicestep`
    /// helper's behavior, kept as the default for non-Python front-ends).
    Empty,
    /// Raise the language value-error (Python: `ValueError: slice step
    /// cannot be zero`). The concrete type name is [`Options::value_error`].
    Raise,
}

/// Per-language slice semantics. Bounds always wrap from the end and clamp
/// (the canonical, ecma-`slice`-compatible form every front-end already gets
/// on the no-step path); the one genuine divergence is the `step == 0` policy,
/// so front-ends build this from their profile properties — never a name check.
#[derive(Clone, Copy)]
pub struct Options {
    /// `step == 0` policy.
    pub zero_step: ZeroStep,
    /// Exception type raised when `zero_step == Raise`. Normalized through
    /// `errors::canonical_exception_name`, so e.g. Python `ValueError` and
    /// Dart `FormatException` unify for cross-language `catch`.
    pub value_error: &'static str,
}

impl Options {
    /// Build from the `slice_step_zero_raises` profile property. Property-
    /// driven, so a language opts into Python-style `ValueError` by setting
    /// the flag in its profile — no per-language branching here.
    pub const fn new(zero_step_raises: bool) -> Options {
        Options {
            zero_step: if zero_step_raises {
                ZeroStep::Raise
            } else {
                ZeroStep::Empty
            },
            value_error: "ValueError",
        }
    }
}

/// No-step slice. Stack: `[obj, start, end] -> [result]`.
///
/// Routes through `ecma:array.slice`, which dispatches string->substring and
/// array->array internally with negative-index wrap and out-of-range clamp
/// (ECMA-262 §23.1.3.28). This is the canonical contiguous slice for every
/// language; the retired `__vybe_slice` helper had the same contract.
pub fn emit_contiguous(chunks: &mut [Chunk], current: usize, line: u32) {
    // Stash [obj, start, end] so the source is reachable after the slice.
    let base = chunks[current].alloc_scratch(3);
    let obj = base;
    let start = base + 1;
    let end = base + 2;
    set(chunks, current, end, line);
    set(chunks, current, start, line);
    set(chunks, current, obj, line);

    // A typed array (bytes/bytearray = Uint8Array) is a real JS type, not a
    // tagged array: slice it through the typed-array slice so the result keeps
    // its type. Everything else (array/string) goes through ecma:array.slice.
    get(chunks, current, obj, line);
    let is_view = chunks[current].add_import("ecma:arraybuffer", "isView");
    chunks[current].emit_call(is_view, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(chunks, current, obj, line);
    get(chunks, current, start, line);
    get(chunks, current, end, line);
    let ta_slice = chunks[current].add_import("ecma:uint8array", "slice");
    chunks[current].emit_call(ta_slice, 3, line);
    chunks[current].emit_else(line);
    get(chunks, current, obj, line);
    get(chunks, current, start, line);
    get(chunks, current, end, line);
    collections::emit_slice(chunks, current, line);
    chunks[current].emit_end(line); // [result]

    get(chunks, current, obj, line); // [result, obj]
    tuples::emit_propagate_tag(chunks, current, line); // [result]
}

// ── local-slot helpers ──────────────────────────────────────────────────

fn get(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// Sentinel bound value: emit `len` at runtime.
const LEN: i32 = i32::MIN;
/// Sentinel bound value: emit `len - 1` at runtime.
const LEN_M1: i32 = i32::MIN + 1;

/// Push a bound value: [`LEN`]/[`LEN_M1`] sentinels resolve to `len`/`len-1`;
/// any other `v` is a literal. Keeps the `-1` literal (a real Python default)
/// distinct from the `len-1` default.
fn emit_bound_const(chunks: &mut [Chunk], current: usize, len: u16, v: i32, line: u32) {
    match v {
        LEN => get(chunks, current, len, line),
        LEN_M1 => {
            get(chunks, current, len, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_SUB, line);
        }
        _ => core_wasm::i32_const(&mut chunks[current], line, v),
    }
}

/// Push i32 bool: `slot < 0`.
fn is_neg(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    get(chunks, current, slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
}

/// Push i32 bool: `obj` is a string (via `wasm:js-string.test`, yields i32).
fn is_string(chunks: &mut [Chunk], current: usize, obj: u16, line: u32) {
    get(chunks, current, obj, line);
    let idx = chunks[current].add_import("wasm:js-string", "test");
    chunks[current].emit_call(idx, 1, line);
}

/// Emit `dst = neg ? a : b`, where `a`/`b` are provided as closures that push
/// one value. Uses step sign as the selector when `neg` names the step slot.
fn set_by_step_sign(
    chunks: &mut [Chunk],
    current: usize,
    step: u16,
    dst: u16,
    line: u32,
    neg_branch: impl Fn(&mut [Chunk], usize, u32),
    pos_branch: impl Fn(&mut [Chunk], usize, u32),
) {
    is_neg(chunks, current, step, line);
    chunks[current].emit_if(line);
    neg_branch(chunks, current, line);
    set(chunks, current, dst, line);
    chunks[current].emit_else(line);
    pos_branch(chunks, current, line);
    set(chunks, current, dst, line);
    chunks[current].emit_end(line);
}

/// Normalize one bound (`raw` -> `dst`) per CPython `PySlice_AdjustIndices`.
/// `default_neg`/`default_pos` supply the value when `raw` is NULL (absent),
/// chosen by step sign. `clamp_low_neg`/`clamp_low_pos` supply the value when
/// the bound underflows (`< 0`); overflow (`>= len`) clamps to `len-1`/`len`.
#[allow(clippy::too_many_arguments)]
fn normalize_bound(
    chunks: &mut [Chunk],
    current: usize,
    raw: u16,
    dst: u16,
    step: u16,
    len: u16,
    line: u32,
    default_neg: i32,
    default_pos: i32,
    clamp_low_neg: i32,
) {
    // if raw is NULL -> default by step sign
    get(chunks, current, raw, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    set_by_step_sign(
        chunks,
        current,
        step,
        dst,
        line,
        move |c, cur, l| emit_bound_const(c, cur, len, default_neg, l),
        move |c, cur, l| emit_bound_const(c, cur, len, default_pos, l),
    );
    chunks[current].emit_else(line);
    // provided: dst = raw
    get(chunks, current, raw, line);
    set(chunks, current, dst, line);
    // negative wrap: if dst < 0 { dst += len } — canonical, ecma-slice parity.
    is_neg(chunks, current, dst, line);
    chunks[current].emit_if(line);
    get(chunks, current, dst, line);
    get(chunks, current, len, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(chunks, current, dst, line);
    chunks[current].emit_end(line);
    // clamp low: if dst < 0 { dst = (step<0) ? clamp_low_neg : 0 }
    is_neg(chunks, current, dst, line);
    chunks[current].emit_if(line);
    set_by_step_sign(
        chunks,
        current,
        step,
        dst,
        line,
        move |c, cur, l| core_wasm::i32_const(&mut c[cur], l, clamp_low_neg),
        |c, cur, l| core_wasm::i32_const(&mut c[cur], l, 0),
    );
    chunks[current].emit_else(line);
    // clamp high: else if dst >= len { dst = (step<0) ? len-1 : len }
    get(chunks, current, dst, line);
    get(chunks, current, len, line);
    ops::emit_dyn_ge(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    set_by_step_sign(
        chunks,
        current,
        step,
        dst,
        line,
        move |c, cur, l| {
            get(c, cur, len, l);
            core_wasm::i32_const(&mut c[cur], l, 1);
            c[cur].emit_op(Op::I32_SUB, l); // len - 1
        },
        move |c, cur, l| get(c, cur, len, l),
    );
    chunks[current].emit_end(line); // end (dst >= len)
    chunks[current].emit_end(line); // end (dst < 0)
    chunks[current].emit_end(line); // end (raw NULL?)
}

/// Fill the `step_n`/`lo`/`hi` slots from the raw `lower_in`/`upper_in`/
/// `step_in` slots and `len`, per CPython `PySlice_AdjustIndices`. Shared by
/// strided read, assignment, and deletion. `opts.zero_step` is honored (Raise
/// emits a throw; Empty forces an empty range).
#[allow(clippy::too_many_arguments)]
fn normalize_slice_into(
    chunks: &mut [Chunk],
    current: usize,
    lower_in: u16,
    upper_in: u16,
    step_in: u16,
    len: u16,
    step_n: u16,
    lo: u16,
    hi: u16,
    opts: Options,
    line: u32,
) {
    // step_n = step_in ?? 1
    get(chunks, current, step_in, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    set(chunks, current, step_n, line);
    chunks[current].emit_else(line);
    get(chunks, current, step_in, line);
    set(chunks, current, step_n, line);
    chunks[current].emit_end(line);

    // step == 0 policy (Raise diverges; Empty handled after normalization).
    if opts.zero_step == ZeroStep::Raise {
        get(chunks, current, step_n, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        ops::emit_dyn_eq(&mut chunks[current], line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
        chunks[current].emit_dup(line);
        chunks[current].emit_string_const("slice step cannot be zero", line);
        errors::emit_exception_new_finalize(&mut chunks[current], opts.value_error, line);
        errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_end(line);
    }

    // lo default: step<0 ? len-1 : 0 ; clamp low: step<0 ? -1 : 0
    normalize_bound(
        chunks, current, lower_in, lo, step_n, len, line, LEN_M1, 0, -1,
    );
    // hi default: step<0 ? -1 : len ; clamp low: step<0 ? -1 : 0
    normalize_bound(
        chunks, current, upper_in, hi, step_n, len, line, -1, LEN, -1,
    );

    // Empty policy: if step still 0, force an empty range (lo=hi=0, step=1).
    if opts.zero_step == ZeroStep::Empty {
        get(chunks, current, step_n, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        ops::emit_dyn_eq(&mut chunks[current], line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        set(chunks, current, lo, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        set(chunks, current, hi, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        set(chunks, current, step_n, line);
        chunks[current].emit_end(line);
    }
}

/// Strided slice. Stack: `[obj, lower, upper, step] -> [result]`, where any of
/// `lower`/`upper`/`step` may be NULL (absent bound). Emits inline CPython
/// slice normalization + strided copy; string operands rejoin to a string.
pub fn emit_stepped(chunks: &mut [Chunk], current: usize, line: u32, opts: Options) {
    let base = chunks[current].alloc_scratch(11);
    let obj = base;
    let lower_in = base + 1;
    let upper_in = base + 2;
    let step_in = base + 3;
    let len = base + 4;
    let step_n = base + 5;
    let lo = base + 6;
    let hi = base + 7;
    let i = base + 8;
    let result = base + 9;
    let cond = base + 10;

    // Pop [obj, lower, upper, step] (step on top).
    set(chunks, current, step_in, line);
    set(chunks, current, upper_in, line);
    set(chunks, current, lower_in, line);
    set(chunks, current, obj, line);

    // result = []
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, result, line);

    // len = length(obj)
    get(chunks, current, obj, line);
    collections::emit_len(chunks, current, line);
    set(chunks, current, len, line);

    normalize_slice_into(
        chunks, current, lower_in, upper_in, step_in, len, step_n, lo, hi, opts, line,
    );

    // i = lo
    get(chunks, current, lo, line);
    set(chunks, current, i, line);

    // while ( (step>0 && i<hi) || (step<0 && i>hi) ) { result.push(obj[i]); i += step }
    let st = loops::emit_loop_start(chunks, current, line);
    // cond = (step_n > 0) ? (i < hi) : (i > hi)
    get(chunks, current, step_n, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_gt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(chunks, current, i, line);
    get(chunks, current, hi, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    set(chunks, current, cond, line);
    chunks[current].emit_else(line);
    get(chunks, current, i, line);
    get(chunks, current, hi, line);
    ops::emit_dyn_gt(&mut chunks[current], line);
    set(chunks, current, cond, line);
    chunks[current].emit_end(line);
    get(chunks, current, cond, line);
    loops::emit_loop_cond(chunks, current, line); // break when false

    // result.push( is_string(obj) ? charAt(obj,i) : obj[i] )
    get(chunks, current, result, line);
    is_string(chunks, current, obj, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, obj, line);
    get(chunks, current, i, line);
    {
        let idx = chunks[current].add_import("ecma:string", "charAt");
        chunks[current].emit_call(idx, 2, line);
    }
    chunks[current].emit_else(line);
    get(chunks, current, obj, line);
    get(chunks, current, i, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // i += step_n
    get(chunks, current, i, line);
    get(chunks, current, step_n, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(chunks, current, i, line);

    loops::emit_loop_end(chunks, current, st, line);

    // Rejoin strings: is_string(obj) ? result.join("") : result
    is_string(chunks, current, obj, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, result, line);
    chunks[current].emit_string_const("", line);
    collections::emit_join(chunks, current, line);
    chunks[current].emit_else(line);
    get(chunks, current, result, line);
    chunks[current].emit_end(line);

    // Preserve tuple-ness: tuple[i:j:k] stays a tuple (no-op for lists/strings).
    get(chunks, current, obj, line);
    tuples::emit_propagate_tag(chunks, current, line);
}

/// Push the loop condition `(step>0 ? i<hi : i>hi)` as a dyn bool. Stack: `-> [bool]`.
fn emit_stride_cond(chunks: &mut [Chunk], current: usize, step_n: u16, i: u16, hi: u16, line: u32) {
    get(chunks, current, step_n, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_gt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(chunks, current, i, line);
    get(chunks, current, hi, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(chunks, current, i, line);
    get(chunks, current, hi, line);
    ops::emit_dyn_gt(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

/// Contiguous slice assignment `obj[lo:hi] = value` (variable length: splice).
/// Stack: `[obj, lower|null, upper|null, value] -> []`. Bounds wrap from the end
/// and clamp; the array is mutated in place (remove the range, insert `value`).
pub fn emit_splice_assign(chunks: &mut [Chunk], current: usize, line: u32) {
    let b = chunks[current].alloc_scratch(8);
    let obj = b;
    let lower_in = b + 1;
    let upper_in = b + 2;
    let value = b + 3;
    let len = b + 4;
    let one = b + 5;
    let lo = b + 6;
    let hi = b + 7;

    set(chunks, current, value, line);
    set(chunks, current, upper_in, line);
    set(chunks, current, lower_in, line);
    set(chunks, current, obj, line);

    get(chunks, current, obj, line);
    collections::emit_len(chunks, current, line);
    set(chunks, current, len, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    set(chunks, current, one, line);

    // Contiguous (step = 1): reuse the CPython bound normalization.
    normalize_bound(chunks, current, lower_in, lo, one, len, line, LEN_M1, 0, -1);
    normalize_bound(chunks, current, upper_in, hi, one, len, line, -1, LEN, -1);

    // count = max(hi - lo, 0)  (stored back into hi)
    get(chunks, current, hi, line);
    get(chunks, current, lo, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    set(chunks, current, hi, line);
    get(chunks, current, hi, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(chunks, current, hi, line);
    chunks[current].emit_end(line);

    // remove_range(obj, lo, count)
    get(chunks, current, obj, line);
    get(chunks, current, lo, line);
    get(chunks, current, hi, line);
    collections::emit_remove_range(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // insert_range(obj, lo, value)
    get(chunks, current, obj, line);
    get(chunks, current, lo, line);
    get(chunks, current, value, line);
    collections::emit_insert_range(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Extended slice assignment `obj[lo:hi:step] = value` (positional: each strided
/// slot gets the matching `value` element). Stack:
/// `[obj, lower|null, upper|null, step|null, value] -> []`.
pub fn emit_strided_assign(chunks: &mut [Chunk], current: usize, line: u32, opts: Options) {
    let b = chunks[current].alloc_scratch(11);
    let obj = b;
    let lower_in = b + 1;
    let upper_in = b + 2;
    let step_in = b + 3;
    let value = b + 4;
    let len = b + 5;
    let step_n = b + 6;
    let lo = b + 7;
    let hi = b + 8;
    let i = b + 9;
    let m = b + 10;

    set(chunks, current, value, line);
    set(chunks, current, step_in, line);
    set(chunks, current, upper_in, line);
    set(chunks, current, lower_in, line);
    set(chunks, current, obj, line);

    get(chunks, current, obj, line);
    collections::emit_len(chunks, current, line);
    set(chunks, current, len, line);

    normalize_slice_into(
        chunks, current, lower_in, upper_in, step_in, len, step_n, lo, hi, opts, line,
    );

    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(chunks, current, m, line);
    get(chunks, current, lo, line);
    set(chunks, current, i, line);

    let st = loops::emit_loop_start(chunks, current, line);
    emit_stride_cond(chunks, current, step_n, i, hi, line);
    loops::emit_loop_cond(chunks, current, line);
    // obj[i] = value[m]
    get(chunks, current, obj, line);
    get(chunks, current, i, line);
    get(chunks, current, value, line);
    get(chunks, current, m, line);
    collections::emit_get(chunks, current, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // m += 1; i += step
    get(chunks, current, m, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(chunks, current, m, line);
    get(chunks, current, i, line);
    get(chunks, current, step_n, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(chunks, current, i, line);
    loops::emit_loop_end(chunks, current, st, line);
}

/// Strided deletion `del obj[lo:hi:step]`. Stack:
/// `[obj, lower|null, upper|null, step|null] -> []`. Removes each strided index
/// in place; a running counter keeps positions valid without collecting indices.
pub fn emit_strided_del(chunks: &mut [Chunk], current: usize, line: u32, opts: Options) {
    let b = chunks[current].alloc_scratch(11);
    let obj = b;
    let lower_in = b + 1;
    let upper_in = b + 2;
    let step_in = b + 3;
    let len = b + 4;
    let step_n = b + 5;
    let lo = b + 6;
    let hi = b + 7;
    let i = b + 8;
    let removed = b + 9;
    let target = b + 10;

    set(chunks, current, step_in, line);
    set(chunks, current, upper_in, line);
    set(chunks, current, lower_in, line);
    set(chunks, current, obj, line);

    get(chunks, current, obj, line);
    collections::emit_len(chunks, current, line);
    set(chunks, current, len, line);

    normalize_slice_into(
        chunks, current, lower_in, upper_in, step_in, len, step_n, lo, hi, opts, line,
    );

    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(chunks, current, removed, line);
    get(chunks, current, lo, line);
    set(chunks, current, i, line);

    let st = loops::emit_loop_start(chunks, current, line);
    emit_stride_cond(chunks, current, step_n, i, hi, line);
    loops::emit_loop_cond(chunks, current, line);
    // target = (step > 0) ? i - removed : i  (ascending removals shift positions)
    get(chunks, current, step_n, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_gt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(chunks, current, i, line);
    get(chunks, current, removed, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    set(chunks, current, target, line);
    chunks[current].emit_else(line);
    get(chunks, current, i, line);
    set(chunks, current, target, line);
    chunks[current].emit_end(line);
    // remove_range(obj, target, 1)
    get(chunks, current, obj, line);
    get(chunks, current, target, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_remove_range(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // removed += 1; i += step
    get(chunks, current, removed, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(chunks, current, removed, line);
    get(chunks, current, i, line);
    get(chunks, current, step_n, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(chunks, current, i, line);
    loops::emit_loop_end(chunks, current, st, line);
}
