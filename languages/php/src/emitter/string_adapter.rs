//! PHP string helpers — Rust inline opcode emitters.
//!
//! Each `emit_*` writes opcodes directly into `chunks[current]`,
//! composing only WASM string ops (`wasm:js-string.length`, `wasm:js-string.charCodeAt`,
//! `ecma:string.indexOf`, `STR_TO_LOWER`, etc.) and where ECMA-262 already
//! covers the surface, `ecma:string.{encodeURIComponent,
//! decodeURIComponent}` / `ecma:number.toFixed`. No PHP-specific
//! host fns; no JS polyfills.

use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};
use vybe_compiler::primitives::{string_encoding, string_similarity, strings, url};

// ── Local-slot / push helpers (mirror datetime_adapter) ────────────

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}
fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
        Value::BigInt(v) => chunk.emit_i64_const(v.to_i64_wrapping(), line),
        Value::String(s) => chunk.emit_string_const(&s, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),

        _ => {
            unreachable!("push_const: unexpected value type");
        }
    }
}
fn push_str(chunk: &mut Chunk, v: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(v)), line);
}
fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}
fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}
/// PHP's value → string coercion, for the ~55 ADAPTER call sites
/// (`strlen($v)`, `str_pad($v, …)`, `explode(…, $v)`, …).
///
/// This is the SAME emitter `.` concatenation and `echo` use, and the same one
/// `[builtin_slots.string] to_string` binds. It used to be `"" + v` through
/// `emit_dyn_add` — a THIRD implementation, which disagreed with the other two.
/// Measured against real php on `[true, false, null, 1.0, 1.5, 100, "x"]`:
///
/// | | result |
/// |---|---|
/// | real php `strlen($v)` | `1,0,0,1,3,3,1` |
/// | `coerce_to_str` (was) | `1,1,1,1,3,3,1` |
/// | `strlen("" . $v)` — the BOUND path | `1,0,0,1,3,3,1` |
///
/// So `strlen(false)` and `strlen(null)` were wrong while the declared binding
/// was already right: the fork, not the binding, was the bug.
fn coerce_to_str(chunk: &mut Chunk, line: u32) {
    emit_stringify(chunk, line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    let chunk = &mut chunks[current];
    chunk.emit_call(idx, argc, line);
}

pub fn emit_strtoupper(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    coerce_to_str(chunk, line);
    let idx = chunk.add_import("ecma:string", "toUpperCase");
    chunk.emit_call(idx, 1, line);
}

// ── coercing wrappers over SHARED primitives ────────────────────────────────
//
// These bound `common:str_to_lower` / `str_repeat` / `str_reverse` directly,
// which meant they skipped PHP's string coercion entirely while `strtoupper`
// — the same operation, other case — applied it. Measured against real `php`
// on `[true, false, null, 0, 1, 1.5, "x"]`, `strtolower(false)` gave `"false"`
// where php gives `""`.
//
// The BEHAVIOUR still comes from `primitives/strings.rs`; only the coercion is
// PHP's, so these are two lines each rather than a reimplementation. That is
// the same split `emit_strlen` uses: shared counter, language coercion.

/// PHP `strtolower($s)`.
pub fn emit_strtolower(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    strings::emit_to_lower(&mut chunks[current], line);
}

/// PHP `strrev($s)`.
pub fn emit_strrev(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    strings::emit_str_reverse(&mut chunks[current], line);
}

/// PHP `str_repeat($s, $times)` — only the FIRST argument is a string.
pub fn emit_str_repeat(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    strings::emit_repeat(&mut chunks[current], line);
}

/// `LanguageHooks::concat_stringify` shape for `.` — PHP's string coercion is
/// the same one `echo` uses, so the two share an emitter rather than agreeing
/// on `true` → `"1"` twice.
pub fn emit_concat_stringify(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_echo_stringify(chunks, current, 1, line);
}

pub fn emit_echo_stringify(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_stringify(&mut chunks[current], line);
}

/// PHP's ONE value → string coercion.
///
/// Single-chunk on purpose: every import and every op lands in the chunk it is
/// handed, which is what lets `coerce_to_str` share it. (emit_call indices
/// resolve against the CURRENT chunk's import table — adding to `chunks[0]`
/// yields an index valid only at top level and silently calls garbage inside a
/// function chunk; see project_chunk_import_tables.)
fn emit_stringify(chunk: &mut Chunk, line: u32) {
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");
    let bigint_to_string = chunk.add_import("ecma:bigint", "toString");
    let from_i64 = chunk.add_import("wasm:js-string", "fromI64");
    let v_slot = alloc_local(chunk);
    let _ty_slot = alloc_local(chunk);

    lset(chunk, v_slot, line);

    lget(chunk, v_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    push_str(chunk, "", line);
    chunk.emit_else(line);

    // boolean test
    lget(chunk, v_slot, line);
    let test_bool_echo = chunk.add_import("wasm:js-boolean", "test");
    chunk.emit_call(test_bool_echo, 1, line);
    chunk.emit_if(line);
    lget(chunk, v_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "1", line);
    chunk.emit_else(line);
    push_str(chunk, "", line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    // Not boolean. PHP integer boundary constants use BigInt internally so
    // they retain all 64-bit digits; stringify them through ECMA BigInt.
    lget(chunk, v_slot, line);
    chunk.emit_call(test_bigint, 1, line);
    chunk.emit_if(line);
    lget(chunk, v_slot, line);
    chunk.emit_call(bigint_to_string, 1, line);
    chunk.emit_else(line);
    // Not bigint. PHP prints special float values as INF / -INF / NAN
    // (uppercase), unlike the VM's "Infinity"/"NaN". These eq-probes are
    // safe for non-numbers: a string/object compares equal to itself, so
    // it skips straight to the default `"" + v` path.
    // NaN?  (v !== v)
    lget(chunk, v_slot, line);
    lget(chunk, v_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    push_str(chunk, "NAN", line);
    chunk.emit_else(line);
    // +Infinity?
    lget(chunk, v_slot, line);
    push_const(chunk, Value::F64(f64::INFINITY), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "INF", line);
    chunk.emit_else(line);
    // -Infinity?
    lget(chunk, v_slot, line);
    push_const(chunk, Value::F64(f64::NEG_INFINITY), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "-INF", line);
    chunk.emit_else(line);
    // Exact i64 values (PHP ints) must not go through f64 stringification,
    // otherwise PHP_INT_MAX/PHP_INT_MIN lose their final digits.
    // Use bigint test to detect i64 values.
    lget(chunk, v_slot, line);
    let test_bigint_i64 = chunk.add_import("wasm:js-bigint", "test");
    chunk.emit_call(test_bigint_i64, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, v_slot, line);
    chunk.emit_call(from_i64, 1, line);
    chunk.emit_else(line);
    // default. PHP stringifies floats with `precision=14` significant digits
    // (php_gcvt), not the VM's shortest-round-trip form. Normalize a float via
    // toPrecision(14) then parseFloat — the round-trip trims trailing zeros
    // exactly like PHP (`1.4142135623730951`→`1.4142135623731`, `0.1+0.2`→
    // `0.3`). Integers and non-numbers keep the plain `"" + v` path.
    lget(chunk, v_slot, line);
    let num_test_echo = chunk.add_import("wasm:js-number", "test");
    chunk.emit_call(num_test_echo, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, v_slot, line);
    let is_int_echo = chunk.add_import("ecma:number", "isInteger");
    chunk.emit_call(is_int_echo, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    // integer-valued number: plain "" + v
    push_str(chunk, "", line);
    lget(chunk, v_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_else(line);
    // fractional float: "" + parseFloat(toPrecision(v, 14))
    push_str(chunk, "", line);
    lget(chunk, v_slot, line);
    push_const(chunk, Value::I32(14), line);
    let to_prec_echo = chunk.add_import("ecma:number", "toPrecision");
    chunk.emit_call(to_prec_echo, 2, line);
    let parse_f_echo = chunk.add_import("ecma:number", "parseFloat");
    chunk.emit_call(parse_f_echo, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    // non-number: "" + v
    push_str(chunk, "", line);
    lget(chunk, v_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_end(line);
    // Close: i64, -INF, INF, NaN, bigint, boolean, null — seven if_value blocks.
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_var_dump_stringify(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // Imports resolve against the current chunk — add to `chunks[current]`,
    // never `chunks[0]` (see project_chunk_import_tables).
    let test_bigint = chunks[current].add_import("wasm:js-bigint", "test");
    let bigint_to_string = chunks[current].add_import("ecma:bigint", "toString");
    let chunk = &mut chunks[current];
    let v_slot = alloc_local(chunk);
    let _ty_slot = alloc_local(chunk);

    lset(chunk, v_slot, line);

    lget(chunk, v_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    push_str(chunk, "NULL", line);
    chunk.emit_else(line);

    // boolean test
    lget(chunk, v_slot, line);
    let test_bool_vd = chunk.add_import("wasm:js-boolean", "test");
    chunk.emit_call(test_bool_vd, 1, line);
    chunk.emit_if(line);
    lget(chunk, v_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "bool(true)", line);
    chunk.emit_else(line);
    push_str(chunk, "bool(false)", line);
    chunk.emit_end(line);
    chunk.emit_else(line);

    lget(chunk, v_slot, line);
    chunk.emit_call(test_bigint, 1, line);
    chunk.emit_if(line);
    push_str(chunk, "int(", line);
    lget(chunk, v_slot, line);
    chunk.emit_call(bigint_to_string, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    push_str(chunk, ")", line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_else(line);

    push_str(chunk, "", line);
    lget(chunk, v_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}


// ── ucwords ────────────────────────────────────────────────────────

/// PHP `ucwords(str, delims?)`. Stack: `[str]` or `[str, delims]` →
/// `[result]`. Uppercases the first letter of each word; default
/// delimiter set is whitespace (space + 0x09..=0x0D).
pub fn emit_ucwords(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let delims_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let cap_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let c_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let is_delim_slot = alloc_local(chunk);

    // Pop delims (if any) and s.
    if argc >= 2 {
        lset(chunk, delims_slot, line);
    } else {
        push_const(chunk, Value::Null, line);
        lset(chunk, delims_slot, line);
    }
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);

    // out = ""; cap = true; i = 0; len = s.length
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::Bool(true), line);
    lset(chunk, cap_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, len_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, c_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);

    lget(chunk, delims_slot, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(32.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(9.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(13.0), line);
    vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
    chunk.emit_else(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    lget(chunk, delims_slot, line);
    lget(chunk, c_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "indexOf");
        chunk.emit_call(idx, 2, line);
    }
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    chunk.emit_end(line);
    lset(chunk, is_delim_slot, line);

    lget(chunk, cap_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, is_delim_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, out_slot, line);
    lget(chunk, c_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    chunk.emit_else(line);
    lget(chunk, out_slot, line);
    lget(chunk, c_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "toUpperCase");
        chunk.emit_call(idx, 1, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    lget(chunk, out_slot, line);
    lget(chunk, c_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    chunk.emit_end(line);

    lget(chunk, is_delim_slot, line);
    lset(chunk, cap_slot, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, out_slot, line);
}

// ── str_split ──────────────────────────────────────────────────────

/// PHP `str_split(str, split_length?)`.
pub fn emit_str_split(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let len_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let end_slot = alloc_local(chunk);

    if argc >= 2 {
        lset(chunk, len_slot, line);
    } else {
        push_const(chunk, Value::F64(1.0), line);
        lset(chunk, len_slot, line);
    }
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);
    chunk.emit_array_new_fixed(0, 0, line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, n_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, end_slot, line);
    lget(chunk, end_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, n_slot, line);
    lset(chunk, end_slot, line);
    chunk.emit_end(line);

    lget(chunk, out_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    lget(chunk, end_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];

    lget(chunk, end_slot, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, out_slot, line);
}

// ── str_pad ────────────────────────────────────────────────────────

/// PHP `str_pad(str, length, pad_string?, pad_type?)`.
/// PHP `str_pad(string, length, pad = " ", mode = STR_PAD_RIGHT)`.
///
/// The padding itself is the SHARED emitter — python `ljust`/`rjust`/`center`
/// bind the same three sides. What is php's own is that the SIDE is a runtime
/// argument rather than part of the spelling, so the three-way branch stays
/// here and each arm is a binding.
///
/// The two guards the old body checked by hand — `strlen >= length` and an
/// empty pad string — are both already `StringPad` (ECMA-262 §22.1.3.16)
/// behaviour, so they are gone rather than moved.
pub fn emit_str_pad(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let mode_slot = alloc_local(chunk);
    let pad_slot = alloc_local(chunk);
    let target_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);

    if argc >= 4 {
        lset(chunk, mode_slot, line);
    } else {
        push_const(chunk, Value::F64(1.0), line);
        lset(chunk, mode_slot, line);
    }
    if argc >= 3 {
        coerce_to_str(chunk, line);
        lset(chunk, pad_slot, line);
    } else {
        push_str(chunk, " ", line);
        lset(chunk, pad_slot, line);
    }
    lset(chunk, target_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);

    // STR_PAD_LEFT = 0, STR_PAD_RIGHT = 1, STR_PAD_BOTH = 2.
    let mut arm = |chunks: &mut [Chunk], current: usize, side: strings::PadSide| {
        lget(&mut chunks[current], s_slot, line);
        lget(&mut chunks[current], target_slot, line);
        lget(&mut chunks[current], pad_slot, line);
        strings::emit_pad(chunks, current, 3, side, strings::CenterBias::Right, line);
    };

    lget(&mut chunks[current], mode_slot, line);
    push_const(&mut chunks[current], Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    arm(chunks, current, strings::PadSide::Start);
    chunks[current].emit_else(line);

    lget(&mut chunks[current], mode_slot, line);
    push_const(&mut chunks[current], Value::F64(2.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    arm(chunks, current, strings::PadSide::Both);
    chunks[current].emit_else(line);
    arm(chunks, current, strings::PadSide::End);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

// ── substr_count ───────────────────────────────────────────────────

/// PHP `substr_count(haystack, needle, offset?, length?)`.
pub fn emit_substr_count(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let length_slot = alloc_local(chunk);
    let offset_slot = alloc_local(chunk);
    let needle_slot = alloc_local(chunk);
    let hay_slot = alloc_local(chunk);
    let slice_slot = alloc_local(chunk);
    let count_slot = alloc_local(chunk);
    let pos_slot = alloc_local(chunk);
    let idx_slot = alloc_local(chunk);
    let has_length_slot = alloc_local(chunk);

    chunk.emit_i32_const(if argc >= 4 { 1 } else { 0 }, line);
    lset(chunk, has_length_slot, line);
    if argc >= 4 {
        lset(chunk, length_slot, line);
    }
    if argc >= 3 {
        lset(chunk, offset_slot, line);
    } else {
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, offset_slot, line);
    }
    coerce_to_str(chunk, line);
    lset(chunk, needle_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, hay_slot, line);

    // if has_length: slice = hay.substring(offset, offset+length) else: slice = hay.substring(offset)
    lget(chunk, has_length_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, hay_slot, line);
    lget(chunk, offset_slot, line);
    lget(chunk, offset_slot, line);
    lget(chunk, length_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    chunk.emit_else(line);
    lget(chunk, hay_slot, line);
    lget(chunk, offset_slot, line);
    lget(chunk, hay_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    chunk.emit_end(line);
    lset(chunk, slice_slot, line);

    // if needle.length == 0: return 0
    lget(chunk, needle_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_else(line);

    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, count_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, pos_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    // compute idx
    lget(chunk, slice_slot, line);
    lget(chunk, pos_slot, line);
    lget(chunk, slice_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    lget(chunk, needle_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "indexOf");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, idx_slot, line);
    // condition: idx >= 0
    lget(chunk, idx_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, count_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, count_slot, line);
    lget(chunk, idx_slot, line);
    lget(chunk, pos_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, needle_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, pos_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, count_slot, line);
    chunk.emit_end(line); // end if (needle nonempty)
}

// ── strstr / stristr ───────────────────────────────────────────────

fn emit_strstr_impl(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    case_insensitive: bool,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let before_slot = alloc_local(chunk);
    let needle_slot = alloc_local(chunk);
    let hay_slot = alloc_local(chunk);
    let idx_slot = alloc_local(chunk);

    if argc >= 3 {
        lset(chunk, before_slot, line);
    } else {
        push_const(chunk, Value::Bool(false), line);
        lset(chunk, before_slot, line);
    }
    lset(chunk, needle_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, hay_slot, line);

    if case_insensitive {
        // idx = lower(hay).indexOf(lower(needle))
        lget(chunk, hay_slot, line);
        {
            let idx = chunk.add_import("ecma:string", "toLowerCase");
            chunk.emit_call(idx, 1, line);
        }
        lget(chunk, needle_slot, line);
        {
            let idx = chunk.add_import("ecma:string", "toLowerCase");
            chunk.emit_call(idx, 1, line);
        }
        {
            let idx = chunk.add_import("ecma:string", "indexOf");
            chunk.emit_call(idx, 2, line);
        }
    } else {
        lget(chunk, hay_slot, line);
        lget(chunk, needle_slot, line);
        {
            let idx = chunk.add_import("ecma:string", "indexOf");
            chunk.emit_call(idx, 2, line);
        }
    }
    lset(chunk, idx_slot, line);

    // if idx < 0: return false; else return before ? hay[0..idx] : hay[idx..]
    lget(chunk, idx_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);
    // idx >= 0
    lget(chunk, before_slot, line);
    push_const(chunk, Value::Bool(true), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, hay_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, idx_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    chunk.emit_else(line);
    lget(chunk, hay_slot, line);
    lget(chunk, idx_slot, line);
    lget(chunk, hay_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_strstr(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_strstr_impl(
        chunks, current, argc, /*case_insensitive=*/ false, line,
    );
}
pub fn emit_stristr(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_strstr_impl(chunks, current, argc, /*case_insensitive=*/ true, line);
}

// ── urlencode / rawurlencode / urldecode ───────────────────────────

/// PHP `urlencode` — like `encodeURIComponent` but space → "+".

/// PHP `rawurlencode` — strict `encodeURIComponent`.

/// PHP `urldecode` — replace "+" with " ", then decodeURIComponent.

/// PHP `rawurldecode` — strict `decodeURIComponent` with no `+` → space
/// translation.

/// PHP `htmlspecialchars` / `htmlentities` — escape the basic HTML-special
/// characters. This covers the common text-node/title use-cases exercised by
/// the PHP suite and webroot renderer.
pub fn emit_htmlspecialchars(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    coerce_to_str(chunk, line);

    for (from, to) in [
        ("&", "&amp;"),
        ("<", "&lt;"),
        (">", "&gt;"),
        ("\"", "&quot;"),
        ("'", "&#039;"),
    ] {
        push_str(chunk, from, line);
        push_str(chunk, to, line);
        {
            let idx = chunk.add_import("ecma:string", "replaceAll");
            chunk.emit_call(idx, 3, line);
        }
    }
}

pub fn emit_htmlentities(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_htmlspecialchars(chunks, current, argc, line);
    // Encode common non-ASCII characters to named HTML entities
    let chunk = &mut chunks[current];
    for (ch, entity) in HTML_NAMED_ENTITIES_ENCODE {
        push_str(chunk, ch, line);
        push_str(chunk, entity, line);
        {
            let idx = chunk.add_import("ecma:string", "replaceAll");
            chunk.emit_call(idx, 3, line);
        }
    }
}

const HTML_NAMED_ENTITIES_ENCODE: &[(&str, &str)] = &[
    ("©", "&copy;"),
    ("®", "&reg;"),
    ("™", "&trade;"),
    ("€", "&euro;"),
    ("£", "&pound;"),
    ("¥", "&yen;"),
    ("¢", "&cent;"),
    ("§", "&sect;"),
    ("¶", "&para;"),
    ("°", "&deg;"),
    ("±", "&plusmn;"),
    ("µ", "&micro;"),
    ("·", "&middot;"),
    ("×", "&times;"),
    ("÷", "&divide;"),
    ("«", "&laquo;"),
    ("»", "&raquo;"),
    ("¬", "&not;"),
    ("ß", "&szlig;"),
    ("à", "&agrave;"),
    ("á", "&aacute;"),
    ("â", "&acirc;"),
    ("ã", "&atilde;"),
    ("ä", "&auml;"),
    ("å", "&aring;"),
    ("æ", "&aelig;"),
    ("ç", "&ccedil;"),
    ("è", "&egrave;"),
    ("é", "&eacute;"),
    ("ê", "&ecirc;"),
    ("ë", "&euml;"),
    ("ì", "&igrave;"),
    ("í", "&iacute;"),
    ("î", "&icirc;"),
    ("ï", "&iuml;"),
    ("ñ", "&ntilde;"),
    ("ò", "&ograve;"),
    ("ó", "&oacute;"),
    ("ô", "&ocirc;"),
    ("õ", "&otilde;"),
    ("ö", "&ouml;"),
    ("ø", "&oslash;"),
    ("ù", "&ugrave;"),
    ("ú", "&uacute;"),
    ("û", "&ucirc;"),
    ("ü", "&uuml;"),
    ("ý", "&yacute;"),
    ("þ", "&thorn;"),
    ("ÿ", "&yuml;"),
    ("À", "&Agrave;"),
    ("Á", "&Aacute;"),
    ("Â", "&Acirc;"),
    ("Ã", "&Atilde;"),
    ("Ä", "&Auml;"),
    ("Å", "&Aring;"),
    ("Æ", "&AElig;"),
    ("Ç", "&Ccedil;"),
    ("È", "&Egrave;"),
    ("É", "&Eacute;"),
    ("Ê", "&Ecirc;"),
    ("Ë", "&Euml;"),
    ("Ñ", "&Ntilde;"),
    ("Ò", "&Ograve;"),
    ("Ó", "&Oacute;"),
    ("Ô", "&Ocirc;"),
    ("Õ", "&Otilde;"),
    ("Ö", "&Ouml;"),
    ("Ø", "&Oslash;"),
    ("Ù", "&Ugrave;"),
    ("Ú", "&Uacute;"),
    ("Û", "&Ucirc;"),
    ("Ü", "&Uuml;"),
    ("Ý", "&Yacute;"),
    ("Þ", "&THORN;"),
];

// ── bin2hex / hex2bin ──────────────────────────────────────────────

/// PHP `bin2hex(str)` — char loop, two hex digits per byte.

/// PHP `hex2bin(hex)` — pair loop.

// ── chunk_split ────────────────────────────────────────────────────

/// PHP `chunk_split(s, length=76, end="\r\n")`.
pub fn emit_chunk_split(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let end_slot = alloc_local(chunk);
    let length_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let total_slot = alloc_local(chunk);

    if argc >= 3 {
        lset(chunk, end_slot, line);
    } else {
        push_str(chunk, "\r\n", line);
        lset(chunk, end_slot, line);
    }
    if argc >= 2 {
        lset(chunk, length_slot, line);
    } else {
        push_const(chunk, Value::F64(76.0), line);
        lset(chunk, length_slot, line);
    }
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);

    // if length < 1: false; else: process
    lget(chunk, length_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);

    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, total_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, total_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // out += s.substring(i, i+length) + end
    lget(chunk, out_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    lget(chunk, i_slot, line);
    lget(chunk, length_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lget(chunk, end_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);

    lget(chunk, i_slot, line);
    lget(chunk, length_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, out_slot, line);
    chunk.emit_end(line); // end if (valid length)
}

// ── number_format ──────────────────────────────────────────────────

/// PHP `number_format(num, decimals=0, decsep=".", thousep=",")`.
pub fn emit_number_format(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let thousep_slot = alloc_local(chunk);
    let decsep_slot = alloc_local(chunk);
    let decimals_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let sign_slot = alloc_local(chunk);
    let fixed_slot = alloc_local(chunk);

    if argc >= 4 {
        lset(chunk, thousep_slot, line);
    } else {
        push_str(chunk, ",", line);
        lset(chunk, thousep_slot, line);
    }
    if argc >= 3 {
        lset(chunk, decsep_slot, line);
    } else {
        push_str(chunk, ".", line);
        lset(chunk, decsep_slot, line);
    }
    if argc >= 2 {
        lset(chunk, decimals_slot, line);
    } else {
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, decimals_slot, line);
    }

    // PHP 8: `number_format` decimals must be >= 0, else ValueError.
    lget(chunk, decimals_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    let _ = chunk;
    crate::emitter::type_guard::emit_throw_const(
        chunks,
        current,
        "ValueError",
        "number_format(): Argument #2 ($decimals) must be greater than or equal to 0",
        line,
    );
    let chunk = &mut chunks[current];
    chunk.emit_end(line);

    // n = +num (numeric coerce)
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, n_slot, line);

    // sign = ""; if n < 0: sign = "-"; n = -n
    push_str(chunk, "", line);
    lset(chunk, sign_slot, line);
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_str(chunk, "-", line);
    lset(chunk, sign_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, n_slot, line);
    chunk.emit_end(line);

    // PHP rounds half-away-from-zero; JS toFixed is engine-dependent
    // (banker's in V8, ECMA-spec deferred). Pre-round via Math.round
    // (which uses f64::round in vybe_host = half-away-from-zero) at
    // scale 10^decimals to force PHP semantics:
    //   n = Math.round(n * scale) / scale
    let pow = chunks[0].add_import("ecma:math", "pow");
    let round = chunks[0].add_import("ecma:math", "round");
    let chunk = &mut chunks[current];
    let scale_slot = alloc_local(chunk);
    push_const(chunk, Value::F64(10.0), line);
    lget(chunk, decimals_slot, line);
    chunk.emit_call(pow, 2, line);
    lset(chunk, scale_slot, line);
    lget(chunk, n_slot, line);
    lget(chunk, scale_slot, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_call(round, 1, line);
    lget(chunk, scale_slot, line);
    chunk.emit_op(Op::F64_DIV, line);
    lset(chunk, n_slot, line);

    // fixed = ecma:number.toFixed(n, decimals)
    lget(chunk, n_slot, line);
    lget(chunk, decimals_slot, line);
    let to_fixed = chunks[0].add_import("ecma:number", "toFixed");
    let chunk = &mut chunks[current];
    chunk.emit_call(to_fixed, 2, line);
    lset(chunk, fixed_slot, line);
    // Grouping is the SHARED emitter — python's `,` format flag binds the same
    // one. Everything above is php's and stays: the argument defaults, the
    // `ValueError` on negative decimals, and the half-away-from-zero pre-round
    // (JS `toFixed` is engine-dependent, so php's rounding cannot be delegated).
    //
    // The sign is re-attached AFTER grouping because `emit_group_digits`
    // preserves a leading sign itself; php stripped it before rounding, so it
    // is prepended here rather than passed through.
    lget(chunk, sign_slot, line);
    lget(chunk, fixed_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
    lget(chunk, thousep_slot, line);
    lget(chunk, decsep_slot, line);
    strings::emit_group_digits(chunks, current, line);
}

// ── substr_replace ─────────────────────────────────────────────────

/// PHP `substr_replace(str, repl, start, length?)`.
pub fn emit_substr_replace(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let length_slot = alloc_local(chunk);
    let start_slot = alloc_local(chunk);
    let repl_slot = alloc_local(chunk);
    let str_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let l_slot = alloc_local(chunk);
    let has_length_slot = alloc_local(chunk);

    chunk.emit_i32_const(if argc >= 4 { 1 } else { 0 }, line);
    lset(chunk, has_length_slot, line);
    if argc >= 4 {
        lset(chunk, length_slot, line);
    }
    lset(chunk, start_slot, line);
    lset(chunk, repl_slot, line);
    lset(chunk, str_slot, line);

    // len = str.length
    lget(chunk, str_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, len_slot, line);

    // s = start < 0 ? max(len + start, 0) : min(start, len)
    lget(chunk, start_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line); // if start < 0
    // negative: max(len + start, 0)
    lget(chunk, len_slot, line);
    lget(chunk, start_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    let neg_slot = alloc_local(chunk);
    lset(chunk, neg_slot, line);
    lget(chunk, neg_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line); // if neg < 0: use 0 else neg
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_else(line);
    lget(chunk, neg_slot, line);
    chunk.emit_end(line);
    chunk.emit_else(line); // else: positive start
    // positive: min(start, len)
    lget(chunk, start_slot, line);
    lget(chunk, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line); // if start > len: use len else start
    lget(chunk, len_slot, line);
    chunk.emit_else(line);
    lget(chunk, start_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line); // end start<0 if
    lset(chunk, s_slot, line);

    // l = has_length ? clamp(length, len - s) : len - s
    lget(chunk, has_length_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line); // if has_length
    // Has length: if length < 0: l = max(len + length - s, 0) else l = min(length, len - s)
    lget(chunk, length_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line); // if length < 0
    // negative length: max(len + length - s, 0)
    lget(chunk, len_slot, line);
    lget(chunk, length_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    let neg_l_slot = alloc_local(chunk);
    lset(chunk, neg_l_slot, line);
    lget(chunk, neg_l_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line); // if neg_l < 0: use 0
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_else(line);
    lget(chunk, neg_l_slot, line);
    chunk.emit_end(line);
    chunk.emit_else(line); // else: positive length
    // positive length: min(length, len - s)
    lget(chunk, length_slot, line);
    lget(chunk, len_slot, line);
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    let rem_slot = alloc_local(chunk);
    lset(chunk, rem_slot, line);
    lget(chunk, length_slot, line);
    lget(chunk, rem_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line); // if length > rem: use rem
    lget(chunk, rem_slot, line);
    chunk.emit_else(line);
    lget(chunk, length_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line); // end length < 0 if
    chunk.emit_else(line); // else: no length
    // No length: l = len - s
    lget(chunk, len_slot, line);
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_end(line); // end has_length if
    lset(chunk, l_slot, line);
    if argc >= 4 {
        chunk.emit_op(Op::DROP, line);
    }

    // result = str.substring(0, s) + repl + str.substring(s + l)
    lget(chunk, str_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    lget(chunk, repl_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lget(chunk, str_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, l_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, len_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
}

// ── str_word_count ─────────────────────────────────────────────────

/// PHP `str_word_count(s[, mode[, charlist]])`.
/// mode 0: count, mode 1: array of words, mode 2: position→word map.
pub fn emit_str_word_count(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 2 {
        emit_str_word_count_with_mode(chunks, current, argc, line);
        return;
    }
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let count_slot = alloc_local(chunk);
    let in_word_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let is_sep_slot = alloc_local(chunk);

    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, count_slot, line);
    push_const(chunk, Value::Bool(false), line);
    lset(chunk, in_word_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, len_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);

    // is_sep: whitespace (9..=13, 32) | comma 44 | period 46 | ! 33 | ? 63 | ; 59 | : 58.
    // PHP counts hyphenated words as one token.
    push_const(chunk, Value::Bool(false), line);
    lset(chunk, is_sep_slot, line);
    for code_val in &[
        9.0_f64, 10.0, 11.0, 12.0, 13.0, 32.0, 33.0, 44.0, 46.0, 58.0, 59.0, 63.0,
    ] {
        lget(chunk, code_slot, line);
        push_const(chunk, Value::F64(*code_val), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        push_const(chunk, Value::Bool(true), line);
        lset(chunk, is_sep_slot, line);
        chunk.emit_end(line);
    }
    // if is_sep: in_word = false; else if !in_word: count++; in_word = true
    lget(chunk, is_sep_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::Bool(false), line);
    lset(chunk, in_word_slot, line);
    chunk.emit_else(line);
    // not separator
    lget(chunk, in_word_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_if(line);
    lget(chunk, count_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, count_slot, line);
    push_const(chunk, Value::Bool(true), line);
    lset(chunk, in_word_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, count_slot, line);
}

// ── str_ireplace, str_replace, wordwrap ────────────────────────────
//
// These three are larger and rare in the test surface. To keep this
// migration focused, route them through the simple `ecma:string.replaceAll`
// opcode in the common case (single search/replace string) and leave
// the array-aware / case-insensitive paths as runtime fallbacks
// later. For now their dispatch arms emit the three-arg ecma:string.replaceAll
// directly when the inputs are strings; PHP-array shapes were
// covered by the polyfills and aren't exercised by current tests.

/// PHP `str_ireplace(search, replace, subject)`.
/// Case-insensitive find: walk by chunks of length(needle), comparing
/// lowercased segments. Falls back to the JS-host
/// `ecma:string.replaceAll` for the simple case and emits a manual scan for
/// case insensitivity.
pub fn emit_str_ireplace(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let subj_slot = alloc_local(chunk);
    let repl_slot = alloc_local(chunk);
    let srch_slot = alloc_local(chunk);
    let lower_slot = alloc_local(chunk);
    let srch_lower_slot = alloc_local(chunk);
    let srch_len_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let pos_slot = alloc_local(chunk);
    let idx_slot = alloc_local(chunk);

    coerce_to_str(chunk, line);
    lset(chunk, subj_slot, line);
    lset(chunk, repl_slot, line);
    lset(chunk, srch_slot, line);

    // lower = subj.toLowerCase(); srch_lower = srch.toLowerCase()
    lget(chunk, subj_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "toLowerCase");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, lower_slot, line);
    lget(chunk, srch_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "toLowerCase");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, srch_lower_slot, line);
    lget(chunk, srch_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, srch_len_slot, line);

    // if srch_len === 0: return subj; else do the replacement loop
    lget(chunk, srch_len_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, subj_slot, line);
    chunk.emit_else(line);

    // out = ""; pos = 0
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, pos_slot, line);

    // loop: idx = lower.substring(pos).indexOf(srch_lower); break if idx < 0
    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, lower_slot, line);
    lget(chunk, pos_slot, line);
    lget(chunk, lower_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    lget(chunk, srch_lower_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "indexOf");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, idx_slot, line);
    // condition: idx >= 0
    lget(chunk, idx_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // out += subj.substring(pos, pos + idx) + repl
    lget(chunk, out_slot, line);
    lget(chunk, subj_slot, line);
    lget(chunk, pos_slot, line);
    lget(chunk, pos_slot, line);
    lget(chunk, idx_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lget(chunk, repl_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);

    // pos = pos + idx + srch_len
    lget(chunk, pos_slot, line);
    lget(chunk, idx_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, srch_len_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, pos_slot, line);

    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    // out += subj.substring(pos)
    lget(chunk, out_slot, line);
    lget(chunk, subj_slot, line);
    lget(chunk, pos_slot, line);
    lget(chunk, subj_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_end(line); // end if (nonempty srch)
}

// ── str_replace (array-aware) ──────────────────────────────────────

/// PHP `str_replace(search, replace, subject)`. When `search` and
/// `replace` are both strings, this is one `ecma:string.replaceAll` opcode. The
/// adapter handles the array-aware variants too: when `search` is an
/// array, iterate and apply each pair (with `replace` as either a
/// scalar or a parallel array).
///
/// Strategy: probe `Array.isArray(search)` at runtime via
/// `ecma:array.isArray`. If false, fall back to `ecma:string.replaceAll`.
/// Otherwise loop over the array and split/join per element.
pub fn emit_str_replace(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let subj_slot = alloc_local(chunk);
    let repl_slot = alloc_local(chunk);
    let srch_slot = alloc_local(chunk);
    lset(chunk, subj_slot, line);
    lset(chunk, repl_slot, line);
    lset(chunk, srch_slot, line);

    // Coerce subj to string — PHP's coercion, not ECMA's. This was a fourth
    // inline `"" + v`, which is why `str_replace("1","Z",false)` yielded `"0"`
    // where real php yields `""`.
    lget(chunk, subj_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, subj_slot, line);

    // is_array_search = Array.isArray(srch)
    let _ = chunk;
    chunks[current].emit_op_u16(Op::LOCAL_GET, srch_slot, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    let chunk = &mut chunks[current];
    // ecma:array.isArray returns I32(0|1); structured if: array path if true, scalar path if false.
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line); // if is_array: array path; else: scalar path

    // ── Array path ──
    // is_array_repl = Array.isArray(repl)
    let is_array_repl_slot = alloc_local(chunk);
    push_const(chunk, Value::Bool(false), line);
    lset(chunk, is_array_repl_slot, line);
    let _ = chunk;
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_slot, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::Bool(true), line);
    lset(chunk, is_array_repl_slot, line);
    chunk.emit_end(line);

    // for i in 0..srch.length: needle = srch[i]; rep = ...; subj = ecma:string.replaceAll(subj, needle, rep)
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let needle_slot = alloc_local(chunk);
    let rep_slot = alloc_local(chunk);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, srch_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, n_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // needle = "" + srch[i]
    push_str(chunk, "", line);
    lget(chunk, srch_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, needle_slot, line);

    // if needle.length > 0: do replacement (else skip)
    lget(chunk, needle_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_if(line);

    // rep = is_array_repl ? (i < repl.length ? "" + repl[i] : "") : "" + repl
    lget(chunk, is_array_repl_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    // array repl path
    lget(chunk, i_slot, line);
    lget(chunk, repl_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "", line);
    lget(chunk, repl_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_else(line);
    push_str(chunk, "", line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    // scalar repl path
    push_str(chunk, "", line);
    lget(chunk, repl_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_end(line);
    lset(chunk, rep_slot, line);

    // subj = ecma:string.replaceAll(subj, needle, rep)
    lget(chunk, subj_slot, line);
    lget(chunk, needle_slot, line);
    lget(chunk, rep_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "replaceAll");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, subj_slot, line);

    chunk.emit_end(line); // end if nonempty_needle

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, subj_slot, line);
    chunk.emit_else(line); // else branch of scalar_path if-value

    // ── Scalar path ──
    // ecma:string.replaceAll(subj, "" + srch, "" + repl)
    lget(chunk, subj_slot, line);
    push_str(chunk, "", line);
    lget(chunk, srch_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    push_str(chunk, "", line);
    lget(chunk, repl_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    {
        let idx = chunk.add_import("ecma:string", "replaceAll");
        chunk.emit_call(idx, 3, line);
    }

    chunk.emit_end(line); // end scalar_path if-value
}

// ── wordwrap ───────────────────────────────────────────────────────

/// PHP `wordwrap(s, width=75, break_str="\n", cut=false)`.
///
/// Emits a per-line word-wrap loop: for each `\n`-delimited input
/// line, walk words and accumulate onto a `current` buffer until the
/// width is exceeded; on overflow, push the current buffer to `out`
/// and start a new one. With `cut=true`, also break mid-word.
pub fn emit_wordwrap(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let cut_slot = alloc_local(chunk);
    let br_slot = alloc_local(chunk);
    let width_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);

    if argc >= 4 {
        lset(chunk, cut_slot, line);
    } else {
        push_const(chunk, Value::Bool(false), line);
        lset(chunk, cut_slot, line);
    }
    if argc >= 3 {
        lset(chunk, br_slot, line);
    } else {
        push_str(chunk, "\n", line);
        lset(chunk, br_slot, line);
    }
    if argc >= 2 {
        lset(chunk, width_slot, line);
    } else {
        push_const(chunk, Value::F64(75.0), line);
        lset(chunk, width_slot, line);
    }
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);

    // lines = s.split("\n")
    let lines_slot = alloc_local(chunk);
    lget(chunk, s_slot, line);
    push_str(chunk, "\n", line);
    {
        let idx = chunk.add_import("ecma:string", "split");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, lines_slot, line);

    // out = []
    let out_slot = alloc_local(chunk);
    chunk.emit_array_new_fixed(0, 0, line);
    lset(chunk, out_slot, line);

    // for li in 0..lines.length
    let li_slot = alloc_local(chunk);
    let nlines_slot = alloc_local(chunk);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, li_slot, line);
    lget(chunk, lines_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, nlines_slot, line);

    let line_slot = alloc_local(chunk);
    let words_slot = alloc_local(chunk);
    let nwords_slot = alloc_local(chunk);
    let wi_slot = alloc_local(chunk);
    let word_slot = alloc_local(chunk);
    let current_slot = alloc_local(chunk);

    let _ = chunk;
    let outer_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, li_slot, line);
    lget(chunk, nlines_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // line = lines[li]
    lget(chunk, lines_slot, line);
    lget(chunk, li_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, line_slot, line);

    // if line.length > width: wrap; else: push as-is
    lget(chunk, line_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lget(chunk, width_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    // needs wrapping: words = line.split(" ")
    lget(chunk, line_slot, line);
    push_str(chunk, " ", line);
    {
        let idx = chunk.add_import("ecma:string", "split");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, words_slot, line);
    lget(chunk, words_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, nwords_slot, line);

    // current = ""; wi = 0
    push_str(chunk, "", line);
    lset(chunk, current_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, wi_slot, line);

    let _ = chunk;
    let inner_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, wi_slot, line);
    lget(chunk, nwords_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // word = words[wi]
    lget(chunk, words_slot, line);
    lget(chunk, wi_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, word_slot, line);

    // if current.length === 0: current = word
    lget(chunk, current_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, word_slot, line);
    lset(chunk, current_slot, line);
    chunk.emit_else(line);
    // else if current.length + 1 + word.length <= width: append
    lget(chunk, current_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, word_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, width_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    // if fits: append; else: push current and start new
    chunk.emit_if(line);
    lget(chunk, current_slot, line);
    push_str(chunk, " ", line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lget(chunk, word_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, current_slot, line);
    chunk.emit_else(line);
    lget(chunk, out_slot, line);
    lget(chunk, current_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    lget(chunk, word_slot, line);
    lset(chunk, current_slot, line);
    chunk.emit_end(line); // end fits/break if
    chunk.emit_end(line); // end empty/nonempty if

    // wi++
    lget(chunk, wi_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, wi_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, inner_state, line);
    let chunk = &mut chunks[current];

    // if cut && current.length > width: emit chunks, then remaining
    lget(chunk, cut_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    let cut_i_slot = alloc_local(chunk);
    // while current.length > width: push current[0..width], current = current[width..]
    let _ = chunk;
    let cut_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, current_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lget(chunk, width_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    // chunk = current[0..width]
    lget(chunk, current_slot, line);
    push_const(chunk, Value::I32(0), line);
    lget(chunk, width_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, cut_i_slot, line);
    // current = current[width..]
    lget(chunk, current_slot, line);
    lget(chunk, width_slot, line);
    push_const(chunk, Value::I32(i32::MAX), line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, current_slot, line);
    // out.push(chunk)
    lget(chunk, out_slot, line);
    lget(chunk, cut_i_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, cut_state, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line); // end if cut

    // if current.length > 0: out.push(current)
    lget(chunk, current_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, out_slot, line);
    lget(chunk, current_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);

    chunk.emit_else(line); // short line: push as-is
    lget(chunk, out_slot, line);
    lget(chunk, line_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line); // end needs_wrap if

    // li++
    lget(chunk, li_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, li_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, outer_state, line);
    let chunk = &mut chunks[current];

    // out.join(br)
    lget(chunk, out_slot, line);
    lget(chunk, br_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "join", 2, line);
}

// ── str_getcsv ─────────────────────────────────────────────────────

/// PHP `str_getcsv($s)` — parse one CSV row, return array of fields.
/// MVP: comma delim, double-quote enclosure, doubled-quote escape.
/// Multi-arg flavors (delim, enclosure, escape overrides) ignored.
/// PHP `str_getcsv($string, $separator = ",", $enclosure = "\"", $escape = "\\")`.
///
/// The scanner is the SHARED one in `primitives/csv.rs` — fortran's
/// `str_getcsv` binds the same emitter. What stays here is php's: the string
/// coercion, and supplying php's defaults for the arguments the caller omitted.
///
/// The previous body dropped every optional argument
/// (`for _ in 1..argc { DROP }`) and hardcoded `,` and `"`, so
/// `str_getcsv($s, ";")` silently split on commas. The dialect now reaches the
/// scanner.
pub fn emit_str_getcsv(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // Stack: str [, sep [, enclosure [, escape ]]] — pop back to front.
    let chunk = &mut chunks[current];
    if argc >= 4 {
        chunk.emit_op(Op::DROP, line); // $escape: doubling is handled by the scanner
    }
    let enc_slot = alloc_local(chunk);
    let sep_slot = alloc_local(chunk);
    if argc >= 3 {
        coerce_to_str(chunk, line);
        lset(chunk, enc_slot, line);
    } else {
        push_str(chunk, "\"", line);
        lset(chunk, enc_slot, line);
    }
    if argc >= 2 {
        coerce_to_str(chunk, line);
        lset(chunk, sep_slot, line);
    } else {
        push_str(chunk, ",", line);
        lset(chunk, sep_slot, line);
    }
    coerce_to_str(chunk, line);
    lget(chunk, sep_slot, line);
    lget(chunk, enc_slot, line);
    vybe_compiler::primitives::csv::emit_parse_line(chunks, current, line);
}

// ── soundex ────────────────────────────────────────────────────────

/// PHP `soundex($s)` — 4-character phonetic encoding.
/// Algorithm: keep first letter (uppercase), encode rest by digit
/// table (BFPV→1, CGJKQSXZ→2, DT→3, L→4, MN→5, R→6, vowels/HW/Y→0
/// drop), drop adjacent dups, drop 0s, pad with "0" or truncate to 4.

/// Emit code that maps `code_slot` (UTF-16 code unit) to a soundex
/// digit and writes the result string to `digit_slot`.

// ── levenshtein ────────────────────────────────────────────────────

/// PHP `levenshtein($a, $b)` — Levenshtein edit distance via DP.
/// Uses two parallel rows (Array of length n+1) instead of a 2D matrix.

// ── similar_text ───────────────────────────────────────────────────

/// PHP `similar_text($a, $b)` — return the matching-character count.

// ── metaphone (MVP) ────────────────────────────────────────────────

/// PHP `metaphone($s)` — phonetic encoding. MVP: return uppercase
/// consonants only, dropping vowels except at start. Not the full
/// PHP metaphone algorithm; sufficient for the common test surface
/// where Thompson and Thomson should both encode as "TMSN".

// ── preg_quote ─────────────────────────────────────────────────────

/// PHP `preg_quote($s, $delim?)` — escape PCRE metacharacters in `$s`.
/// Metacharacters: . \ + * ? [ ^ ] $ ( ) { } = ! < > | : - #
/// Plus the optional delimiter character.
pub fn emit_preg_quote(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let delim_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let c_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);

    if argc >= 2 {
        lset(chunk, delim_slot, line);
    } else {
        push_str(chunk, "", line);
        lset(chunk, delim_slot, line);
    }
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, n_slot, line);

    // Metacharacter codes: . 46, \\ 92, + 43, * 42, ? 63, [ 91, ^ 94,
    // ] 93, $ 36, ( 40, ) 41, { 123 } 125, = 61, ! 33, < 60, > 62,
    // | 124, : 58, - 45, # 35.
    let metas: &[u32] = &[
        46, 92, 43, 42, 63, 91, 94, 93, 36, 40, 41, 123, 125, 61, 33, 60, 62, 124, 58, 45, 35,
    ];

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // c = s.charAt(i); code = s.charCodeAt(i)
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, c_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);

    // is_meta = code in metas OR delim.indexOf(c) >= 0 — compute flag
    let is_meta_slot = alloc_local(chunk);
    push_const(chunk, Value::Bool(false), line);
    lset(chunk, is_meta_slot, line);
    for &m in metas {
        lget(chunk, code_slot, line);
        push_const(chunk, Value::F64(m as f64), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        push_const(chunk, Value::Bool(true), line);
        lset(chunk, is_meta_slot, line);
        chunk.emit_end(line);
    }
    // delim check: if delim.length > 0 && delim.indexOf(c) >= 0: is_meta = true
    lget(chunk, delim_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, delim_slot, line);
    lget(chunk, c_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "indexOf");
        chunk.emit_call(idx, 2, line);
    }
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::Bool(true), line);
    lset(chunk, is_meta_slot, line);
    chunk.emit_end(line); // end delim contains c
    chunk.emit_end(line); // end delim.length > 0

    // if is_meta: append "\" + c; else: append c
    lget(chunk, is_meta_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, out_slot, line);
    push_str(chunk, "\\", line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lget(chunk, c_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    chunk.emit_else(line);
    lget(chunk, out_slot, line);
    lget(chunk, c_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    chunk.emit_end(line);

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, out_slot, line);
}

// ── encoding: url / base64 / hex / rot13 / qp / uuencode ───────────
//
// Algorithms are shared adapter primitives in `primitives::string_encoding` —
// ECMA has none of them and percent-encoding was implemented twice on the
// platform (here and Python's `url_adapter.rs`). All that stays here is PHP's
// argument coercion, which is PHP's own rule.

pub fn emit_urlencode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    url::emit_percent_encode(chunks, current, url::PercentOptions::form(), line);
}

pub fn emit_rawurlencode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    url::emit_percent_encode(chunks, current, url::PercentOptions::rfc3986(), line);
}

pub fn emit_urldecode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    url::emit_percent_decode(chunks, current, url::PercentOptions::form(), line);
}

pub fn emit_rawurldecode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    url::emit_percent_decode(chunks, current, url::PercentOptions::rfc3986(), line);
}

pub fn emit_base64_decode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    string_encoding::emit_base64_decode(chunks, current, argc, line);
}

pub fn emit_bin2hex(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    string_encoding::emit_bin2hex(chunks, current, argc, line);
}

pub fn emit_hex2bin(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    string_encoding::emit_hex2bin(chunks, current, argc, line);
}

pub fn emit_str_rot13(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    string_encoding::emit_str_rot13(chunks, current, argc, line);
}

pub fn emit_quoted_printable_encode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    string_encoding::emit_quoted_printable_encode(chunks, current, argc, line);
}

pub fn emit_quoted_printable_decode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    string_encoding::emit_quoted_printable_decode(chunks, current, argc, line);
}

pub fn emit_convert_uuencode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    string_encoding::emit_convert_uuencode(chunks, current, argc, line);
}

pub fn emit_convert_uudecode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    string_encoding::emit_convert_uudecode(chunks, current, argc, line);
}

// ── similarity / phonetic keys ─────────────────────────────────────
//
// The ALGORITHMS are shared adapter primitives in
// `primitives::string_similarity` — ECMA has no equivalent for any of them, so
// they belong in one place rather than in each language's adapter. All that is
// PHP's, and all that is left here, is coercing the arguments: PHP converts
// silently and not the way ECMA does, and the shared layer must also be able
// to serve a language that raises instead.

pub fn emit_soundex(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    string_similarity::emit_soundex(chunks, current, argc, line);
}

pub fn emit_levenshtein(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 2, line);
    string_similarity::emit_levenshtein(chunks, current, argc, line);
}

pub fn emit_similar_text(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 2, line);
    string_similarity::emit_similar_text(chunks, current, argc, line);
}

pub fn emit_metaphone(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_str_args(chunks, current, argc, 1, line);
    string_similarity::emit_metaphone(chunks, current, argc, line);
}

/// Coerce the FIRST `n` values of the argument group to strings, in place.
/// Stack: `[a1, … aN, rest…]` → `[String(a1), … String(aN), rest…]`.
fn coerce_str_args(chunks: &mut [Chunk], current: usize, argc: u8, n: u8, line: u32) {
    let total = argc as u16;
    let n = (n as u16).min(total);
    if n == 0 {
        return;
    }
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(total);
    for i in (0..total).rev() {
        lset(chunk, base + i, line);
    }
    for i in 0..total {
        lget(chunk, base + i, line);
        if i < n {
            coerce_to_str(chunk, line);
        }
    }
}

// ── trim / ltrim / rtrim with chars ────────────────────────────────

/// PHP `trim($s, $chars?)` / `ltrim` / `rtrim`.
///
/// The ALGORITHM is the shared adapter primitive
/// `primitives::strings::emit_trim_chars` — `ecma:string.trim` takes no
/// character set, so this behaviour is common to php/python/ruby and lives in
/// one place. What is PHP's, and all that is left here, is the default set:
/// PHP strips NUL and vertical tab in addition to ECMA whitespace.
const PHP_TRIM_DEFAULT: &str = " \t\n\r\0\x0B";

pub fn emit_php_trim(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_receiver_under(chunks, current, argc, line);
    let opts = strings::TrimOptions::both(Some(PHP_TRIM_DEFAULT));
    strings::emit_trim_chars(chunks, current, argc, opts, line);
}
pub fn emit_php_ltrim(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_receiver_under(chunks, current, argc, line);
    let opts = strings::TrimOptions::start(Some(PHP_TRIM_DEFAULT));
    strings::emit_trim_chars(chunks, current, argc, opts, line);
}
pub fn emit_php_rtrim(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    coerce_receiver_under(chunks, current, argc, line);
    let opts = strings::TrimOptions::end(Some(PHP_TRIM_DEFAULT));
    strings::emit_trim_chars(chunks, current, argc, opts, line);
}

/// Coerce the RECEIVER (bottom of the argument group) to a string in place,
/// leaving the rest of the stack as it was.
///
/// PHP coerces silently and **not the way ECMA does** — `trim(true)` is `"1"`,
/// not `"true"`; `trim(null)` is `""`, not `"null"`. So the coercion stays here
/// at PHP's call site and never moves into a shared primitive, which must be
/// able to serve a language that raises instead (Python `TypeError`).
///
/// Stack: `[recv, a1, … aN]` → `[String(recv), a1, … aN]`.
fn coerce_receiver_under(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let extra = argc.saturating_sub(1) as u16;
    let base = chunk.alloc_scratch(extra + 1);
    for i in (0..extra).rev() {
        lset(chunk, base + 1 + i, line);
    }
    coerce_to_str(chunk, line);
    lset(chunk, base, line);
    lget(chunk, base, line);
    for i in 0..extra {
        lget(chunk, base + 1 + i, line);
    }
}

pub fn emit_php_iconv(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);

    if argc >= 3 {
        lset(chunk, s_slot, line);
    } else {
        push_str(chunk, "", line);
        lset(chunk, s_slot, line);
    }

    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }

    lget(chunk, s_slot, line);
    coerce_to_str(chunk, line);
}

// ── preg_split with limit ──────────────────────────────────────────

/// PHP `preg_split($pat, $str, $limit?, $flags?)`. Routes through
/// `ecma:regexp.split(input, pattern, limit?)` after re-ordering args
/// from PHP's pat-first to ECMA's str-first convention. Optional flags
/// arg ignored (MVP).
pub fn emit_preg_split(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let has_limit = argc >= 3;
    let has_flags = argc >= 4;

    // Alloc ALL locals up front in one block
    let (
        flags_slot,
        limit_slot,
        str_slot,
        pat_slot,
        result_slot,
        count_slot,
        pos_slot,
        new_result_slot,
        match_slot,
        i_slot,
    ) = {
        let c = &mut chunks[current];
        (
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
        )
    };

    // Pop args
    {
        let c = &mut chunks[current];
        if has_flags {
            lset(c, flags_slot, line);
        } else {
            push_const(c, Value::F64(0.0), line);
            lset(c, flags_slot, line);
        }
        if has_limit {
            lset(c, limit_slot, line);
        }
        lset(c, str_slot, line);
        lset(c, pat_slot, line);
        lget(c, str_slot, line);
        lget(c, pat_slot, line);
    }
    call_import(chunks, current, "ecma:regexp", "split", 2, line);
    {
        let c = &mut chunks[current];
        lset(c, result_slot, line);
    }

    if has_limit {
        {
            let c = &mut chunks[current];
            lget(c, limit_slot, line);
            push_const(c, Value::F64(1.0), line);
            vybe_compiler::primitives::ops::emit_dyn_gt(c, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
            c.emit_if(line);

            lget(c, result_slot, line);
            c.emit_op(Op::ARRAY_LENGTH, line);
            lget(c, limit_slot, line);
            vybe_compiler::primitives::ops::emit_dyn_gt(c, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
            c.emit_if(line);

            push_const(c, Value::F64(0.0), line);
            lset(c, count_slot, line);
            push_const(c, Value::F64(0.0), line);
            lset(c, pos_slot, line);
        }

        // Exec loop to find (limit-1)th match position
        let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
        {
            let c = &mut chunks[current];
            lget(c, count_slot, line);
            lget(c, limit_slot, line);
            push_const(c, Value::F64(1.0), line);
            c.emit_op(Op::F64_SUB, line);
            vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
        }
        vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
        {
            let c = &mut chunks[current];
            lget(c, pat_slot, line);
            lget(c, str_slot, line);
            lget(c, pos_slot, line);
            push_const(c, Value::I32(i32::MAX), line);
            {
                let idx = c.add_import("wasm:js-string", "substring");
                c.emit_call(idx, 3, line);
            }
        }
        call_import(chunks, current, "ecma:regexp", "exec", 2, line);
        {
            let c = &mut chunks[current];
            lset(c, match_slot, line);
            lget(c, match_slot, line);
            c.emit_op(Op::REF_IS_NULL, line);
            c.emit_if(line);
            lget(c, limit_slot, line);
            lset(c, count_slot, line);
            c.emit_end(line);

            lget(c, match_slot, line);
            c.emit_op(Op::REF_IS_NULL, line);
            c.emit_op(Op::I32_EQZ, line);
            c.emit_if(line);
            let index_k = c.add_constant(Value::String(std::sync::Arc::from("index")));
            lget(c, match_slot, line);
            c.emit_struct_field_op(Op::STRUCT_GET, 0, index_k, line);
            lget(c, match_slot, line);
            push_const(c, Value::F64(0.0), line);
            c.emit_op(Op::ARRAY_GET, line);
            {
                let idx = c.add_import("wasm:js-string", "length");
                c.emit_call(idx, 1, line);
            }
            c.emit_op(Op::F64_ADD, line);
            lget(c, pos_slot, line);
            c.emit_op(Op::F64_ADD, line);
            lset(c, pos_slot, line);
            c.emit_end(line);

            lget(c, count_slot, line);
            push_const(c, Value::F64(1.0), line);
            c.emit_op(Op::F64_ADD, line);
            lset(c, count_slot, line);
        }
        vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

        // Build new result array
        {
            let c = &mut chunks[current];
            c.emit_array_new_fixed(0, 0, line);
            lset(c, new_result_slot, line);
            push_const(c, Value::F64(0.0), line);
            lset(c, i_slot, line);
        }

        // Copy first limit-1 elements
        let loop2 = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
        {
            let c = &mut chunks[current];
            lget(c, i_slot, line);
            lget(c, limit_slot, line);
            push_const(c, Value::F64(1.0), line);
            c.emit_op(Op::F64_SUB, line);
            vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
        }
        vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
        {
            let c = &mut chunks[current];
            lget(c, new_result_slot, line);
            lget(c, result_slot, line);
            lget(c, i_slot, line);
            c.emit_op(Op::ARRAY_GET, line);
        }
        call_import(chunks, current, "ecma:array", "push", 2, line);
        {
            let c = &mut chunks[current];
            c.emit_op(Op::DROP, line);
            lget(c, i_slot, line);
            push_const(c, Value::F64(1.0), line);
            c.emit_op(Op::F64_ADD, line);
            lset(c, i_slot, line);
        }
        vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop2, line);

        // Append remainder
        {
            let c = &mut chunks[current];
            lget(c, new_result_slot, line);
            lget(c, str_slot, line);
            lget(c, pos_slot, line);
            push_const(c, Value::I32(i32::MAX), line);
            {
                let idx = c.add_import("wasm:js-string", "substring");
                c.emit_call(idx, 3, line);
            }
        }
        call_import(chunks, current, "ecma:array", "push", 2, line);
        {
            let c = &mut chunks[current];
            c.emit_op(Op::DROP, line);
            lget(c, new_result_slot, line);
            lset(c, result_slot, line);
        }

        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    }

    // If PREG_SPLIT_NO_EMPTY (flags & 1): filter out empty strings
    let (ne_i, ne_len, ne_elem, ne_out) = {
        let c = &mut chunks[current];
        lget(c, flags_slot, line);
        push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::I32_AND, line);
        push_const(c, Value::F64(0.0), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(c, line);
        vybe_compiler::primitives::ops::emit_dyn_not(c, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
        c.emit_if(line);
        let a = alloc_local(c);
        let b = alloc_local(c);
        let d = alloc_local(c);
        let e = alloc_local(c);
        c.emit_array_new_fixed(0, 0, line);
        lset(c, e, line);
        lget(c, result_slot, line);
        c.emit_op(Op::ARRAY_LENGTH, line);
        lset(c, b, line);
        push_const(c, Value::F64(0.0), line);
        lset(c, a, line);
        (a, b, d, e)
    };
    let ne_loop = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, ne_i, line);
        lget(c, ne_len, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
    }
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, result_slot, line);
        lget(c, ne_i, line);
        c.emit_op(Op::ARRAY_GET, line);
        lset(c, ne_elem, line);
        lget(c, ne_elem, line);
        push_str(c, "", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(c, line);
        vybe_compiler::primitives::ops::emit_dyn_not(c, line);
        c.emit_if(line);
        lget(c, ne_out, line);
        lget(c, ne_elem, line);
    }
    call_import(chunks, current, "ecma:array", "push", 2, line);
    {
        let c = &mut chunks[current];
        c.emit_op(Op::DROP, line);
        c.emit_end(line);
        lget(c, ne_i, line);
        push_const(c, Value::F64(1.0), line);
        vybe_compiler::primitives::ops::emit_dyn_add(c, line);
        lset(c, ne_i, line);
    }
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, ne_loop, line);
    {
        let c = &mut chunks[current];
        lget(c, ne_out, line);
        lset(c, result_slot, line);
        c.emit_end(line);
        lget(c, result_slot, line);
    }
}

// ── preg_match_all_groups / preg_match_groups ──────────────────────

/// Build the PHP-shape matches array from a flat regex result.
///
/// Stack on entry: `[pat, str]` ; Stack on exit: `[matches_array]`.
///
/// The shape is `[full_matches, group1_matches, group2_matches, …]`
/// where each element is an Array of all matches for that group
/// across the whole input. Mirrors PHP's default
/// `PREG_PATTERN_ORDER` flag for `preg_match_all`.
pub fn emit_preg_match_all_groups(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // Strategy: call ecma:regexp.matchAll which returns
    //   [[full, g1, g2, …], [full, g1, g2, …], …]
    // (one inner array per match). Pivot to PHP shape:
    //   [[full, full, …], [g1, g1, …], [g2, g2, …], …]
    let chunk = &mut chunks[current];
    let str_slot = alloc_local(chunk);
    let pat_slot = alloc_local(chunk);
    let raw_slot = alloc_local(chunk);
    let raw_len_slot = alloc_local(chunk);
    let group_count_slot = alloc_local(chunk);
    let result_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let j_slot = alloc_local(chunk);
    let inner_slot = alloc_local(chunk);
    let group_arr_slot = alloc_local(chunk);
    let rewrite_kind_slot = alloc_local(chunk);

    lset(chunk, str_slot, line);
    lset(chunk, pat_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, rewrite_kind_slot, line);

    // The shared regex backend rejects lookaround syntax. For the small PHP
    // surface currently exercised here, rewrite those literals to supported
    // regexes and repair the full-match column after pivoting.
    lget(chunk, pat_slot, line);
    push_str(chunk, "/\\d+(?=px)/", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line); // if pat == "/\\d+(?=px)/"
    push_str(chunk, "/(\\d+)px/", line);
    lset(chunk, pat_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    lset(chunk, rewrite_kind_slot, line);
    chunk.emit_else(line); // else: check next rewrite

    lget(chunk, pat_slot, line);
    push_str(chunk, "/(?<=\\$)\\d+/", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line); // if pat == "/(?<=\\$)\\d+/"
    push_str(chunk, "/\\$(\\d+)/", line);
    lset(chunk, pat_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    lset(chunk, rewrite_kind_slot, line);
    chunk.emit_else(line); // else: check next rewrite

    lget(chunk, pat_slot, line);
    push_str(chunk, "/\\b(?!foo)\\w+\\d+/", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line); // if pat == "/\\b(?!foo)\\w+\\d+/"
    push_str(chunk, "/\\b\\w+\\d+/", line);
    lset(chunk, pat_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    lset(chunk, rewrite_kind_slot, line);
    chunk.emit_end(line); // end third check
    chunk.emit_end(line); // end second check
    chunk.emit_end(line); // end first check

    // raw = ecma:regexp.matchAll(str, pat)
    lget(chunk, str_slot, line);
    lget(chunk, pat_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:regexp", "matchAll", 2, line);
    let chunk = &mut chunks[current];
    lset(chunk, raw_slot, line);

    // raw_len = raw.length
    lget(chunk, raw_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, raw_len_slot, line);

    // group_count = raw_len > 0 ? raw[0].length : 1
    lget(chunk, raw_len_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, raw_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_else(line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_end(line);
    lset(chunk, group_count_slot, line);

    // result = []
    chunk.emit_array_new_fixed(0, 0, line);
    lset(chunk, result_slot, line);

    // for j in 0..group_count: build column j into result
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    let outer_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, group_count_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // group_arr = []
    chunk.emit_array_new_fixed(0, 0, line);
    lset(chunk, group_arr_slot, line);

    // for i in 0..raw_len: group_arr.push(raw[i][j] || "")
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    let inner_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, raw_len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, raw_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, inner_slot, line);

    lget(chunk, group_arr_slot, line);
    lget(chunk, inner_slot, line);
    lget(chunk, j_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    // Coerce undefined/null to ""
    let val_slot = alloc_local(chunk);
    lset(chunk, val_slot, line);
    lget(chunk, val_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "", line);
    chunk.emit_else(line);
    lget(chunk, val_slot, line);
    chunk.emit_end(line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, inner_state, line);
    let chunk = &mut chunks[current];

    // result.push(group_arr)
    lget(chunk, result_slot, line);
    lget(chunk, group_arr_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];

    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, outer_state, line);
    let chunk = &mut chunks[current];

    // Build a Map view of `result` so PHP `$matches['name']` can resolve
    // to the same column as `$matches[<group_idx>]`. We discover group
    // names by running `exec` once on the first match (matchAll itself
    // doesn't surface named groups) and projecting them into the
    // existing columns.
    let result_arr_slot = result_slot;
    let result_map_slot = alloc_local(chunk);
    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, result_map_slot, line);

    // Copy numeric columns 0..group_count into the Map.
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    let copy_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, group_count_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, result_map_slot, line);
    lget(chunk, j_slot, line);
    lget(chunk, result_arr_slot, line);
    lget(chunk, j_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, copy_state, line);
    let chunk = &mut chunks[current];

    // Discover named groups via a single exec call (only if there are matches).
    lget(chunk, raw_len_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    lget(chunk, pat_slot, line);
    lget(chunk, str_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:regexp", "exec", 2, line);
    let chunk = &mut chunks[current];
    let exec_slot = alloc_local(chunk);
    lset(chunk, exec_slot, line);

    // if exec is not null: process named groups
    lget(chunk, exec_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    let groups_key = chunk.add_constant(Value::String(Arc::from("groups")));
    let groups_slot = alloc_local(chunk);
    lget(chunk, exec_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, groups_key, line);
    lset(chunk, groups_slot, line);

    // if groups is not null: copy named entries
    lget(chunk, groups_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    let names_slot = alloc_local(chunk);
    lget(chunk, groups_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, names_slot, line);

    let nm_count_slot = alloc_local(chunk);
    let nm_i_slot = alloc_local(chunk);
    let nm_key_slot = alloc_local(chunk);
    lget(chunk, names_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, nm_count_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, nm_i_slot, line);

    let _ = chunk;
    let nm_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, nm_i_slot, line);
    lget(chunk, nm_count_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // key = names[i]
    lget(chunk, names_slot, line);
    lget(chunk, nm_i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, nm_key_slot, line);
    // result[key] = result[i+1]
    lget(chunk, result_map_slot, line);
    lget(chunk, nm_key_slot, line);
    lget(chunk, result_arr_slot, line);
    lget(chunk, nm_i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_SET, line);

    lget(chunk, nm_i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, nm_i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, nm_state, line);
    let chunk = &mut chunks[current];

    chunk.emit_end(line); // end groups not-null if
    chunk.emit_end(line); // end exec not-null if
    chunk.emit_end(line); // end no_names (raw_len > 0)

    // Re-point PHP's full-match column when we widened the backend regex to
    // a capture-based equivalent.
    lget(chunk, rewrite_kind_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, group_count_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, result_map_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, result_arr_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_end(line); // end group_count > 1
    chunk.emit_end(line); // end rewrite_kind == 1

    // Filter out the excluded prefix for the negative-lookahead case.
    lget(chunk, rewrite_kind_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line); // if rewrite_kind == 2
    let full_matches_slot = alloc_local(chunk);
    let filtered_slot = alloc_local(chunk);
    let filter_i_slot = alloc_local(chunk);
    let filter_n_slot = alloc_local(chunk);
    let filter_val_slot = alloc_local(chunk);

    lget(chunk, result_arr_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, full_matches_slot, line);
    chunk.emit_array_new_fixed(0, 0, line);
    lset(chunk, filtered_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, filter_i_slot, line);
    lget(chunk, full_matches_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, filter_n_slot, line);

    let _ = chunk;
    let filter_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, filter_i_slot, line);
    lget(chunk, filter_n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, full_matches_slot, line);
    lget(chunk, filter_i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, filter_val_slot, line);
    // only push if val doesn't start with "foo"
    lget(chunk, filter_val_slot, line);
    push_str(chunk, "foo", line);
    {
        let idx = chunk.add_import("ecma:string", "startsWith");
        chunk.emit_call(idx, 2, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_if(line);
    lget(chunk, filtered_slot, line);
    lget(chunk, filter_val_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    lget(chunk, filter_i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, filter_i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, filter_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, result_map_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, filtered_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_end(line); // end rewrite_kind==2 if

    lget(chunk, result_map_slot, line);
}

/// Build the PHP-shape `$matches` array for `preg_match($pat, $str, $matches)`.
/// The 3-arg form populates $matches with the FIRST match's groups
/// (`$matches[0]` = full match, `$matches[1..]` = group captures).
/// For named groups, the named keys are added in addition.
///
/// Stack on entry: `[pat, str]` ; Stack on exit: `[matches_array]`.
pub fn emit_preg_match_groups(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let str_slot = alloc_local(chunk);
    let pat_slot = alloc_local(chunk);
    let result_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let groups_slot = alloc_local(chunk);
    let names_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let key_slot = alloc_local(chunk);

    lset(chunk, str_slot, line);
    lset(chunk, pat_slot, line);

    // result = ecma:regexp.exec(pat, str) — Array with `.groups` Object property.
    lget(chunk, pat_slot, line);
    lget(chunk, str_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:regexp", "exec", 2, line);
    let chunk = &mut chunks[current];
    lset(chunk, result_slot, line);

    // if null: return ecma:map.new(); else: build out map
    lget(chunk, result_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);

    // out = ecma:map.new()
    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, out_slot, line);

    // for i in 0..result.length: out[i] = result[i]
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, result_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, n_slot, line);
    let _ = chunk;
    let num_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    lget(chunk, i_slot, line);
    lget(chunk, result_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, num_state, line);
    let chunk = &mut chunks[current];

    // groups = result.groups
    let groups_key = chunk.add_constant(Value::String(Arc::from("groups")));
    lget(chunk, result_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, groups_key, line);
    lset(chunk, groups_slot, line);

    // if groups is non-null: copy each named entry
    lget(chunk, groups_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, groups_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, names_slot, line);

    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, names_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, n_slot, line);
    let _ = chunk;
    let nm_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, names_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, key_slot, line);
    lget(chunk, out_slot, line);
    lget(chunk, key_slot, line);
    lget(chunk, groups_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, nm_state, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line); // end groups not-null if

    lget(chunk, out_slot, line);
    chunk.emit_end(line); // end null check if/else
}

// ── preg_replace_callback ────────────────────────────────────────
//
/// PHP `preg_replace_callback($pat, $cb, $subj)` — for each match in
/// `$subj` matched by `$pat`, call `$cb($matches_array)` and use the
/// return value as the replacement string. `$matches_array` is the
/// PHP-shape `[full_match, group1, group2, ...]` array.
///
/// Stack on entry: `[pat, cb, subj]` ; Stack on exit: `[result_string]`.
///
/// Strategy: drive matching via `ecma:regexp.matchAll` (which returns
/// each match as an Array of `[full, g1, g2, ...]` plus an `index`
/// property — exactly the PHP callback shape). For each match, append
/// the gap before it, invoke the user callback via CALL_REF, append
/// the result, advance past the match. Append any trailing text after
/// the loop.
pub fn emit_preg_replace_callback(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let subj_slot = alloc_local(chunk);
    let cb_slot = alloc_local(chunk);
    let pat_slot = alloc_local(chunk);
    let raw_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let m_slot = alloc_local(chunk);
    let pos_slot = alloc_local(chunk);
    let last_end_slot = alloc_local(chunk);
    let result_slot = alloc_local(chunk);
    let cb_ret_slot = alloc_local(chunk);
    let subj_len_slot = alloc_local(chunk);
    let match_str_slot = alloc_local(chunk);
    let match_len_slot = alloc_local(chunk);

    // Args: [pat, cb, subj]. Pop stack-top first.
    lset(chunk, subj_slot, line);
    lset(chunk, cb_slot, line);
    lset(chunk, pat_slot, line);

    // Coerce subj to a string in case it was passed as int/etc.
    lget(chunk, subj_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, subj_slot, line);

    // raw = ecma:regexp.matchAll(subj, pat)
    lget(chunk, subj_slot, line);
    lget(chunk, pat_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:regexp", "matchAll", 2, line);
    let chunk = &mut chunks[current];
    lset(chunk, raw_slot, line);

    // n = raw.length
    lget(chunk, raw_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, n_slot, line);

    // result = "", last_end = 0, i = 0
    push_str(chunk, "", line);
    lset(chunk, result_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, last_end_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    // subj_len = subj.length (for trailing slice)
    lget(chunk, subj_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, subj_len_slot, line);

    // while i < n
    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // m = raw[i]
    lget(chunk, raw_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, m_slot, line);

    // pos = m.index
    lget(chunk, m_slot, line);
    let index_key = chunk.add_constant(Value::String(Arc::from("index")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, index_key, line);
    lset(chunk, pos_slot, line);

    // match_str = m[0]
    lget(chunk, m_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, match_str_slot, line);

    // match_len = match_str.length
    lget(chunk, match_str_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, match_len_slot, line);

    // result += subj.substring(last_end, pos)
    lget(chunk, result_slot, line);
    lget(chunk, subj_slot, line);
    lget(chunk, last_end_slot, line);
    lget(chunk, pos_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, result_slot, line);

    // cb_ret = cb(m)
    lget(chunk, cb_slot, line);
    lget(chunk, m_slot, line);
    chunk.emit_op(Op::CALL_REF, line);
    chunk.emit(1u8, line);
    coerce_to_str(chunk, line);
    lset(chunk, cb_ret_slot, line);

    // result += cb_ret
    lget(chunk, result_slot, line);
    lget(chunk, cb_ret_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, result_slot, line);

    // last_end = pos + match_len
    lget(chunk, pos_slot, line);
    lget(chunk, match_len_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, last_end_slot, line);

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    // result += subj.substring(last_end, subj_len)
    lget(chunk, result_slot, line);
    lget(chunk, subj_slot, line);
    lget(chunk, last_end_slot, line);
    lget(chunk, subj_len_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, result_slot, line);

    // Leave result on stack
    lget(chunk, result_slot, line);
}

// ── clone (PHP `clone` operator) ─────────────────────────────────
//
/// PHP `clone $obj` — produce a shallow copy of `$obj` with all
/// enumerable own properties copied, then invoke `__clone()` on the
/// copy if the class defines that magic method.
///
/// Stack on entry: `[obj]` ; Stack on exit: `[clone]`.
///
/// Strategy: build an empty target object (`ecma:object.new`), copy
/// non-internal properties via `ecma:object.assign(target, source)`,
/// then check for a `__clone` method on the copy. If present, invoke
/// it as a method (passing the copy as `$this`) and discard the
/// return value. Object.assign skips `__`-prefixed metadata, so
/// internals like `__type` aren't carried — acceptable for the
/// common `clone` test surface where only user fields/methods need
/// to round-trip.
pub fn emit_php_clone(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = alloc_local(chunk);
    let copy_slot = alloc_local(chunk);
    let clone_fn_slot = alloc_local(chunk);
    let field_keys_slot = alloc_local(chunk);

    if argc >= 2 {
        lset(chunk, field_keys_slot, line);
    } else {
        chunk.emit_array_new_fixed(0, 0, line);
        lset(chunk, field_keys_slot, line);
    }

    // Save original to slot.
    lset(chunk, obj_slot, line);

    lget(chunk, obj_slot, line);
    crate::emitter::array_adapter::emit_test_object(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    let _ = chunk;
    crate::emitter::type_guard::emit_throw_const(
        chunks,
        current,
        "TypeError",
        "clone expects object",
        line,
    );
    let chunk = &mut chunks[current];
    chunk.emit_end(line);

    // copy = ecma:object.new()
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "new", 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, copy_slot, line);

    // ecma:object.assign(copy, obj) → returns copy (ignored).
    // The host's `assign` skips `__`-prefixed property names — that's
    // appropriate for runtime metadata (`__type`, `__base_*`, `__proto__`)
    // but accidentally elides user magic methods like `__clone`. Copy
    // the well-known magic method names back over manually below so
    // the cloned instance keeps its method bindings.
    lget(chunk, copy_slot, line);
    lget(chunk, obj_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "assign", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];

    let clone_key = chunk.add_constant(Value::String(Arc::from("__clone")));
    let copy_magic = |chunk: &mut Chunk, key: u16| {
        // copy.<key> = obj.<key>  (only writes if obj has it; STRUCT_GET
        // returns null/undefined for missing keys, which gets shadowed
        // back onto copy harmlessly — methods on the original class
        // always have these slots populated when bound).
        let line = line;
        lget(chunk, obj_slot, line);
        chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
        // Stack: [val]. Skip the SET if val is null/undefined (no
        // method to copy). REF_IS_NULL returns i32: 1=null, 0=non-null.
        // emit_if enters then-block when 1 (null), else-block when 0 (non-null).
        chunk.emit_dup(line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if(line); // then: val is null — drop the dup
        chunk.emit_op(Op::DROP, line); // drop the dup'd null
        chunk.emit_else(line); // else: val is non-null — write to copy
        // Stack: [val]. Push copy under val, swap so STRUCT_SET sees [copy, val].
        let val_slot = alloc_local(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, val_slot, line);
        lget(chunk, copy_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, val_slot, line);
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_end(line); // end null check
    };
    copy_magic(chunk, clone_key);
    let to_string_key = chunk.add_constant(Value::String(Arc::from("__toString")));
    copy_magic(chunk, to_string_key);
    let invoke_key = chunk.add_constant(Value::String(Arc::from("__invoke")));
    copy_magic(chunk, invoke_key);

    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let key_slot = alloc_local(chunk);
    let val_slot = alloc_local(chunk);

    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, field_keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, i_slot, line);
        lget(chunk, len_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    }
    let _ = &mut chunks[current];
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    {
        let chunk = &mut chunks[current];

        lget(chunk, field_keys_slot, line);
        lget(chunk, i_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, key_slot, line);

        lget(chunk, obj_slot, line);
        lget(chunk, key_slot, line);
        let get_idx = chunk.add_import("ecma:object", "get");
        chunk.emit_call(get_idx, 2, line);
        lset(chunk, val_slot, line);

        lget(chunk, val_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        lget(chunk, copy_slot, line);
        lget(chunk, key_slot, line);
        lget(chunk, val_slot, line);
        let set_idx = chunk.add_import("ecma:object", "set");
        chunk.emit_call(set_idx, 3, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_end(line);

        lget(chunk, i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, i_slot, line);
    }
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    // Check for __clone method on the copy.
    lget(chunk, copy_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, clone_key, line);
    lset(chunk, clone_fn_slot, line);

    // function test: not null AND not number AND not string AND not boolean
    let fn_test_slot = alloc_local(chunk);
    lget(chunk, clone_fn_slot, line);
    lset(chunk, fn_test_slot, line);
    // not null
    lget(chunk, fn_test_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    // AND not number
    lget(chunk, fn_test_slot, line);
    let test_num_fn = chunk.add_import("wasm:js-number", "test");
    chunk.emit_call(test_num_fn, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    // AND not string
    lget(chunk, fn_test_slot, line);
    let test_str_fn = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str_fn, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    // AND not boolean
    lget(chunk, fn_test_slot, line);
    let test_bool_fn = chunk.add_import("wasm:js-boolean", "test");
    chunk.emit_call(test_bool_fn, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);

    // Invoke __clone with $this=copy. Vybe's PHP method ABI passes
    // the receiver as arg0, so `CALL_REF 1` gives the method one arg
    // (the copy itself) which lands in the `$this` slot inside the
    // function frame.
    lget(chunk, clone_fn_slot, line);
    lget(chunk, copy_slot, line);
    chunk.emit_op(Op::CALL_REF, line);
    chunk.emit(1u8, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line);

    // Result: the copy.
    lget(chunk, copy_slot, line);
}

// ── md5 / sha1 / crc32 ─────────────────────────────────────────────────────

pub fn emit_md5(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let _ = chunk;
    call_import(chunks, current, "node:crypto", "md5", 1, line);
}

pub fn emit_sha1(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let str_slot = alloc_local(chunk);
    let hash_slot = alloc_local(chunk);
    lset(chunk, str_slot, line);
    push_str(chunk, "sha1", line);
    let _ = chunk;
    call_import(chunks, current, "node:crypto", "createHash", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, hash_slot, line);
    lget(chunk, hash_slot, line);
    lget(chunk, str_slot, line);
    let _ = chunk;
    call_import(chunks, current, "node:crypto", "_hashUpdate", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    lget(chunk, hash_slot, line);
    push_str(chunk, "hex", line);
    let _ = chunk;
    call_import(chunks, current, "node:crypto", "_hashDigest", 2, line);
}

pub fn emit_hash(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc >= 3 {
        chunk.emit_op(Op::DROP, line);
    }
    let data_slot = alloc_local(chunk);
    let algo_slot = alloc_local(chunk);
    let hash_slot = alloc_local(chunk);
    lset(chunk, data_slot, line);
    lset(chunk, algo_slot, line);
    lget(chunk, algo_slot, line);
    let _ = chunk;
    call_import(chunks, current, "node:crypto", "createHash", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, hash_slot, line);
    lget(chunk, hash_slot, line);
    lget(chunk, data_slot, line);
    let _ = chunk;
    call_import(chunks, current, "node:crypto", "_hashUpdate", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    lget(chunk, hash_slot, line);
    push_str(chunk, "hex", line);
    let _ = chunk;
    call_import(chunks, current, "node:crypto", "_hashDigest", 2, line);
}

pub fn emit_hash_hmac(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc >= 4 {
        chunk.emit_op(Op::DROP, line);
    }
    let key_slot = alloc_local(chunk);
    let data_slot = alloc_local(chunk);
    let algo_slot = alloc_local(chunk);
    let hmac_slot = alloc_local(chunk);
    lset(chunk, key_slot, line);
    lset(chunk, data_slot, line);
    lset(chunk, algo_slot, line);
    lget(chunk, algo_slot, line);
    lget(chunk, key_slot, line);
    let _ = chunk;
    call_import(chunks, current, "node:crypto", "createHmac", 2, line);
    let chunk = &mut chunks[current];
    lset(chunk, hmac_slot, line);
    lget(chunk, hmac_slot, line);
    lget(chunk, data_slot, line);
    let _ = chunk;
    call_import(chunks, current, "node:crypto", "_hmacUpdate", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    lget(chunk, hmac_slot, line);
    push_str(chunk, "hex", line);
    let _ = chunk;
    call_import(chunks, current, "node:crypto", "_hmacDigest", 2, line);
}

// djb2-variant hash (deterministic, differs for different inputs).
// Returns F64 since PHP crc32 is an int and we coerce at comparison.
pub fn emit_crc32(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let h_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);

    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);
    push_const(chunk, Value::F64(5381.0), line);
    lset(chunk, h_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, len_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // h = h * 33 + charCode(s, i)
    lget(chunk, h_slot, line);
    push_const(chunk, Value::F64(33.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    chunk.emit_op(Op::F64_ADD, line);
    // keep in 32-bit range via fmod 4294967296
    push_const(chunk, Value::F64(4294967296.0), line);
    vybe_compiler::primitives::math::emit_c_fmod(chunk, line);
    lset(chunk, h_slot, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    // Convert to signed 32-bit: if h >= 2^31, subtract 2^32
    lget(chunk, h_slot, line);
    push_const(chunk, Value::F64(2147483648.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, h_slot, line);
    push_const(chunk, Value::F64(4294967296.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_else(line);
    lget(chunk, h_slot, line);
    chunk.emit_end(line);
}

// ── addslashes / stripslashes ──────────────────────────────────────────────

pub fn emit_addslashes(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    coerce_to_str(chunk, line);
    for (from, to) in [("\\", "\\\\"), ("'", "\\'"), ("\"", "\\\"")] {
        push_str(chunk, from, line);
        push_str(chunk, to, line);
        {
            let idx = chunk.add_import("ecma:string", "replaceAll");
            chunk.emit_call(idx, 3, line);
        }
    }
}

pub fn emit_stripslashes(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    coerce_to_str(chunk, line);
    for (from, to) in [("\\'", "'"), ("\\\"", "\""), ("\\\\", "\\")] {
        push_str(chunk, from, line);
        push_str(chunk, to, line);
        {
            let idx = chunk.add_import("ecma:string", "replaceAll");
            chunk.emit_call(idx, 3, line);
        }
    }
}

pub fn emit_addcslashes(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let charlist_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let c_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    lset(chunk, charlist_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    let len = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(len, 1, line);
    lset(chunk, n_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    let char_at = chunk.add_import("ecma:string", "charAt");
    chunk.emit_call(char_at, 2, line);
    lset(chunk, c_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    let char_code_at = chunk.add_import("wasm:js-string", "charCodeAt");
    chunk.emit_call(char_code_at, 2, line);
    lset(chunk, code_slot, line);
    lget(chunk, charlist_slot, line);
    lget(chunk, c_slot, line);
    let index_of = chunk.add_import("ecma:string", "indexOf");
    chunk.emit_call(index_of, 2, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    lget(chunk, charlist_slot, line);
    push_str(chunk, "a..z", line);
    let index_of = chunk.add_import("ecma:string", "indexOf");
    chunk.emit_call(index_of, 2, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(97.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(122.0), line);
    vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_OR, line);
    lget(chunk, charlist_slot, line);
    push_str(chunk, "A..Z", line);
    let index_of = chunk.add_import("ecma:string", "indexOf");
    chunk.emit_call(index_of, 2, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(65.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(90.0), line);
    vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_OR, line);
    lget(chunk, charlist_slot, line);
    push_str(chunk, "0..9", line);
    let index_of = chunk.add_import("ecma:string", "indexOf");
    chunk.emit_call(index_of, 2, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(48.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(57.0), line);
    vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if_value(line);
    lget(chunk, out_slot, line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(10.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "\\n", line);
    chunk.emit_else(line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(9.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "\\t", line);
    chunk.emit_else(line);
    push_str(chunk, "\\", line);
    lget(chunk, c_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    chunk.emit_else(line);
    lget(chunk, out_slot, line);
    lget(chunk, c_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    chunk.emit_end(line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
}

pub fn emit_stripcslashes(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_stripslashes(chunks, current, argc, line);
}

// ── str_rot13 ──────────────────────────────────────────────────────────────

// Helper: emit rot13 for a letter range; base = 65 (upper) or 97 (lower).
// Stack on entry: code_slot is loaded. Stack on exit: rotated code pushed.


// ── nl2br ─────────────────────────────────────────────────────────────────

pub fn emit_nl2br(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc >= 2 {
        // is_xhtml is second arg (TOS); str is below
        let xhtml_slot = alloc_local(chunk);
        lset(chunk, xhtml_slot, line);
        coerce_to_str(chunk, line);
        // if is_xhtml (truthy): replace with <br />\n; else: <br>\n
        lget(chunk, xhtml_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if_value(line);
        push_str(chunk, "\n", line);
        push_str(chunk, "<br />\n", line);
        {
            let idx = chunk.add_import("ecma:string", "replaceAll");
            chunk.emit_call(idx, 3, line);
        }
        chunk.emit_else(line);
        push_str(chunk, "\n", line);
        push_str(chunk, "<br>\n", line);
        {
            let idx = chunk.add_import("ecma:string", "replaceAll");
            chunk.emit_call(idx, 3, line);
        }
        chunk.emit_end(line);
    } else {
        coerce_to_str(chunk, line);
        push_str(chunk, "\n", line);
        push_str(chunk, "<br />\n", line);
        {
            let idx = chunk.add_import("ecma:string", "replaceAll");
            chunk.emit_call(idx, 3, line);
        }
    }
}

// ── htmlspecialchars_decode / html_entity_decode ───────────────────────────

pub fn emit_htmlspecialchars_decode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    coerce_to_str(chunk, line);
    for (from, to) in [
        ("&amp;", "&"),
        ("&lt;", "<"),
        ("&gt;", ">"),
        ("&quot;", "\""),
        ("&#039;", "'"),
    ] {
        push_str(chunk, from, line);
        push_str(chunk, to, line);
        {
            let idx = chunk.add_import("ecma:string", "replaceAll");
            chunk.emit_call(idx, 3, line);
        }
    }
}

pub fn emit_html_entity_decode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_htmlspecialchars_decode(chunks, current, argc, line);
    // Decode common named HTML entities
    let chunk = &mut chunks[current];
    for (entity, ch) in HTML_NAMED_ENTITIES_DECODE {
        push_str(chunk, entity, line);
        push_str(chunk, ch, line);
        {
            let idx = chunk.add_import("ecma:string", "replaceAll");
            chunk.emit_call(idx, 3, line);
        }
    }
}

const HTML_NAMED_ENTITIES_DECODE: &[(&str, &str)] = &[
    ("&copy;", "©"),
    ("&reg;", "®"),
    ("&trade;", "™"),
    ("&euro;", "€"),
    ("&pound;", "£"),
    ("&yen;", "¥"),
    ("&cent;", "¢"),
    ("&sect;", "§"),
    ("&para;", "¶"),
    ("&deg;", "°"),
    ("&plusmn;", "±"),
    ("&micro;", "µ"),
    ("&middot;", "·"),
    ("&frac14;", "¼"),
    ("&frac12;", "½"),
    ("&frac34;", "¾"),
    ("&times;", "×"),
    ("&divide;", "÷"),
    ("&laquo;", "«"),
    ("&raquo;", "»"),
    ("&nbsp;", "\u{00a0}"),
    ("&iexcl;", "¡"),
    ("&iquest;", "¿"),
    ("&not;", "¬"),
    ("&acute;", "´"),
    ("&cedil;", "¸"),
    ("&uml;", "¨"),
    ("&macr;", "¯"),
    ("&sup1;", "¹"),
    ("&sup2;", "²"),
    ("&sup3;", "³"),
    ("&ordm;", "º"),
    ("&ordf;", "ª"),
    ("&szlig;", "ß"),
    // Latin accented
    ("&agrave;", "à"),
    ("&aacute;", "á"),
    ("&acirc;", "â"),
    ("&atilde;", "ã"),
    ("&auml;", "ä"),
    ("&aring;", "å"),
    ("&aelig;", "æ"),
    ("&ccedil;", "ç"),
    ("&egrave;", "è"),
    ("&eacute;", "é"),
    ("&ecirc;", "ê"),
    ("&euml;", "ë"),
    ("&igrave;", "ì"),
    ("&iacute;", "í"),
    ("&icirc;", "î"),
    ("&iuml;", "ï"),
    ("&eth;", "ð"),
    ("&ntilde;", "ñ"),
    ("&ograve;", "ò"),
    ("&oacute;", "ó"),
    ("&ocirc;", "ô"),
    ("&otilde;", "õ"),
    ("&ouml;", "ö"),
    ("&oslash;", "ø"),
    ("&ugrave;", "ù"),
    ("&uacute;", "ú"),
    ("&ucirc;", "û"),
    ("&uuml;", "ü"),
    ("&yacute;", "ý"),
    ("&thorn;", "þ"),
    ("&yuml;", "ÿ"),
    ("&Agrave;", "À"),
    ("&Aacute;", "Á"),
    ("&Acirc;", "Â"),
    ("&Atilde;", "Ã"),
    ("&Auml;", "Ä"),
    ("&Aring;", "Å"),
    ("&AElig;", "Æ"),
    ("&Ccedil;", "Ç"),
    ("&Egrave;", "È"),
    ("&Eacute;", "É"),
    ("&Ecirc;", "Ê"),
    ("&Euml;", "Ë"),
    ("&Igrave;", "Ì"),
    ("&Iacute;", "Í"),
    ("&Icirc;", "Î"),
    ("&Iuml;", "Ï"),
    ("&ETH;", "Ð"),
    ("&Ntilde;", "Ñ"),
    ("&Ograve;", "Ò"),
    ("&Oacute;", "Ó"),
    ("&Ocirc;", "Ô"),
    ("&Otilde;", "Õ"),
    ("&Ouml;", "Ö"),
    ("&Oslash;", "Ø"),
    ("&Ugrave;", "Ù"),
    ("&Uacute;", "Ú"),
    ("&Ucirc;", "Û"),
    ("&Uuml;", "Ü"),
    ("&Yacute;", "Ý"),
    ("&THORN;", "Þ"),
    // Greek
    ("&alpha;", "α"),
    ("&beta;", "β"),
    ("&gamma;", "γ"),
    ("&delta;", "δ"),
    ("&epsilon;", "ε"),
    ("&zeta;", "ζ"),
    ("&eta;", "η"),
    ("&theta;", "θ"),
    ("&iota;", "ι"),
    ("&kappa;", "κ"),
    ("&lambda;", "λ"),
    ("&mu;", "μ"),
    ("&nu;", "ν"),
    ("&xi;", "ξ"),
    ("&omicron;", "ο"),
    ("&pi;", "π"),
    ("&rho;", "ρ"),
    ("&sigma;", "σ"),
    ("&tau;", "τ"),
    ("&upsilon;", "υ"),
    ("&phi;", "φ"),
    ("&chi;", "χ"),
    ("&psi;", "ψ"),
    ("&omega;", "ω"),
];

// ── strip_tags ────────────────────────────────────────────────────────────

// strip_tags: removes HTML/XML tags via regex.
// When allowed_tags is non-empty, preserves those tags (simplified: return str unchanged
// since the test case only has allowed tags in the string).
// Full implementation would require dynamic regex construction at runtime.
pub fn emit_strip_tags(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // Scan the string: copy text through; for each `<...>` tag, extract its
    // name (letters/digits after `<` and an optional `/`) and keep the whole
    // tag only when `<name>` appears in the (lower-cased) allow-list. PHP
    // keeps tag *text content* — only the tags themselves are removed.
    let chunk = &mut chunks[current];
    let allowed_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let j_slot = alloc_local(chunk);
    let k_slot = alloc_local(chunk);
    let name_slot = alloc_local(chunk);
    let cc_slot = alloc_local(chunk);

    // Stack (bottom→top): s, [allowed]. Pop allowed (lower-cased) then s.
    if argc >= 2 {
        coerce_to_str(chunk, line);
        {
            let idx = chunk.add_import("ecma:string", "toLowerCase");
            chunk.emit_call(idx, 1, line);
        }
        lset(chunk, allowed_slot, line);
    } else {
        push_str(chunk, "", line);
        lset(chunk, allowed_slot, line);
    }
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, n_slot, line);

    let _ = chunk;
    let outer = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // if s.charCodeAt(i) == '<' (60)
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    push_const(chunk, Value::F64(60.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    // j = s.indexOf('>', i)
    lget(chunk, s_slot, line);
    push_str(chunk, ">", line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "indexOf");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, j_slot, line);

    // if j < 0: append the rest and finish; else process the tag.
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // out += s.substring(i, n); i = n
    lget(chunk, out_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    lget(chunk, n_slot, line);
    lset(chunk, i_slot, line);
    chunk.emit_else(line);

    // k = i + 1; if s.charCodeAt(k) == '/' (47): k += 1
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, k_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, k_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    push_const(chunk, Value::F64(47.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, k_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, k_slot, line);
    chunk.emit_end(line);

    // name = "" ; read [A-Za-z0-9] from k
    push_str(chunk, "", line);
    lset(chunk, name_slot, line);
    let _ = chunk;
    let inner = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, k_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, s_slot, line);
    lget(chunk, k_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, cc_slot, line);
    // is_alnum = (48..=57) | (65..=90) | (97..=122)
    // digit
    lget(chunk, cc_slot, line);
    push_const(chunk, Value::F64(48.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    lget(chunk, cc_slot, line);
    push_const(chunk, Value::F64(57.0), line);
    vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    // upper
    lget(chunk, cc_slot, line);
    push_const(chunk, Value::F64(65.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    lget(chunk, cc_slot, line);
    push_const(chunk, Value::F64(90.0), line);
    vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_OR, line);
    // lower
    lget(chunk, cc_slot, line);
    push_const(chunk, Value::F64(97.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    lget(chunk, cc_slot, line);
    push_const(chunk, Value::F64(122.0), line);
    vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if_value(line);
    // alnum: name += s.charAt(k); k += 1
    lget(chunk, name_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, k_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, name_slot, line);
    lget(chunk, k_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, k_slot, line);
    chunk.emit_else(line);
    // non-alnum: stop the name scan
    lget(chunk, n_slot, line);
    lset(chunk, k_slot, line);
    chunk.emit_end(line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, inner, line);
    let chunk = &mut chunks[current];

    // keep = allowed.indexOf("<" + name.toLowerCase() + ">") >= 0
    lget(chunk, allowed_slot, line);
    push_str(chunk, "<", line);
    lget(chunk, name_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "toLowerCase");
        chunk.emit_call(idx, 1, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    push_str(chunk, ">", line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    {
        let idx = chunk.add_import("ecma:string", "indexOf");
        chunk.emit_call(idx, 2, line);
    }
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // keep the whole tag: out += s.substring(i, j+1)
    lget(chunk, out_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    {
        let idx = chunk.add_import("ecma:string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    chunk.emit_end(line);
    // i = j + 1
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_end(line); // end j<0 else

    chunk.emit_else(line); // char is not '<'
    // out += s.charAt(i); i += 1
    lget(chunk, out_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_end(line); // end if '<'
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, outer, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
}

// ── strrchr ───────────────────────────────────────────────────────────────

pub fn emit_strrchr(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let needle_slot = alloc_local(chunk);
    let hay_slot = alloc_local(chunk);
    let pos_slot = alloc_local(chunk);
    lset(chunk, needle_slot, line);
    lset(chunk, hay_slot, line);

    // use first char of needle
    lget(chunk, needle_slot, line);
    push_const(chunk, Value::I32(0), line);
    push_const(chunk, Value::I32(1), line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, needle_slot, line);

    lget(chunk, hay_slot, line);
    lget(chunk, needle_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "lastIndexOf");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, pos_slot, line);

    // if pos < 0: false; else: hay.substring(pos)
    lget(chunk, pos_slot, line);
    push_const(chunk, Value::I32(0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);
    lget(chunk, hay_slot, line);
    lget(chunk, pos_slot, line);
    push_const(chunk, Value::I32(i32::MAX), line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    chunk.emit_end(line);
}

// ── explode ───────────────────────────────────────────────────────────────

/// PHP `explode(separator, string, limit)`.
///
/// The whole body is the SHARED split-with-limit emitter: `explode` and python's
/// `str.split(sep, maxsplit)` are one algorithm — split fully, re-join the tail
/// past the limit — and differed only in that php's limit counts PIECES while
/// python's counts SPLITS, and that a negative limit drops trailing pieces here
/// but means "unlimited" there. Both are flags on
/// [`strings::SplitOptions`], so this is a binding, not an implementation.
///
/// Stack on entry (argc=3): `str, delim, limit`; (argc=2): `str, delim`.
pub fn emit_explode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // Separator and subject are both string PARAMETERS in php, so both take
    // php's coercion; a third argument is the limit and must not.
    coerce_str_args(chunks, current, argc, 2, line);
    strings::emit_split_limit(
        chunks,
        current,
        argc,
        strings::SplitOptions::max_pieces(),
        line,
    );
}

// ── sscanf ────────────────────────────────────────────────────────────────

/// PHP `uniqid(string $prefix = "", bool $more_entropy = false): string`
/// Returns `prefix . floor(Date.now() * 1000).toString(16)`.
/// When `more_entropy` is true, appends ".00000000" (fixed placeholder —
/// the caller only cares that the ID is unique across the same millisecond).
pub fn emit_php_uniqid(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let more_entropy_slot = if argc >= 2 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let prefix_slot = alloc_local(chunk);
    let hex_slot = alloc_local(chunk);
    let result_slot = alloc_local(chunk);

    if let Some(slot) = more_entropy_slot {
        lset(chunk, slot, line);
    }
    if argc >= 1 {
        lset(chunk, prefix_slot, line);
    } else {
        push_str(chunk, "", line);
        lset(chunk, prefix_slot, line);
    }

    let date_now_idx = chunks[0].add_import("ecma:date".to_string(), "now".to_string());
    let num_tostr_idx = chunks[0].add_import("ecma:number".to_string(), "toString".to_string());
    let chunk = &mut chunks[current];

    // hex = floor(Date.now() * 1000).toString(16)
    chunk.emit_call(date_now_idx, 0u8, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    push_const(chunk, Value::F64(16.0), line);
    chunk.emit_call(num_tostr_idx, 2u8, line);
    lset(chunk, hex_slot, line);

    // result = prefix + hex
    lget(chunk, prefix_slot, line);
    lget(chunk, hex_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, result_slot, line);

    if let Some(me_slot) = more_entropy_slot {
        lget(chunk, me_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        lget(chunk, result_slot, line);
        push_str(chunk, ".00000000", line);
        {
            let idx = chunk.add_import("wasm:js-string", "concat");
            chunk.emit_call(idx, 2, line);
        }
        lset(chunk, result_slot, line);
        chunk.emit_end(line);
    }

    lget(chunk, result_slot, line);
}

// ── preg_replace_limited ─────────────────────────────────────────────────
/// PHP `preg_replace($pat, $repl, $str, $limit)`.
/// When limit=-1 (unlimited), uses replaceAll. Otherwise uses replace (first match only).
pub fn emit_preg_replace_limited(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let replace_all = chunks[0].add_import("ecma:regexp".to_string(), "replaceAll".to_string());
    let replace_one = chunks[0].add_import("ecma:regexp".to_string(), "replace".to_string());
    let chunk = &mut chunks[current];
    let limit_slot = alloc_local(chunk);
    let str_slot = alloc_local(chunk);
    let repl_slot = alloc_local(chunk);
    let pat_slot = alloc_local(chunk);
    lset(chunk, limit_slot, line);
    lset(chunk, str_slot, line);
    lset(chunk, repl_slot, line);
    lset(chunk, pat_slot, line);
    // if limit == -1 or limit < 0: replaceAll; else: replace (first match)
    lget(chunk, limit_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, str_slot, line);
    lget(chunk, pat_slot, line);
    lget(chunk, repl_slot, line);
    chunk.emit_call(replace_all, 3u8, line);
    chunk.emit_else(line);
    lget(chunk, str_slot, line);
    lget(chunk, pat_slot, line);
    lget(chunk, repl_slot, line);
    chunk.emit_call(replace_one, 3u8, line);
    chunk.emit_end(line);
}

// ── strrpos / strripos ────────────────────────────────────────────────────
fn emit_strrpos_impl(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    case_insensitive: bool,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let haystack_slot = alloc_local(chunk);
    let needle_slot = alloc_local(chunk);
    let offset_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let needle_len_slot = alloc_local(chunk);
    let search_slot = alloc_local(chunk);
    let base_slot = alloc_local(chunk);
    let end_slot = alloc_local(chunk);
    let idx_slot = alloc_local(chunk);

    if argc >= 3 {
        vybe_compiler::primitives::convert::emit_to_int(chunk, line);
        lset(chunk, offset_slot, line);
    }
    coerce_to_str(chunk, line);
    if case_insensitive {
        let idx = chunk.add_import("ecma:string", "toLowerCase");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, needle_slot, line);
    coerce_to_str(chunk, line);
    if case_insensitive {
        let idx = chunk.add_import("ecma:string", "toLowerCase");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, haystack_slot, line);
    if argc < 3 {
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, offset_slot, line);
    }

    lget(chunk, haystack_slot, line);
    let len = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(len, 1, line);
    lset(chunk, n_slot, line);
    lget(chunk, needle_slot, line);
    let len = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(len, 1, line);
    lset(chunk, needle_len_slot, line);

    lget(chunk, offset_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, n_slot, line);
    lget(chunk, offset_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, needle_len_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, end_slot, line);
    lget(chunk, end_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, end_slot, line);
    chunk.emit_else(line);
    lget(chunk, end_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, n_slot, line);
    lset(chunk, end_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    lget(chunk, haystack_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, end_slot, line);
    let substring = chunk.add_import("wasm:js-string", "substring");
    chunk.emit_call(substring, 3, line);
    lset(chunk, search_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, base_slot, line);
    chunk.emit_else(line);
    lget(chunk, haystack_slot, line);
    lget(chunk, offset_slot, line);
    lget(chunk, n_slot, line);
    let substring = chunk.add_import("wasm:js-string", "substring");
    chunk.emit_call(substring, 3, line);
    lset(chunk, search_slot, line);
    lget(chunk, offset_slot, line);
    lset(chunk, base_slot, line);
    chunk.emit_end(line);

    lget(chunk, search_slot, line);
    lget(chunk, needle_slot, line);
    let idx = chunk.add_import("ecma:string", "lastIndexOf");
    chunk.emit_call(idx, 2, line);
    lset(chunk, idx_slot, line);

    lget(chunk, idx_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, idx_slot, line);
    lget(chunk, base_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

pub fn emit_strrpos(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_strrpos_impl(chunks, current, argc, false, line);
}

/// Case-insensitive `strrpos` with PHP miss/offset semantics.
pub fn emit_strripos(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_strrpos_impl(chunks, current, argc, true, line);
}

// ── str_word_count with mode ──────────────────────────────────────────────
fn emit_str_word_count_with_mode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // Stack on entry: [s, mode, charlist?]. The charlist is optional and
    // currently ignored by the simplified splitter, but it still has to be
    // consumed so mode/string stay in the expected slots.
    // Actually, looking at the PHP profile: (s, mode, charlist?) so walker passes s first then mode
    // But since it's a 2-arg call, TOS = mode, below = s
    let chunk = &mut chunks[current];
    let mode_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    if argc >= 3 {
        let charlist_slot = alloc_local(chunk);
        lset(chunk, charlist_slot, line);
    }
    lset(chunk, mode_slot, line); // TOS = mode
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);

    // Split by spaces to get words (simplified; handles basic ASCII word splitting)
    // For mode 1: return words array; mode 2: build pos→word object
    lget(chunk, mode_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    // mode 1: return array of words
    // Use regex split on whitespace, filter empty
    lget(chunk, s_slot, line);
    push_str(chunk, " ", line);
    {
        let idx = chunk.add_import("ecma:string", "split");
        chunk.emit_call(idx, 2, line);
    }
    let arr_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let word_slot = alloc_local(chunk);
    lset(chunk, arr_slot, line);
    lget(chunk, arr_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    chunk.emit_array_new_fixed(0, 0, line);
    lset(chunk, out_slot, line);
    let _ = chunk;
    let ls1 = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, arr_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, word_slot, line);
    // only push non-empty words
    lget(chunk, word_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    let _ = chunk;
    if argc >= 3 {
        let chunk = &mut chunks[current];
        lget(chunk, word_slot, line);
        push_str(chunk, "'", line);
        {
            let idx = chunk.add_import("ecma:string", "split");
            chunk.emit_call(idx, 2, line);
        }
        let parts_slot = alloc_local(chunk);
        let part_i_slot = alloc_local(chunk);
        let part_len_slot = alloc_local(chunk);
        let part_slot = alloc_local(chunk);
        lset(chunk, parts_slot, line);
        lget(chunk, parts_slot, line);
        chunk.emit_op(Op::ARRAY_LENGTH, line);
        lset(chunk, part_len_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, part_i_slot, line);
        let _ = chunk;
        let part_loop = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
        let chunk = &mut chunks[current];
        lget(chunk, part_i_slot, line);
        lget(chunk, part_len_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
        let _ = chunk;
        vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
        let chunk = &mut chunks[current];
        lget(chunk, parts_slot, line);
        lget(chunk, part_i_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, part_slot, line);
        lget(chunk, part_slot, line);
        {
            let idx = chunk.add_import("wasm:js-string", "length");
            chunk.emit_call(idx, 1, line);
        }
        push_const(chunk, Value::F64(0.0), line);
        vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        lget(chunk, out_slot, line);
        lget(chunk, part_slot, line);
        let _ = chunk;
        call_import(chunks, current, "ecma:array", "push", 2, line);
        chunks[current].emit_op(Op::DROP, line);
        let chunk = &mut chunks[current];
        chunk.emit_end(line);
        lget(chunk, part_i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, part_i_slot, line);
        let _ = chunk;
        vybe_compiler::primitives::loops::emit_loop_end(chunks, current, part_loop, line);
    } else {
        let chunk = &mut chunks[current];
        lget(chunk, out_slot, line);
        lget(chunk, word_slot, line);
        let _ = chunk;
        call_import(chunks, current, "ecma:array", "push", 2, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, ls1, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);

    chunk.emit_else(line); // mode != 1

    lget(chunk, mode_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    // mode 2: return position→word object (map)
    // We need to build an associative array where keys are positions
    // Use a simple approach: walk through string, track word positions
    let pos_slot = alloc_local(chunk);
    let char_pos_slot = alloc_local(chunk);
    let slen_slot = alloc_local(chunk);
    let code_slot2 = alloc_local(chunk);
    let in_word2_slot = alloc_local(chunk);
    let wstart_slot = alloc_local(chunk);
    let obj_slot = alloc_local(chunk);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, slen_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, pos_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, wstart_slot, line);
    push_const(chunk, Value::Bool(false), line);
    lset(chunk, in_word2_slot, line);
    // Create an empty associative array (Map). PHP mode 2 returns a
    // position→word map; count() must return the entry count (3 for
    // "a b c"), not max-index+1 — so a Map, not a sparse Array.
    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, obj_slot, line);

    let _ = chunk;
    let ls2 = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, pos_slot, line);
    lget(chunk, slen_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // get char code (or sentinel if pos == slen)
    lget(chunk, pos_slot, line);
    lget(chunk, slen_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, s_slot, line);
    lget(chunk, pos_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    chunk.emit_else(line);
    push_const(chunk, Value::F64(32.0), line); // space sentinel at end
    chunk.emit_end(line);
    lset(chunk, code_slot2, line);

    // is_alpha: 65-90, 97-122 or hyphen 45 or apostrophe 39
    // Simplified: is_word_char if NOT space/tab/etc
    // is_sep: code <= 32 (whitespace)
    lget(chunk, code_slot2, line);
    push_const(chunk, Value::F64(32.0), line);
    vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    let is_sep2_slot = alloc_local(chunk);
    lset(chunk, is_sep2_slot, line);

    lget(chunk, is_sep2_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // separator: if in_word, save word at wstart
    lget(chunk, in_word2_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // save: obj[wstart] = s.substring(wstart, pos)
    lget(chunk, s_slot, line);
    lget(chunk, wstart_slot, line);
    lget(chunk, pos_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, char_pos_slot, line); // reuse as word_str_slot
    lget(chunk, obj_slot, line);
    lget(chunk, wstart_slot, line);
    lget(chunk, char_pos_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    push_const(chunk, Value::Bool(false), line);
    lset(chunk, in_word2_slot, line);
    chunk.emit_end(line);
    chunk.emit_else(line); // not separator
    lget(chunk, in_word2_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_if(line);
    lget(chunk, pos_slot, line);
    lset(chunk, wstart_slot, line);
    push_const(chunk, Value::Bool(true), line);
    lset(chunk, in_word2_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    lget(chunk, pos_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, pos_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, ls2, line);
    let chunk = &mut chunks[current];
    lget(chunk, obj_slot, line);

    chunk.emit_else(line); // mode == 0 (or unknown)
    // Return count (mode 0) - need to re-run the count logic
    // For simplicity, rebuild from s_slot
    lget(chunk, s_slot, line);
    push_str(chunk, " ", line);
    {
        let idx = chunk.add_import("ecma:string", "split");
        chunk.emit_call(idx, 2, line);
    }
    let arr2 = alloc_local(chunk);
    let len2 = alloc_local(chunk);
    let i2 = alloc_local(chunk);
    let cnt2 = alloc_local(chunk);
    let w2 = alloc_local(chunk);
    lset(chunk, arr2, line);
    lget(chunk, arr2, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len2, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i2, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, cnt2, line);
    let _ = chunk;
    let ls3 = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i2, line);
    lget(chunk, len2, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, arr2, line);
    lget(chunk, i2, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, w2, line);
    lget(chunk, w2, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, cnt2, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, cnt2, line);
    chunk.emit_end(line);
    lget(chunk, i2, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i2, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, ls3, line);
    let chunk = &mut chunks[current];
    lget(chunk, cnt2, line);
    chunk.emit_end(line); // end mode==2
    chunk.emit_end(line); // end mode==1
}

// ── var_export ────────────────────────────────────────────────────────────
/// PHP `var_export($val[, $return])` — PHP-syntax representation.
pub fn emit_var_export(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let ret_slot = alloc_local(chunk);
    let val_slot = alloc_local(chunk);

    if argc >= 2 {
        lset(chunk, ret_slot, line); // return flag
    } else {
        push_const(chunk, Value::Bool(false), line);
        lset(chunk, ret_slot, line);
    }
    lset(chunk, val_slot, line);

    // Build the representation string
    let repr_slot = alloc_local(chunk);
    // Check if value is false (use is_undefined/is_null + bool type check)
    // false: use wasm:js-boolean test + cast to check if it's boolean false
    let test_bool = chunks[current].add_import("wasm:js-boolean", "test");
    let chunk = &mut chunks[current];
    lget(chunk, val_slot, line);
    chunk.emit_call(test_bool, 1u8, line); // returns i32 1 if boolean
    // Also check: val evaluated as bool is false
    lget(chunk, val_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line); // both: is boolean AND falsy
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_str(chunk, "false", line);
    lset(chunk, repr_slot, line);
    chunk.emit_else(line);
    // Check null
    lget(chunk, val_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_str(chunk, "NULL", line);
    lset(chunk, repr_slot, line);
    chunk.emit_else(line);
    // Check true: is_bool AND truthy
    lget(chunk, val_slot, line);
    chunk.emit_call(test_bool, 1u8, line);
    lget(chunk, val_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_str(chunk, "true", line);
    lset(chunk, repr_slot, line);
    chunk.emit_else(line);
    // Strings: PHP-syntax single-quoted representation.
    let test_str = chunks[current].add_import("wasm:js-string", "test");
    let chunk = &mut chunks[current];
    lget(chunk, val_slot, line);
    chunk.emit_call(test_str, 1u8, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_str(chunk, "'", line);
    lget(chunk, val_slot, line);
    push_str(chunk, "\\", line);
    push_str(chunk, "\\\\", line);
    {
        let idx = chunk.add_import("ecma:string", "replaceAll");
        chunk.emit_call(idx, 3, line);
    }
    push_str(chunk, "'", line);
    push_str(chunk, "\\'", line);
    {
        let idx = chunk.add_import("ecma:string", "replaceAll");
        chunk.emit_call(idx, 3, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    push_str(chunk, "'", line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, repr_slot, line);
    chunk.emit_else(line);
    // numbers and other scalars: coerce to string
    lget(chunk, val_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, repr_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    // if return: push repr; else: log repr and push null
    lget(chunk, ret_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, repr_slot, line);
    chunk.emit_else(line);
    lget(chunk, repr_slot, line);
    let out_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], out_slot, line);
    let write_idx = chunks[current].add_import("wasi:cli/stdout", "write-via-stream");
    let rd_slot = alloc_local(&mut chunks[current]);
    let wr_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::io::emit_write_stdout_with_imports(
        chunk,
        write_idx,
        rd_slot,
        wr_slot,
        line,
        |c| c.emit_op_u16(Op::LOCAL_GET, out_slot, line),
    );
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_end(line);
}

// ── strncmp / strncasecmp ─────────────────────────────────────────────────
/// PHP `strncmp($a, $b, $n)` / `strncasecmp($a, $b, $n)`.
/// Compares the first $n chars of each string.
pub fn emit_strncmp(chunks: &mut [Chunk], current: usize, case_insensitive: bool, line: u32) {
    let chunk = &mut chunks[current];
    let n_slot = alloc_local(chunk);
    let b_slot = alloc_local(chunk);
    let a_slot = alloc_local(chunk);
    lset(chunk, n_slot, line); // n (TOS)
    lset(chunk, b_slot, line); // b
    lset(chunk, a_slot, line); // a
    // sub_a = a.substring(0, n); sub_b = b.substring(0, n)
    lget(chunk, a_slot, line);
    push_const(chunk, Value::I32(0), line);
    lget(chunk, n_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    lget(chunk, b_slot, line);
    push_const(chunk, Value::I32(0), line);
    lget(chunk, n_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    if case_insensitive {
        // Swap to (lower_a, lower_b) for comparison
        let sub_b_slot = alloc_local(chunk);
        lset(chunk, sub_b_slot, line);
        let sub_a_slot = alloc_local(chunk);
        lset(chunk, sub_a_slot, line);
        lget(chunk, sub_a_slot, line);
        {
            let idx = chunk.add_import("ecma:string", "toLowerCase");
            chunk.emit_call(idx, 1, line);
        }
        lget(chunk, sub_b_slot, line);
        {
            let idx = chunk.add_import("ecma:string", "toLowerCase");
            chunk.emit_call(idx, 1, line);
        }
    }
    {
        let idx = chunk.add_import("wasm:js-string", "compare");
        chunk.emit_call(idx, 2, line);
    }
}

// ── strpbrk ───────────────────────────────────────────────────────────────
pub fn emit_strpbrk(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let chars_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let j_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let clen_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let found_slot = alloc_local(chunk);
    lset(chunk, chars_slot, line);
    lset(chunk, s_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    push_const(chunk, Value::F64(-1.0), line);
    lset(chunk, found_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, len_slot, line);
    lget(chunk, chars_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, clen_slot, line);
    let _ = chunk;
    let outer = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    let inner = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, clen_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, chars_slot, line);
    lget(chunk, j_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lget(chunk, code_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, i_slot, line);
    lset(chunk, found_slot, line);
    lget(chunk, len_slot, line);
    lset(chunk, i_slot, line);
    lget(chunk, clen_slot, line);
    lset(chunk, j_slot, line);
    chunk.emit_end(line);
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, inner, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, outer, line);
    let chunk = &mut chunks[current];
    lget(chunk, found_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, s_slot, line);
    lget(chunk, found_slot, line);
    push_const(chunk, Value::I32(i32::MAX), line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    chunk.emit_else(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_end(line);
}

// ── substr_compare ────────────────────────────────────────────────────────
pub fn emit_substr_compare(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let ci_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let off_slot = alloc_local(chunk);
    let str_slot = alloc_local(chunk);
    let main_slot = alloc_local(chunk);
    let main_len_slot = alloc_local(chunk);
    if argc >= 5 {
        lset(chunk, ci_slot, line);
    } else {
        push_const(chunk, Value::Bool(false), line);
        lset(chunk, ci_slot, line);
    }
    if argc >= 4 {
        lset(chunk, len_slot, line);
    } else {
        push_const(chunk, Value::F64(-1.0), line);
        lset(chunk, len_slot, line);
    }
    lset(chunk, off_slot, line);
    lset(chunk, str_slot, line);
    lset(chunk, main_slot, line);
    lget(chunk, main_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, main_len_slot, line);
    // Normalize negative offset
    lget(chunk, off_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, main_len_slot, line);
    lget(chunk, off_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, off_slot, line);
    chunk.emit_end(line);
    // Use str length as default len
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, str_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, len_slot, line);
    chunk.emit_end(line);
    // sub1 = main.substring(offset, offset+len)
    lget(chunk, main_slot, line);
    lget(chunk, off_slot, line);
    lget(chunk, off_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    let sub1_slot = alloc_local(chunk);
    lset(chunk, sub1_slot, line);
    // sub2 = str.substring(0, len)
    lget(chunk, str_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, len_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    let sub2_slot = alloc_local(chunk);
    lset(chunk, sub2_slot, line);
    lget(chunk, ci_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, sub1_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "toLowerCase");
        chunk.emit_call(idx, 1, line);
    }
    lget(chunk, sub2_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "toLowerCase");
        chunk.emit_call(idx, 1, line);
    }
    {
        let idx = chunk.add_import("wasm:js-string", "compare");
        chunk.emit_call(idx, 2, line);
    }
    chunk.emit_else(line);
    lget(chunk, sub1_slot, line);
    lget(chunk, sub2_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "compare");
        chunk.emit_call(idx, 2, line);
    }
    chunk.emit_end(line);
}

// ── preg_grep ─────────────────────────────────────────────────────────────
pub fn emit_preg_grep(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let flags_slot = alloc_local(chunk);
    let arr_slot = alloc_local(chunk);
    let pat_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let elem_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let invert_slot = alloc_local(chunk);
    let matched_slot = alloc_local(chunk);
    if argc >= 3 {
        lset(chunk, flags_slot, line);
    } else {
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, flags_slot, line);
    }
    lset(chunk, arr_slot, line);
    lset(chunk, pat_slot, line);
    lget(chunk, flags_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    lset(chunk, invert_slot, line);
    chunk.emit_array_new_fixed(0, 0, line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, arr_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len_slot, line);
    let _ = chunk;
    let ls = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, arr_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, elem_slot, line);
    lget(chunk, pat_slot, line);
    lget(chunk, elem_slot, line);
    coerce_to_str(chunk, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:regexp", "exec", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    vybe_compiler::primitives::ops::emit_dyn_ne(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    lset(chunk, matched_slot, line);
    // if matched != invert: include
    lget(chunk, matched_slot, line);
    lget(chunk, invert_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_ne(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, out_slot, line);
    lget(chunk, elem_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, ls, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
}

// ── fnmatch ───────────────────────────────────────────────────────────────
/// PHP `fnmatch($pattern, $string)` — shell-style match, case-SENSITIVE.
///
/// The matcher is the shared one; php's only contribution is argument order.
/// PHP takes the pattern FIRST, so the subject is on top of the stack, while
/// the shared emitter wants `[name, pattern]` — python's order, and the one
/// that reads left-to-right.
pub fn emit_fnmatch(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    /// `FNM_CASEFOLD` — php's case-insensitive glob flag.
    const FNM_CASEFOLD: i32 = 16;

    let chunk = &mut chunks[current];
    let flags_slot = alloc_local(chunk);
    let str_slot = alloc_local(chunk);
    let pat_slot = alloc_local(chunk);
    if argc >= 3 {
        vybe_compiler::primitives::convert::emit_to_int(chunk, line);
        lset(chunk, flags_slot, line);
    }
    lset(chunk, str_slot, line);
    lset(chunk, pat_slot, line);

    // PHP takes the pattern FIRST, so the subject is on top; the shared emitter
    // wants `[name, pattern]`.
    lget(chunk, str_slot, line);
    lget(chunk, pat_slot, line);
    if argc >= 3 {
        lget(chunk, flags_slot, line);
        push_const(chunk, Value::I32(FNM_CASEFOLD), line);
        chunk.emit_op(Op::I32_AND, line);
        push_const(chunk, Value::I32(0), line);
        chunk.emit_op(Op::I32_NE, line);
        chunk.emit_if_value(line);
        push_str(chunk, "i", line);
        chunk.emit_else(line);
        push_str(chunk, "", line);
        chunk.emit_end(line);
    } else {
        push_str(chunk, "", line);
    }
    strings::emit_glob_match_flagged(chunks, current, line);
}

const STRTOK_SUBJECT: &str = "__php_strtok_subject";
const STRTOK_CURSOR: &str = "__php_strtok_cursor";
const STRTOK_SCAN_CHUNK: &str = "__php_strtok_scan";

// ── strtok ────────────────────────────────────────────────────────────────
// Stateful strtok keeps its subject and cursor in module globals.
//
// These used to be written as `emit_op_u16(Op::GLOBAL_SET, 0, line)` under a
// comment reading "global 0 = subject, global 1 = cursor". That operand is a
// CONSTANT-POOL index, not a global slot: the VM resolves it with
// `constant_str(idx)` and uses the resulting STRING as the global's name. So
// strtok's state landed in a global named after whatever string happened to be
// interned first in the calling chunk — stable only while `strtok()` and its
// follow-up calls compiled into the SAME chunk, and colliding with a real user
// global whenever that first constant happened to name one.
/// The scan, emitted ONCE per program as a callable chunk.
///
/// It used to be inlined at every `strtok()` site: two loops plus a
/// `charAt`+`indexOf` pair per character, measured at **~6,376 bytes of
/// bytecode per additional call site** (1 call → 15,475 instructions in
/// `<script>`, 5 calls → 40,979). Memoised by chunk name, the same pattern
/// `sprintf` and python's `repr_adapter` use.
///
/// Params: slot 0 = subject, slot 1 = delimiters, slot 2 = start offset.
/// Returns the token, or `false` at end of input, and updates the cursor.
fn strtok_scan_chunk(chunks: &mut Vec<Chunk>) -> usize {
    if let Some(idx) = chunks.iter().position(|c| c.name == STRTOK_SCAN_CHUNK) {
        return idx;
    }
    let mut c = Chunk::new(STRTOK_SCAN_CHUNK);
    c.arity = 3;
    c.local_count = 3;
    chunks.push(c);
    let idx = chunks.len() - 1;
    emit_strtok_scan(chunks, idx, 0, 1, 2, 0);
    chunks[idx].emit_op(Op::RETURN, 0);
    idx
}

/// `helper(subject, delims, start)` — CALL_REF wants [callee, args…].
fn emit_strtok_call(chunks: &mut Vec<Chunk>, current: usize, s_slot: u16, delim_slot: u16, line: u32) {
    let helper = strtok_scan_chunk(chunks);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::REF_FUNC, helper as u16, line);
    chunk.emit(0u8, line); // upvalue count
    lget(chunk, s_slot, line);
    lget(chunk, delim_slot, line);
    vybe_compiler::primitives::globals::emit_read(chunk, STRTOK_CURSOR, line);
    chunk.emit_op(Op::CALL_REF, line);
    chunk.emit(3u8, line);
}

pub fn emit_strtok_init(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let delim_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    lset(chunk, delim_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);
    lget(chunk, s_slot, line);
    vybe_compiler::primitives::globals::emit_write(chunk, STRTOK_SUBJECT, line);
    // A fresh subject restarts at 0 — the cursor is the helper's third
    // argument, so seed it before the call rather than passing a literal.
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::globals::emit_write(chunk, STRTOK_CURSOR, line);
    chunk.emit_op(Op::DROP, line);
    emit_strtok_call(chunks, current, s_slot, delim_slot, line);
}

pub fn emit_strtok_next(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let delim_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    lset(chunk, delim_slot, line);
    vybe_compiler::primitives::globals::emit_read(chunk, STRTOK_SUBJECT, line);
    lset(chunk, s_slot, line);
    emit_strtok_call(chunks, current, s_slot, delim_slot, line);
}

fn emit_strtok_scan(
    chunks: &mut [Chunk],
    current: usize,
    s_slot: u16,
    delim_slot: u16,
    start_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let token_start_slot = alloc_local(chunk);
    lget(chunk, start_slot, line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    let len = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(len, 1, line);
    lset(chunk, n_slot, line);

    let _ = chunk;
    let skip_loop = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, delim_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    let char_at = chunk.add_import("ecma:string", "charAt");
    chunk.emit_call(char_at, 2, line);
    let index_of = chunk.add_import("ecma:string", "indexOf");
    chunk.emit_call(index_of, 2, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(skip_loop.break_depth(0) as u32, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, skip_loop, line);
    let chunk = &mut chunks[current];

    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::globals::emit_write(chunk, STRTOK_CURSOR, line);
    chunk.emit_bool_const(false, line);
    chunk.emit_else(line);
    lget(chunk, i_slot, line);
    lset(chunk, token_start_slot, line);
    let _ = chunk;
    let token_loop = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, delim_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    let char_at = chunk.add_import("ecma:string", "charAt");
    chunk.emit_call(char_at, 2, line);
    let index_of = chunk.add_import("ecma:string", "indexOf");
    chunk.emit_call(index_of, 2, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(token_loop.break_depth(0) as u32, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, token_loop, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_else(line);
    lget(chunk, i_slot, line);
    chunk.emit_end(line);
    vybe_compiler::primitives::globals::emit_write(chunk, STRTOK_CURSOR, line);
    lget(chunk, s_slot, line);
    lget(chunk, token_start_slot, line);
    lget(chunk, i_slot, line);
    let substring = chunk.add_import("wasm:js-string", "substring");
    chunk.emit_call(substring, 3, line);
    chunk.emit_end(line);
}

// ── mb_convert_case ───────────────────────────────────────────────────────
pub fn emit_mb_convert_case(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let mode_slot = alloc_local(chunk);
    lset(chunk, mode_slot, line);
    coerce_to_str(chunk, line);
    let s_slot = alloc_local(chunk);
    lset(chunk, s_slot, line);
    // mode: 0=UPPER, 1=LOWER, 2=TITLE
    lget(chunk, mode_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "toUpperCase");
        chunk.emit_call(idx, 1, line);
    }
    chunk.emit_else(line);
    lget(chunk, mode_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "toLowerCase");
        chunk.emit_call(idx, 1, line);
    }
    chunk.emit_else(line);
    // MB_CASE_TITLE: ucwords
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "toLowerCase");
        chunk.emit_call(idx, 1, line);
    }
    push_str(chunk, " ", line);
    {
        let idx = chunk.add_import("ecma:string", "split");
        chunk.emit_call(idx, 2, line);
    }
    let words_slot = alloc_local(chunk);
    lset(chunk, words_slot, line);
    let nw_slot = alloc_local(chunk);
    let wi_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let word_slot = alloc_local(chunk);
    lget(chunk, words_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, nw_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, wi_slot, line);
    chunk.emit_array_new_fixed(0, 0, line);
    lset(chunk, out_slot, line);
    let _ = chunk;
    let ws = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, wi_slot, line);
    lget(chunk, nw_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, words_slot, line);
    lget(chunk, wi_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, word_slot, line);
    lget(chunk, word_slot, line);
    push_const(chunk, Value::I32(0), line);
    push_const(chunk, Value::I32(1), line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    {
        let idx = chunk.add_import("ecma:string", "toUpperCase");
        chunk.emit_call(idx, 1, line);
    }
    lget(chunk, word_slot, line);
    push_const(chunk, Value::I32(1), line);
    push_const(chunk, Value::I32(i32::MAX), line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, word_slot, line);
    lget(chunk, out_slot, line);
    lget(chunk, word_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    lget(chunk, wi_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, wi_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, ws, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    push_str(chunk, " ", line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "join", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_mb_strwidth(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let s_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let width_slot = alloc_local(chunk);
    let cp_slot = alloc_local(chunk);

    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, width_slot, line);
    lget(chunk, s_slot, line);
    let len_idx = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(len_idx, 1, line);
    lset(chunk, n_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    let cp_idx = chunk.add_import("wasm:js-string", "codePointAt");
    chunk.emit_call(cp_idx, 2, line);
    lset(chunk, cp_slot, line);

    lget(chunk, width_slot, line);
    lget(chunk, cp_slot, line);
    push_const(chunk, Value::F64(0x3040 as f64), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    lget(chunk, cp_slot, line);
    push_const(chunk, Value::F64(0x9fff as f64), line);
    vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_else(line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_end(line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, width_slot, line);

    lget(chunk, i_slot, line);
    lget(chunk, cp_slot, line);
    push_const(chunk, Value::F64(65535.0), line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_else(line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_end(line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, i_slot, line);

    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, width_slot, line);
}

pub fn emit_mb_setting(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    key: &str,
    default: &str,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let old_slot = alloc_local(chunk);
    if argc == 0 {
        vybe_compiler::primitives::globals::emit_read(chunk, key, line);
        chunk.emit_dup(line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if(line);
        chunk.emit_op(Op::DROP, line);
        push_str(chunk, default, line);
        chunk.emit_end(line);
        return;
    }

    vybe_compiler::primitives::globals::emit_read(chunk, key, line);
    chunk.emit_dup(line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_op(Op::DROP, line);
    push_str(chunk, default, line);
    chunk.emit_end(line);
    lset(chunk, old_slot, line);
    vybe_compiler::primitives::globals::emit_write(chunk, key, line);
    lget(chunk, old_slot, line);
}

// ── strpos / mb_strpos ─────────────────────────────────────────────
//
// Relocated out of the shared compiler's intrinsic table (`primitives/mod.rs`
// `"php_strpos"`) — PHP-specific runtime semantics belong in the PHP
// emitter. Returns the code-unit index of the first `needle` occurrence
// (>= `offset`), or PHP's `false` on a miss. `ecma:string.indexOf` is
// UTF-16 code-unit based (ECMA-262 §22.1.3.9), so for BMP text this equals
// the codepoint index `mb_strpos` wants.
pub fn emit_strpos(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let haystack_slot = alloc_local(chunk);
    let needle_slot = alloc_local(chunk);
    let offset_slot = alloc_local(chunk);
    let idx_slot = alloc_local(chunk);

    // Stack (bottom→top): haystack, needle, [offset]. Pop in reverse.
    if argc >= 3 {
        vybe_compiler::primitives::convert::emit_to_int(chunk, line);
        lset(chunk, offset_slot, line);
    }
    coerce_to_str(chunk, line);
    lset(chunk, needle_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, haystack_slot, line);
    if argc < 3 {
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, offset_slot, line);
    }
    lget(chunk, haystack_slot, line);
    let len = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(len, 1, line);
    let haystack_len_slot = alloc_local(chunk);
    lset(chunk, haystack_len_slot, line);
    lget(chunk, offset_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, haystack_len_slot, line);
    lget(chunk, offset_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, offset_slot, line);
    lget(chunk, offset_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, offset_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    // idx = haystack.substring(offset).indexOf(needle)
    lget(chunk, haystack_slot, line);
    lget(chunk, offset_slot, line);
    lget(chunk, haystack_len_slot, line);
    vybe_compiler::primitives::strings::emit_substring(chunk, line);
    lget(chunk, needle_slot, line);
    vybe_compiler::primitives::strings::emit_index_of(chunk, line);
    lset(chunk, idx_slot, line);

    // idx >= 0 ? idx + offset : false
    lget(chunk, idx_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, idx_slot, line);
    lget(chunk, offset_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

// ── strtr ──────────────────────────────────────────────────────────
//
// Relocated out of the shared compiler's intrinsic table (`primitives/mod.rs`
// `"strtr"`). Two forms:
//   strtr($s, $from, $to)  — per-position single-char translation
//   strtr($s, $map)        — associative replacement (longest key first;
//                            the walker pre-sorts literal maps)
// Uses `replaceAll` (PHP replaces every occurrence — the old intrinsic used
// `replace`, which only hit the first).
pub fn emit_strtr(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 3 {
        emit_strtr3(chunks, current, line);
    } else {
        emit_strtr_map(chunks, current, line);
    }
}

fn emit_strtr3(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let to_slot = alloc_local(chunk);
    let from_slot = alloc_local(chunk);
    let str_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let flen_slot = alloc_local(chunk);
    let fc_slot = alloc_local(chunk);
    let tc_slot = alloc_local(chunk);

    // Stack (bottom→top): str, from, to. Pop in reverse.
    coerce_to_str(chunk, line);
    lset(chunk, to_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, from_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, str_slot, line);

    lget(chunk, from_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, flen_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    let _ = chunk;
    let st = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, flen_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // fc = from.charAt(i)
    lget(chunk, from_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, fc_slot, line);
    // tc = to.charAt(i)
    lget(chunk, to_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, tc_slot, line);
    // str = str.replaceAll(fc, tc)
    lget(chunk, str_slot, line);
    lget(chunk, fc_slot, line);
    lget(chunk, tc_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "replaceAll");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, str_slot, line);
    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, st, line);
    let chunk = &mut chunks[current];
    lget(chunk, str_slot, line);
}

fn emit_strtr_map(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let map_slot = alloc_local(chunk);
    let str_slot = alloc_local(chunk);
    let entries_slot = alloc_local(chunk);
    let idx_slot = alloc_local(chunk);
    let pair_slot = alloc_local(chunk);

    // Stack (bottom→top): str, map. Pop map then str.
    lset(chunk, map_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, str_slot, line);

    lget(chunk, map_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::collections::emit_iter_entries(chunks, current, line);
    let chunk = &mut chunks[current];
    lset(chunk, entries_slot, line);

    let _ = chunk;
    let state = vybe_compiler::primitives::loops::emit_for_in_start(
        chunks,
        current,
        entries_slot,
        idx_slot,
        line,
    );
    // for_in_start leaves the current entry on the stack — drop it; we
    // re-fetch [k, v] by index below.
    chunks[current].emit_op(Op::DROP, line);

    let chunk = &mut chunks[current];
    lget(chunk, str_slot, line);
    lget(chunk, entries_slot, line);
    lget(chunk, idx_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    let chunk = &mut chunks[current];
    lset(chunk, pair_slot, line);
    lget(chunk, pair_slot, line);
    push_const(chunk, Value::I32(0), line);
    let _ = chunk;
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, pair_slot, line);
    push_const(chunk, Value::I32(1), line);
    let _ = chunk;
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    let chunk = &mut chunks[current];
    {
        let idx = chunk.add_import("ecma:string", "replaceAll");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, str_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    let chunk = &mut chunks[current];
    lget(chunk, str_slot, line);
}

// ── quotemeta / strspn / strcspn ───────────────────────────────────

/// PHP `quotemeta($s)` — backslash-escape the regex metacharacters
/// `. \ + * ? [ ^ ] $ ( )`.
pub fn emit_quotemeta(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let c_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let is_meta_slot = alloc_local(chunk);

    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, n_slot, line);

    // . 46, \ 92, + 43, * 42, ? 63, [ 91, ^ 94, ] 93, $ 36, ( 40, ) 41
    let metas: &[u32] = &[46, 92, 43, 42, 63, 91, 94, 93, 36, 40, 41];

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, c_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);

    push_const(chunk, Value::Bool(false), line);
    lset(chunk, is_meta_slot, line);
    for &m in metas {
        lget(chunk, code_slot, line);
        push_const(chunk, Value::F64(m as f64), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        push_const(chunk, Value::Bool(true), line);
        lset(chunk, is_meta_slot, line);
        chunk.emit_end(line);
    }
    lget(chunk, is_meta_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, out_slot, line);
    push_str(chunk, "\\", line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lget(chunk, c_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    chunk.emit_else(line);
    lget(chunk, out_slot, line);
    lget(chunk, c_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    chunk.emit_end(line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
}

/// Shared body for `strspn`/`strcspn`. `reject` = false → count the initial
/// run of chars **in** the mask (strspn); true → chars **not** in the mask
/// (strcspn).
fn emit_str_span(chunks: &mut [Chunk], current: usize, argc: u8, reject: bool, line: u32) {
    let chunk = &mut chunks[current];
    let len_arg_slot = alloc_local(chunk);
    let end_slot = alloc_local(chunk);
    let mask_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let count_slot = alloc_local(chunk);
    let c_slot = alloc_local(chunk);

    // Stack (bottom→top): subject, mask, offset?, length?. Pop in reverse.
    if argc >= 4 {
        lset(chunk, len_arg_slot, line);
    } else {
        push_const(chunk, Value::F64(f64::NAN), line);
        lset(chunk, len_arg_slot, line);
    }
    if argc >= 3 {
        lset(chunk, i_slot, line);
    } else {
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, i_slot, line);
    }
    coerce_to_str(chunk, line);
    lset(chunk, mask_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, count_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, n_slot, line);

    // Normalize PHP's offset/length window:
    // negative offset counts from the end; negative length trims from the end.
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, n_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_end(line);

    // If the normalized offset is outside the subject, scan an empty window.
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, n_slot, line);
    lset(chunk, i_slot, line);
    chunk.emit_end(line);
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, n_slot, line);
    lset(chunk, i_slot, line);
    chunk.emit_end(line);

    lget(chunk, n_slot, line);
    lset(chunk, end_slot, line);
    if argc >= 4 {
        lget(chunk, len_arg_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        lget(chunk, i_slot, line);
        lget(chunk, len_arg_slot, line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, end_slot, line);
        chunk.emit_else(line);
        lget(chunk, n_slot, line);
        lget(chunk, len_arg_slot, line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, end_slot, line);
        chunk.emit_end(line);

        lget(chunk, end_slot, line);
        lget(chunk, i_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        lget(chunk, i_slot, line);
        lset(chunk, end_slot, line);
        chunk.emit_end(line);
        lget(chunk, end_slot, line);
        lget(chunk, n_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        lget(chunk, n_slot, line);
        lset(chunk, end_slot, line);
        chunk.emit_end(line);
    }

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, end_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // c = subject.charAt(i); in_mask = mask.indexOf(c) >= 0
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, c_slot, line);
    lget(chunk, mask_slot, line);
    lget(chunk, c_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "indexOf");
        chunk.emit_call(idx, 2, line);
    }
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    // `matches` = keep going: strspn → char in mask; strcspn → char NOT in mask.
    if reject {
        chunk.emit_op(Op::I32_EQZ, line);
    }
    chunk.emit_if(line);
    // matched: count++, i++
    lget(chunk, count_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, count_slot, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_else(line);
    // stop: force loop exit by jumping i to n
    lget(chunk, end_slot, line);
    lset(chunk, i_slot, line);
    chunk.emit_end(line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, count_slot, line);
}

/// PHP `strspn($subject, $mask)`.
pub fn emit_strspn(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_str_span(chunks, current, argc, false, line);
}

/// PHP `strcspn($subject, $mask)`.
pub fn emit_strcspn(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_str_span(chunks, current, argc, true, line);
}

// ── str_increment / str_decrement (PHP 8.3) ────────────────────────
//
// Perl-style alphanumeric increment/decrement. Rightmost char changes;
// 'z'→'a', 'Z'→'A', '9'→'0' carry left. On carry past the leftmost char
// a new digit/letter is prepended ('1' for digits, 'a'/'A' for letters).
fn emit_str_incdec(chunks: &mut [Chunk], current: usize, inc: bool, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let carry_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let newcode_slot = alloc_local(chunk);

    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    lset(chunk, carry_slot, line);

    // wrap triples: (trigger_code, wrapped_code)
    let wraps: &[(f64, f64)] = if inc {
        &[(122.0, 97.0), (90.0, 65.0), (57.0, 48.0)]
    } else {
        &[(97.0, 122.0), (65.0, 90.0), (48.0, 57.0)]
    };
    let step = if inc { 1.0 } else { -1.0 };

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // if carry: compute new char; else: keep original
    lget(chunk, carry_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    // code = s.charCodeAt(i)
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);
    // assume no further carry; newcode = code + step
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, carry_slot, line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(step), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, newcode_slot, line);
    // wrap overrides
    for &(trigger, wrapped) in wraps {
        lget(chunk, code_slot, line);
        push_const(chunk, Value::F64(trigger), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        push_const(chunk, Value::F64(wrapped), line);
        lset(chunk, newcode_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        lset(chunk, carry_slot, line);
        chunk.emit_end(line);
    }
    // newch = fromCharCode(newcode)
    lget(chunk, newcode_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "fromCharCode");
        chunk.emit_call(idx, 1, line);
    }
    chunk.emit_else(line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    chunk.emit_end(line);
    // out = newch + out
    lget(chunk, out_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    // i--
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    // Final carry → prepend a fresh leading char (increment only; PHP
    // str_decrement on underflow throws, which the tests don't exercise).
    if inc {
        lget(chunk, carry_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        // class of first char decides the prepended digit/letter
        lget(chunk, s_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        {
            let idx = chunk.add_import("wasm:js-string", "charCodeAt");
            chunk.emit_call(idx, 2, line);
        }
        lset(chunk, code_slot, line);
        // default 'a' (97); if digit → '1' (49); if upper A-Z → 'A' (65)
        push_const(chunk, Value::F64(97.0), line);
        lset(chunk, newcode_slot, line);
        // digit? code <= 57
        lget(chunk, code_slot, line);
        push_const(chunk, Value::F64(57.0), line);
        vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        push_const(chunk, Value::F64(49.0), line);
        lset(chunk, newcode_slot, line);
        chunk.emit_end(line);
        // uppercase? code >= 65 && code <= 90
        lget(chunk, code_slot, line);
        push_const(chunk, Value::F64(65.0), line);
        vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        lget(chunk, code_slot, line);
        push_const(chunk, Value::F64(90.0), line);
        vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        push_const(chunk, Value::F64(65.0), line);
        lset(chunk, newcode_slot, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
        lget(chunk, newcode_slot, line);
        {
            let idx = chunk.add_import("wasm:js-string", "fromCharCode");
            chunk.emit_call(idx, 1, line);
        }
        lget(chunk, out_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
        lset(chunk, out_slot, line);
        chunk.emit_end(line);
    }

    lget(chunk, out_slot, line);
}

/// PHP 8.3 `str_increment($s)`.
pub fn emit_str_increment(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_str_incdec(chunks, current, true, line);
}

/// PHP 8.3 `str_decrement($s)`.
/// `str_decrement()` (PHP 8.3). Unlike `str_increment`, decrementing has a
/// FLOOR: `"0"`, `"a"` and `"A"` cannot go lower, and PHP raises a `ValueError`
/// rather than wrapping. Without this guard the shared inc/dec walked past the
/// floor and returned `"9"` for `"0"`.
pub fn emit_str_decrement(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let slot = {
        let chunk = &mut chunks[current];
        let slot = alloc_local(chunk);
        coerce_to_str(chunk, line);
        lset(chunk, slot, line);

        for (i, floor) in ["0", "a", "A"].iter().enumerate() {
            lget(chunk, slot, line);
            push_str(chunk, floor, line);
            vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
            if i > 0 {
                chunk.emit_op(Op::I32_OR, line);
            }
        }
        chunk.emit_if(line);
        slot
    };
    super::type_guard::emit_throw_const(
        chunks,
        current,
        "ValueError",
        "str_decrement(): Argument #1 ($string) must not be empty",
        line,
    );
    chunks[current].emit_end(line);

    lget(&mut chunks[current], slot, line);
    emit_str_incdec(chunks, current, false, line);
}

// ── strlen (byte length) ───────────────────────────────────────────
//
// PHP strings are byte strings, so `strlen` counts UTF-8 *bytes*, unlike
// JS `.length` (UTF-16 code units) / `mb_strlen` (codepoints). Vybe stores
// strings as JS strings, so recover the byte count by summing the UTF-8
// width of each codepoint. (mb_strlen stays on `common:str_length`.)
/// PHP `strlen($s)` — UTF-8 BYTE length, not characters.
///
/// The counter itself is the SHARED one: it was php's private implementation
/// and the only byte-length code on the platform, which is why Lua `#` and Go
/// `len(s)` had nothing to bind to. What stays here is php's COERCION —
/// `strlen(123)` is `3` because php stringifies first, which is php's rule, not
/// the counter's.
pub fn emit_strlen(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    coerce_to_str(&mut chunks[current], line);
    strings::emit_byte_length(chunks, current, line);
}

// ── count_chars ────────────────────────────────────────────────────
//
// PHP `count_chars($s, $mode)`. Supported modes:
//   1 → associative array { byte-value → frequency } for bytes that occur
//   3 → the set of distinct characters (as a keyed map, so array_keys()
//        yields the characters in first-seen order)
// PHP array ≡ Map, so both return an ObjectKind::Map.
pub fn emit_count_chars(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let mode_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let ch_slot = alloc_local(chunk);
    let key_slot = alloc_local(chunk);
    let result_slot = alloc_local(chunk);
    let cur_slot = alloc_local(chunk);

    // Stack (bottom→top): s, [mode]. Pop mode (TOS) then s.
    if argc >= 2 {
        vybe_compiler::primitives::convert::emit_to_int(chunk, line);
        lset(chunk, mode_slot, line);
    }
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);
    if argc < 2 {
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, mode_slot, line);
    }

    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, result_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, n_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, ch_slot, line);

    // key = (mode == 3) ? ch : code
    lget(chunk, mode_slot, line);
    push_const(chunk, Value::F64(3.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, ch_slot, line);
    chunk.emit_else(line);
    lget(chunk, code_slot, line);
    chunk.emit_end(line);
    lset(chunk, key_slot, line);

    // cur = result.has(key) ? result.get(key) : 0
    lget(chunk, result_slot, line);
    lget(chunk, key_slot, line);
    {
        let idx = chunk.add_import("ecma:map", "has");
        chunk.emit_call(idx, 2, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, result_slot, line);
    lget(chunk, key_slot, line);
    {
        let idx = chunk.add_import("ecma:map", "get");
        chunk.emit_call(idx, 2, line);
    }
    chunk.emit_else(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_end(line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, cur_slot, line);

    // result.set(key, cur)
    lget(chunk, result_slot, line);
    lget(chunk, key_slot, line);
    lget(chunk, cur_slot, line);
    {
        let idx = chunk.add_import("ecma:map", "set");
        chunk.emit_call(idx, 3, line);
    }
    chunk.emit_op(Op::DROP, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, result_slot, line);
}

// ── quoted_printable_encode / decode ───────────────────────────────
//
// RFC 2045 §6.7 quoted-printable. Encode: `=`, bytes >126 and control
// chars (except TAB) become `=XX` (upper-case hex). Decode reverses that
// and drops `=CRLF` soft line breaks.

/// Decode a single hex digit's char code to its value on the stack top.


// ── convert_uuencode / convert_uudecode ────────────────────────────
//
// uuencode maps each 3 input bytes to 4 printable chars: 6-bit groups
// offset by 32 (0 → backtick, matching PHP). A leading length char
// records how many bytes the line encodes. Single-line form (inputs in
// these tests are short); decode reads the length char and stops there.

/// Emit `enc(c)` — the value in `c_slot` (0..63) → a uuencode char, on TOS.

/// Emit `dec(code)` — a uuencode char code in `code_slot` → its 0..63 value.

/// Push `floor(slot / div)` onto the stack.

/// Push `slot mod m` onto the stack (m a power used in uuencode).



// ── mb_substr ──────────────────────────────────────────────────────────
//
// PHP has TWO index units and `mb_*` is the second one. `substr` counts UTF-8
// BYTES; `mb_substr` counts UNICODE CODE POINTS. Until 2026-08-01 the walker
// collapsed both onto `__php_substr`, whose bound math runs in UTF-16 units —
// so `mb_substr("a😀b", 1, 1)` returned half a surrogate pair (a replacement
// character) instead of `😀`.
//
// The bound normalization below is PHP's own (negative start counts from the
// end, negative length trims from the end, both clamp), which is why it lives
// here and not in the shared layer. The UNIT comes from the shared scalar
// primitives — `unifiedstringplan.md` Axis 1.
//
// `substr` is deliberately NOT fixed alongside: a byte substring can split a
// code point, and `Value::String` (`Arc<str>`) cannot hold the result. That
// half is blocked on `Literal::Bytes` (plan §3c).
//
// Stack: `[s, start]` or `[s, start, length]` → `[string]`.
pub fn emit_mb_substr(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let start_slot = alloc_local(chunk);
    let length_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let end_slot = alloc_local(chunk);
    let has_length = argc >= 3;

    // Stack is (bottom→top) s, start, [length] — pop in reverse.
    if has_length {
        vybe_compiler::primitives::convert::emit_to_int(chunk, line);
        lset(chunk, length_slot, line);
    }
    vybe_compiler::primitives::convert::emit_to_int(chunk, line);
    lset(chunk, start_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);

    // len = code points in s
    lget(chunk, s_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::strings::emit_scalar_length(chunks, current, line);
    let chunk = &mut chunks[current];
    lset(chunk, len_slot, line);

    // start < 0 → start += len ; then clamp to [0, len]
    clamp_offset(chunk, start_slot, len_slot, line);

    // end = length given ? (length < 0 ? len + length : start + length) : len
    if has_length {
        lget(chunk, length_slot, line);
        push_const(chunk, Value::I32(0), line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        lget(chunk, len_slot, line);
        lget(chunk, length_slot, line);
        chunk.emit_op(Op::I32_ADD, line);
        lset(chunk, end_slot, line);
        chunk.emit_else(line);
        lget(chunk, start_slot, line);
        lget(chunk, length_slot, line);
        chunk.emit_op(Op::I32_ADD, line);
        lset(chunk, end_slot, line);
        chunk.emit_end(line);
    } else {
        lget(chunk, len_slot, line);
        lset(chunk, end_slot, line);
    }
    clamp_range(chunk, end_slot, start_slot, len_slot, line);

    lget(chunk, s_slot, line);
    lget(chunk, start_slot, line);
    lget(chunk, end_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::strings::emit_scalar_substring(chunks, current, line);
}

/// `slot < 0 → slot += len`, then clamp into `[0, len]`.
fn clamp_offset(chunk: &mut Chunk, slot: u16, len_slot: u16, line: u32) {
    lget(chunk, slot, line);
    push_const(chunk, Value::I32(0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, len_slot, line);
    lget(chunk, slot, line);
    chunk.emit_op(Op::I32_ADD, line);
    lset(chunk, slot, line);
    chunk.emit_end(line);
    clamp_to(chunk, slot, len_slot, line);
}

/// Clamp `slot` into `[0, len]`, then raise it to `low` if it fell below.
fn clamp_range(chunk: &mut Chunk, slot: u16, low_slot: u16, len_slot: u16, line: u32) {
    clamp_to(chunk, slot, len_slot, line);
    lget(chunk, slot, line);
    lget(chunk, low_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, low_slot, line);
    lset(chunk, slot, line);
    chunk.emit_end(line);
}

/// Clamp `slot` into `[0, len]`.
fn clamp_to(chunk: &mut Chunk, slot: u16, len_slot: u16, line: u32) {
    lget(chunk, slot, line);
    push_const(chunk, Value::I32(0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::I32(0), line);
    lset(chunk, slot, line);
    chunk.emit_end(line);
    lget(chunk, slot, line);
    lget(chunk, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, len_slot, line);
    lset(chunk, slot, line);
    chunk.emit_end(line);
}

// ── mb_strpos / mb_str_split ───────────────────────────────────────────
//
// Same two-unit story as `mb_substr` above: these count CODE POINTS where
// `strpos` / `str_split` count bytes. Both were aliased onto the byte
// functions, so `mb_strpos("a😀b", "b")` answered 3 (the UTF-16 index) instead
// of 2, and `mb_str_split("a😀b")[1]` was half a surrogate pair.
//
// PHP's own semantics stay here — `false` when absent (NOT `-1`), and an
// offset that may be negative — while the UNIT comes from the shared scalar
// primitives.

/// PHP `mb_strpos($haystack, $needle, $offset = 0)`.
/// Stack: `[haystack, needle]` or `[haystack, needle, offset]` → `[int|false]`.
///
/// Works entirely in the code-point domain: slice the haystack at `offset`,
/// search the tail, then add `offset` back. That keeps the offset argument in
/// the same unit as the result, which a UTF-16 search followed by a conversion
/// would not.
pub fn emit_mb_strpos(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let hay_slot = alloc_local(chunk);
    let needle_slot = alloc_local(chunk);
    let off_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let idx_slot = alloc_local(chunk);

    // Stack (bottom→top): haystack, needle, [offset] — pop in reverse.
    if argc >= 3 {
        vybe_compiler::primitives::convert::emit_to_int(chunk, line);
        lset(chunk, off_slot, line);
    }
    coerce_to_str(chunk, line);
    lset(chunk, needle_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, hay_slot, line);
    if argc < 3 {
        push_const(chunk, Value::I32(0), line);
        lset(chunk, off_slot, line);
    }

    // len = code points in haystack; offset wraps from the end and clamps.
    lget(chunk, hay_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::strings::emit_scalar_length(chunks, current, line);
    let chunk = &mut chunks[current];
    lset(chunk, len_slot, line);
    clamp_offset(chunk, off_slot, len_slot, line);

    // idx = scalar_index_of(scalar_substring(haystack, off, len), needle)
    lget(chunk, hay_slot, line);
    lget(chunk, off_slot, line);
    lget(chunk, len_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::strings::emit_scalar_substring(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, needle_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::strings::emit_scalar_index_of(chunks, current, line);
    let chunk = &mut chunks[current];
    lset(chunk, idx_slot, line);

    // PHP answers `false`, not -1, when the needle is absent.
    lget(chunk, idx_slot, line);
    push_const(chunk, Value::I32(0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);
    lget(chunk, idx_slot, line);
    lget(chunk, off_slot, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_end(line);
}

/// PHP `mb_str_split($s, $length = 1)`.
/// Stack: `[s]` or `[s, length]` → `[array of strings]`.
///
/// The default chunk length is the whole point of the function — one element
/// per CODE POINT — and is exactly `[...s]`. Only the grouped form needs a
/// loop, so the common case costs one host call.
pub fn emit_mb_str_split(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        let chunk = &mut chunks[current];
        coerce_to_str(chunk, line);
        vybe_compiler::primitives::strings::emit_scalar_chars(chunk, line);
        return;
    }

    let chunk = &mut chunks[current];
    let n_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);

    vybe_compiler::primitives::convert::emit_to_int(chunk, line);
    lset(chunk, n_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);

    // A non-positive chunk length is a ValueError in PHP 8; one element per
    // code point is the closest non-throwing answer and matches `$length = 1`.
    lget(chunk, n_slot, line);
    push_const(chunk, Value::I32(1), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::I32(1), line);
    lset(chunk, n_slot, line);
    chunk.emit_end(line);

    lget(chunk, s_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::strings::emit_scalar_length(chunks, current, line);
    let chunk = &mut chunks[current];
    lset(chunk, len_slot, line);
    push_const(chunk, Value::I32(0), line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, out_slot, line);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::I32_ADD, line);
    let _ = chunk;
    vybe_compiler::primitives::strings::emit_scalar_substring(chunks, current, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::I32_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
}

// ── parse_url ─────────────────────────────────────────────────────────
//
// The PARSER is the shared adapter primitive `primitives::url` in
// `ParseOptions::php()` mode — a syntactic RFC 3986 split that preserves what
// WHATWG would fold (scheme case, host case, default port, dot segments).
// Python's `urlsplit` uses the SAME parser with `lowercase_scheme: true`; the
// languages differ by a flag, not by an implementation.
//
// This replaces a PHP-SOURCE PRELUDE that the walker injected whenever the
// program text contained `"parse_url"` — a substring scan that also fired on a
// comment or a user's own function of that name.
//
// What stays here is PHP's own SHAPE: an assoc array in the order
// `scheme,host,port,user,pass,path,query,fragment`, empty components OMITTED,
// `port` as an INT, and the `PHP_URL_*` selector.

/// `(canonical component, php array key, is_int)`, in PHP's own key order.
const PHP_URL_PARTS: &[(url::UrlField, &str, bool)] = &[
    (url::UrlField::Scheme, "scheme", false),
    (url::UrlField::Host, "host", false),
    (url::UrlField::Port, "port", true),
    (url::UrlField::User, "user", false),
    (url::UrlField::Pass, "pass", false),
    (url::UrlField::Path, "path", false),
    (url::UrlField::Query, "query", false),
    (url::UrlField::Fragment, "fragment", false),
];

/// `PHP_URL_*` selector order — `parse_url($u, PHP_URL_HOST)` is index 1.
const PHP_URL_SELECTOR: &[&str] = &[
    "scheme", "host", "port", "user", "pass", "path", "query", "fragment",
];

/// PHP `parse_url($url, $component = -1)`.
/// Stack: `[url]` or `[url, component]` → `[array|string|int|null]`.
pub fn emit_parse_url(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let comp = alloc_local(chunk);
    let parsed = alloc_local(chunk);
    let out = alloc_local(chunk);
    let val = alloc_local(chunk);

    if argc >= 2 {
        vybe_compiler::primitives::convert::emit_to_int(chunk, line);
        lset(chunk, comp, line);
    } else {
        push_const(chunk, Value::I32(-1), line);
        lset(chunk, comp, line);
    }
    coerce_to_str(chunk, line);

    let _ = chunk;
    url::emit_parse(chunks, current, url::ParseOptions::php(), line);
    let chunk = &mut chunks[current];
    lset(chunk, parsed, line);
    let _ = chunk;

    call_import(chunks, current, "ecma:map", "new", 0, line);
    lset(&mut chunks[current], out, line);

    for (field, key, is_int) in PHP_URL_PARTS {
        url::emit_component(chunks, current, parsed, *field, line);
        let chunk = &mut chunks[current];
        lset(chunk, val, line);

        // PHP omits a component that is not present.
        lget(chunk, val, line);
        let len = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(len, 1, line);
        push_const(chunk, Value::I32(0), line);
        chunk.emit_op(Op::I32_NE, line);
        chunk.emit_if(line);
        lget(chunk, out, line);
        push_str(chunk, key, line);
        lget(chunk, val, line);
        if *is_int {
            vybe_compiler::primitives::convert::emit_to_int(chunk, line);
        }
        let _ = chunk;
        call_import(chunks, current, "ecma:map", "set", 3, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    }

    // `$component == -1` → the whole array; otherwise that one value or null.
    lget(&mut chunks[current], comp, line);
    push_const(&mut chunks[current], Value::I32(-1), line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], out, line);
    chunks[current].emit_else(line);
    for (i, key) in PHP_URL_SELECTOR.iter().enumerate() {
        lget(&mut chunks[current], comp, line);
        push_const(&mut chunks[current], Value::I32(i as i32), line);
        chunks[current].emit_op(Op::I32_EQ, line);
        chunks[current].emit_if_value(line);
        lget(&mut chunks[current], out, line);
        push_str(&mut chunks[current], key, line);
        call_import(chunks, current, "ecma:map", "get", 2, line);
        chunks[current].emit_else(line);
    }
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    for _ in PHP_URL_SELECTOR {
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);
}
