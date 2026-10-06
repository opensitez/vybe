//! Type conversion compilation — maps language-specific casts to WASM ops + host imports.
//!
//! WASM has: i32.trunc_f64_s, f64.convert_i32_s, i32.wrap_i64, etc.
//! String conversion uses host imports (not in WASM spec).

use crate::primitives::Target;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

// ── Direct WASM opcodes ─────────────────────────────────────

/// float → int (truncate). Stack: [f64] → [i32]
pub fn emit_to_int(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_FROM_F64, line);
}

/// int → float. Stack: [i32] → [f64]
pub fn emit_to_float(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_FROM_I32, line);
}

/// i64 → i32. Stack: [i64] → [i32]
pub fn emit_i32_wrap(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_WRAP_I64, line);
}

/// i32 → i64. Stack: [i32] → [i64]
pub fn emit_i64_extend(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I64_EXTEND_I32_S, line);
}

// ── Host imports (string conversions) ───────────────────────

/// Any value → string representation. Stack: [value] → [string]
/// Render an f64 the way a language with a distinct `double` type does:
/// an integral value keeps its `.0`, and the IEEE specials are named.
///
/// ⛔ THIS EXISTED THREE TIMES. dart (`string_adapter::emit_dart_double_to_string`),
/// java (`platforms/jvm/.../list_adapter::emit_double_to_string`) and kotlin
/// (`emitter/numbers::emit_double_to_string`) each wrote it, and **each handled a
/// different subset of the edge cases** — dart had NaN and infinities, kotlin had
/// signed zero, java had neither. That is not three policies; it is one mechanic
/// that drifted, and the drift was a real bug: java printed `0.0` for `-0.0`
/// where Java prints `-0.0`.
///
/// ECMA's `ToString` renders `10.0` as `"10"` and `-0.0` as `"0"`, so this is a
/// `primitives/*` gap by definition — a thing real languages need that ECMA does
/// not hand you — not a licence for each frontend to write its own.
///
/// The general case still delegates to `ecma:number.toString`, which all three
/// already did, so only the specials are decided here.
pub fn emit_float_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    use crate::primitives::instructions::host;
    let v = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v, line);

    // These Number predicates return canonical Booleans, not arbitrary values.
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    host::emit(&mut chunks[current], "ecma:number", "isNaN", 1, line);
    host::emit(&mut chunks[current], "wasm:js-boolean", "cast", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("NaN", line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    host::emit(&mut chunks[current], "ecma:number", "isFinite", 1, line);
    host::emit(&mut chunks[current], "wasm:js-boolean", "cast", 1, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    emit_float_signbit(&mut chunks[current], v, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("-Infinity", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("Infinity", line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);

    // Both tests are pure numeric instructions. Combining their i32 results
    // keeps the ordinary finite formatting body at one emission site.
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    emit_float_signbit(&mut chunks[current], v, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("-0.0", line);
    chunks[current].emit_else(line);
    emit_float_to_string_finite(chunks, current, v, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// The numeric sign-bit predicate, consumed immediately as an i32 condition.
fn emit_float_signbit(chunk: &mut Chunk, v: u16, line: u32) {
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, v, line);
    chunk.emit_op(Op::F64_COPYSIGN, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_LT, line);
}

/// An integral double keeps one fractional digit (`10` → `"10.0"`); anything
/// else renders as ECMA does.
fn emit_float_to_string_finite(chunks: &mut [Chunk], current: usize, v: u16, line: u32) {
    use crate::primitives::instructions::host;
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    host::emit(&mut chunks[current], "ecma:number", "isInteger", 1, line);
    host::emit(&mut chunks[current], "wasm:js-boolean", "cast", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    chunks[current].emit_i32_const(1, line);
    host::emit(&mut chunks[current], "ecma:number", "toFixed", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    host::emit(&mut chunks[current], "ecma:number", "toString", 1, line);
    chunks[current].emit_end(line);
}

pub fn emit_to_string(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "String");
    chunk.emit_call(idx, 1, line);
}

/// String → int (parse). Stack: [string] → [i32]
///
/// Uses `ecma:number.parseInt` (§19.2.5) — stops at the first non-digit
/// so `parseInt("3.7") = 3`, matching VB `CInt`/Python `int`/PHP `intval`
/// semantics for string parsing. `Number(x)` would return 3.7 here.
pub fn emit_parse_int(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:number", "parseInt");
    chunk.emit_call(idx, 1, line);
}

/// String → float (parse). Stack: [string] → [f64]
pub fn emit_parse_float(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(idx, 1, line);
}

/// Dynamic truthiness conversion. Stack: [value] → [bool]
pub fn emit_to_bool(chunk: &mut Chunk, line: u32) {
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
}

// ── Target-aware variants ───────────────────────────────────

/// Target-aware toString. Emits `ecma:string.String` on all targets.
pub fn emit_to_string_targeted(chunk: &mut Chunk, _target: &Target, line: u32) {
    emit_to_string(chunk, line);
}

// ── Linkable chunk builders ──────────────────────────────────────────────────
//
// Linkable chunk builders — the standalone-chunk packaging of what the
// `emit_*` forms above splice inline. Same concept, same module.

// ── isNumeric(value) → bool ─────────────────────────────────
// Check existing numeric tags or a primitive string with a numeric prefix.
pub fn build_is_numeric(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_isnumeric");
    c.arity = 1;
    c.local_count = 2;
    let test_number = c.add_import("wasm:js-number", "test");
    let test_string = c.add_import("wasm:js-string", "test");
    let parse = c.add_import("ecma:number", "parseFloat");
    let done = c.emit_block(0);
    // Same numeric tag set as ecma:value.typeof == "number", including
    // F32/I32/I64 and NaN. Result remains the original boxed-I320/1 ABI.
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(test_number, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(test_string, 1, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0); // all other primitive/object tags retain false
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(parse, 1, 0);
    c.emit_dup(0);
    c.emit_op(Op::F64_EQ, 0);
    c.emit_op_u16(Op::LOCAL_SET, 1, 0);
    c.emit_end(0);
    c.patch_block(done);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── val(s) → number — VB Val: parseFloat with NaN→0 fallback ─────
//
// VB `Val(s)` parses a numeric prefix from the string, returning 0
// for non-numeric / empty input. `ecma:number.parseFloat` matches the
// "stop at first non-numeric" semantic; the only divergence is that
// parseFloat returns NaN on no-match while VB returns 0. Wrap with an
// `r != r` (NaN sentinel) check and select 0 in that case.
pub fn build_val(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_val");
    c.arity = 1;
    c.local_count = 2; // arg(0), result(1)
    let arg = 0u16;
    let result = 1u16;

    let pf_idx = c.add_import("ecma:number", "parseFloat");

    // result = parseFloat(arg)
    c.emit_op_u16(Op::LOCAL_GET, arg, 0);
    c.emit_call(pf_idx, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    // if (result == result) skip — only NaN compares unequal to itself.
    let done = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::F64_EQ, 0);
    c.emit_br_if(0, 0);

    // result = 0
    let zero = c.add_constant(vybe_runtime::Value::F64(0.0));
    crate::primitives::expressions::emit_const_index(&mut c, zero, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_end(0);
    c.patch_block(done);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// `build_cchar` removed — nothing referenced `__vybe_cchar`.

// ── to_bytes(s) → Uint8Array — Python `bytes(s)` / `s.encode()` ──
//
// Encodes `s` (any value) as UTF-8 bytes via WHATWG `TextEncoder`.
// Single host fn call into `web:encoding.encoderNew` + `encode` —
// pure spec-aligned dispatch. Variadic encoding arg in Python (e.g.
// `bytes(s, "utf-8")`) is ignored: WHATWG `TextEncoder` is fixed to
// UTF-8 by spec.
pub fn build_to_bytes(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_to_bytes");
    c.arity = 1;
    c.local_count = 2;
    let s = 0u16;
    let enc = 1u16;

    let new_idx = c.add_import("web:encoding", "encoderNew");
    let encode_idx = c.add_import("web:encoding", "encode");

    c.emit_call(new_idx, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, enc, 0);

    c.emit_op_u16(Op::LOCAL_GET, enc, 0);
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_call(encode_idx, 2, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── Linkable chunk builders ──────────────────────────────────────────────────
//
// Linkable chunk builders — the standalone-chunk packaging of what the
// `emit_*` forms splice inline. A language prefix in a name records which
// frontend first needed a linkable chunk, not a language-specific meaning.

// ── pyhex/pyoct/pybin(n) → string — Python radix conversions ────────
//
// Python's `hex(5)` returns `"0x5"` (with prefix); the underlying
// `ecma:number.toString(n, radix)` produces just `"5"`. Each chunk
// concatenates the prefix and forwards to the host.
pub fn build_pyradix(_imports: &mut Chunk, name: &str, prefix: &str, radix: i32) -> Chunk {
    let mut c = Chunk::new(name);
    c.arity = 1;
    c.local_count = 1;
    let pref = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(prefix)));
    let r = c.add_constant(vybe_runtime::Value::I32(radix));
    let ts_idx = c.add_import("ecma:number", "toString");
    let vybe_runtime::Value::String(prefix) = c.constants[pref as usize].clone() else {
        unreachable!()
    };
    c.emit_string_const(&prefix, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let vybe_runtime::Value::I32(radix) = c.constants[r as usize] else {
        unreachable!()
    };
    c.emit_i32_const(radix, 0);
    c.emit_call(ts_idx, 2, 0);
    let concat = c.add_import("wasm:js-string", "concat");
    c.emit_call(concat, 2, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
