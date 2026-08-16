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
// Check if value is a number type using ref_typeof opcode.
pub fn build_is_numeric(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_isnumeric");
    c.arity = 1;
    c.local_count = 2; // val(0), result(1)
    let val = 0u16;
    let result = 1u16;

    // VB `IsNumeric(v)`:
    //   typeof(v) ∈ {"number", "i32", "i64"}                → true
    //   typeof(v) == "string" && !isNaN(parseFloat(v))      → true
    //   otherwise                                            → false
    //
    // Block-and-br_if cascade so each positive case short-circuits and
    // the next check is skipped.
    let num_str = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("number")));
    let i32_str = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("i32")));
    let i64_str = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("i64")));
    let str_str = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("string")));

    let done = c.emit_block(0);

    // typeof(v) == "number"
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    crate::primitives::reflection::emit_typeof_in_chunk(&mut c, 0);
    crate::primitives::expressions::emit_const_index(&mut c, num_str, 0);
    crate::primitives::strings::emit_str_equals(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_br_if(0, 0);

    // typeof(v) == "i32"
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    crate::primitives::reflection::emit_typeof_in_chunk(&mut c, 0);
    crate::primitives::expressions::emit_const_index(&mut c, i32_str, 0);
    crate::primitives::strings::emit_str_equals(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_br_if(0, 0);

    // typeof(v) == "i64"
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    crate::primitives::reflection::emit_typeof_in_chunk(&mut c, 0);
    crate::primitives::expressions::emit_const_index(&mut c, i64_str, 0);
    crate::primitives::strings::emit_str_equals(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_br_if(0, 0);

    // typeof(v) == "string" — try parseFloat, accept iff !isNaN
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    crate::primitives::reflection::emit_typeof_in_chunk(&mut c, 0);
    crate::primitives::expressions::emit_const_index(&mut c, str_str, 0);
    crate::primitives::strings::emit_str_equals(&mut c, 0);
    // [is_string]
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // not a string → done with result still false

    // result = !isNaN(parseFloat(v))  ≡  parsed == parsed
    let pf_idx = c.add_import("ecma:number", "parseFloat");
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_call(pf_idx, 1, 0);
    c.emit_dup(0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    c.emit_end(0);
    c.patch_block(done);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
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
pub fn build_val(imports: &mut Chunk) -> Chunk {
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
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
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
pub fn build_pyradix(imports: &mut Chunk, name: &str, prefix: &str, radix: i32) -> Chunk {
    let mut c = Chunk::new(name);
    c.arity = 1;
    c.local_count = 1;
    let pref = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(prefix)));
    let r = c.add_constant(vybe_runtime::Value::I32(radix));
    let ts_idx = c.add_import("ecma:number", "toString");
    crate::primitives::expressions::emit_const_index(&mut c, pref, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    crate::primitives::expressions::emit_const_index(&mut c, r, 0);
    c.emit_call(ts_idx, 2, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
