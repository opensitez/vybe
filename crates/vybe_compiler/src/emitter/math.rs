//! Math compilation — maps language-specific math to WASM opcodes + host imports.
//!
//! WASM has: abs, ceil, floor, trunc, nearest, sqrt, min, max, copysign, neg
//! WASM does NOT have: pow, log, sin, cos, tan, atan2, exp, random
//! Those use host imports (standard across all languages).

use crate::emitter::Target;
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

// ── Direct WASM opcodes (no host call) ──────────────────────

pub fn emit_abs(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_ABS, line);
}

/// C-style fmod: `a - trunc(a/b) * b`. Stack: [a, b] → [result].
/// Pure WASM opcodes — no host import needed.
pub fn emit_c_fmod(chunk: &mut Chunk, line: u32) {
    let b_slot = chunk.local_count;
    let a_slot = chunk.local_count + 1;
    chunk.local_count += 2;
    chunk.emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line); // a
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line); // a (for subtraction later)
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line); // b
    chunk.emit_op(Op::F64_DIV, line);               // a/b
    chunk.emit_op(Op::F64_TRUNC, line);             // trunc(a/b)
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line); // b
    chunk.emit_op(Op::F64_MUL, line);               // trunc(a/b)*b
    chunk.emit_op(Op::F64_SUB, line);               // a - trunc(a/b)*b
}

/// Python floor modulo: `a - b * floor(a / b)`. Stack: [a, b] → [result].
/// Differs from C fmod (which truncates toward zero).
pub fn emit_python_floor_mod(chunk: &mut Chunk, line: u32) {
    let b_slot = chunk.local_count;
    let a_slot = chunk.local_count + 1;
    chunk.local_count += 2;
    chunk.emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line); // a
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line); // a
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line); // b
    chunk.emit_op(Op::F64_DIV, line); // a/b
    chunk.emit_op(Op::F64_FLOOR, line); // floor(a/b)
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line); // b
    chunk.emit_op(Op::F64_MUL, line); // b * floor(a/b)
    chunk.emit_op(Op::F64_SUB, line); // a - b*floor(a/b)
}
pub fn emit_floor(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_FLOOR, line);
}
pub fn emit_ceil(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_CEIL, line);
}
pub fn emit_trunc(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_TRUNC, line);
}
pub fn emit_round(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_NEAREST, line);
}
pub fn emit_sqrt(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_SQRT, line);
}
pub fn emit_min(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_MIN, line);
}
pub fn emit_max(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_MAX, line);
}

pub fn emit_neg(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_NEG, line);
}

// ── Host imports (standard math, same across all languages) ──

use std::sync::Arc;
use vybe_bytecode::Value;

/// Legacy: pow via host import. Stack: [base, exponent] → [result].
/// Prefer `emit_pow_push_func` + args + `emit_pow_invoke` for the canonical
/// stdlib path (pure WASM bytecode, runtime-replaceable by Vybe with optimized
/// native pow). The bundle aliases ecma:math.pow to __vybe_pow as a fallback,
/// so this still works, but new code should use the explicit stdlib pattern.
pub fn emit_pow(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:math", "pow");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(2, line);
}

/// Push the __vybe_pow func ref. Use BEFORE compiling base/exponent.
/// Pure WASM — bundle wires `__vybe_pow` to `build_pow` stdlib chunk.
pub fn emit_pow_push_func(chunk: &mut Chunk, line: u32) {
    let name = chunk.add_constant(Value::String(Arc::from("__vybe_pow")));
    chunk.emit_op_u16(Op::GLOBAL_GET, name, line);
}

/// Invoke __vybe_pow after [func, base, exponent] are on the stack.
pub fn emit_pow_invoke(chunk: &mut Chunk, line: u32) {
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
}

/// Stack: [value] → [result]
pub fn emit_log(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:math", "log");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}

pub fn emit_sin(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:math", "sin");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}

pub fn emit_cos(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:math", "cos");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}

pub fn emit_tan(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:math", "tan");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}

pub fn emit_exp(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:math", "exp");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}

/// Stack: [] → [f64 random 0..1]
pub fn emit_random(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:math", "random");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(0, line);
}

// ── Target-aware variants ───────────────────────────────────
// On Vybe: use ecma:math host imports.
// On standard WASM: these must be provided by the embedder or linked from libm.

/// Target-aware pow. Stack: [base, exp] → [result]
pub fn emit_pow_targeted(chunk: &mut Chunk, target: &Target, line: u32) {
    if target.has_module("ecma:math") {
        emit_pow(chunk, line);
    } else {
        // Standard WASM fallback: import from a portable math module.
        // Any compliant embedder must provide "env"/"pow" or "math"/"pow".
        let idx = chunk.add_import("env", "pow");
        chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
        chunk.emit(2, line);
    }
}

/// Target-aware sin/cos/tan/log/exp — all follow same pattern.
pub fn emit_math_fn_targeted(chunk: &mut Chunk, name: &str, target: &Target, line: u32) {
    let (module, func) = if target.has_module("ecma:math") {
        ("ecma:math", name)
    } else {
        ("env", name)
    };
    let idx = chunk.add_import(module, func);
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}
