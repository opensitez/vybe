//! Math compilation — maps language-specific math to WASM opcodes + host imports.
//!
//! WASM has: abs, ceil, floor, trunc, nearest, sqrt, min, max, copysign, neg
//! WASM does NOT have: pow, log, sin, cos, tan, atan2, exp, random
//! Those use host imports (standard across all languages).

use crate::primitives::Target;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

// ── Mathematical constants ──────────────────────────────────
//
// π is π. It is not a .NET surface, a Go surface or a Dart surface, and it does
// not belong to whichever platform happens to be linked — so it lives here,
// once, keyed by the CONCEPT rather than by anyone's spelling of it.
//
// It used to live in eleven places: a `[namespace_constants]` block in each
// language profile, plus `platforms/dotnet`'s `NAMESPACE_CONSTANTS`, which the
// profile parser merged in behind `use_dotnet`. That made the dotnet platform
// the de-facto owner of `Math.PI` — and, for any language that reached the
// merged table, a dependency on a platform it has nothing to do with.

const CONSTANTS: &[(&str, f64)] = &[
    ("pi", std::f64::consts::PI),
    ("e", std::f64::consts::E),
    ("tau", std::f64::consts::TAU),
    ("ln2", std::f64::consts::LN_2),
    ("ln10", std::f64::consts::LN_10),
    ("log2e", std::f64::consts::LOG2_E),
    ("log10e", std::f64::consts::LOG10_E),
    ("sqrt2", std::f64::consts::SQRT_2),
    ("sqrt1_2", std::f64::consts::FRAC_1_SQRT_2),
    ("sqrtpi", 1.772_453_850_905_516),
];

/// The value of a mathematical constant named by its CONCEPT (`pi`, `ln2`,
/// `sqrt2`), case-insensitively — `Pi`, `PI` and `pi` are the same number.
pub fn constant(name: &str) -> Option<f64> {
    CONSTANTS
        .iter()
        .find(|(k, _)| k.eq_ignore_ascii_case(name))
        .map(|(_, v)| *v)
}

/// The value behind a DOTTED constant reference whose owner is the math
/// namespace: `math.pi`, `Math.PI`, `system.math.pi`, `System.Math.Tau`,
/// Go's `math.Pi`, Dart's `math.ln2`.
///
/// The owner segment must actually be spelled `math`, so this never claims an
/// unrelated `Foo.E` or a user class's `.pi`.
pub fn dotted_constant(key: &str) -> Option<f64> {
    let (owner, name) = key.rsplit_once('.')?;
    if !owner
        .rsplit('.')
        .next()
        .is_some_and(|seg| seg.eq_ignore_ascii_case("math"))
    {
        return None;
    }
    constant(name)
}

// ── Direct WASM opcodes (no host call) ──────────────────────

pub fn emit_abs(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_ABS, line);
}

/// C-style fmod: `a - trunc(a/b) * b`. Stack: [a, b] → [result].
/// Pure WASM opcodes — no host import needed.
pub fn emit_c_fmod(chunk: &mut Chunk, line: u32) {
    let b_slot = chunk.local_count;
    let a_slot = chunk.local_count + 1;
    chunk.alloc_scratch(2);
    chunk.emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line); // a
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line); // a (for subtraction later)
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line); // b
    chunk.emit_op(Op::F64_DIV, line); // a/b
    chunk.emit_op(Op::F64_TRUNC, line); // trunc(a/b)
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line); // b
    chunk.emit_op(Op::F64_MUL, line); // trunc(a/b)*b
    chunk.emit_op(Op::F64_SUB, line); // a - trunc(a/b)*b
}

/// Python floor modulo: `a - b * floor(a / b)`. Stack: [a, b] → [result].
/// Differs from C fmod (which truncates toward zero).
pub fn emit_python_floor_mod(chunk: &mut Chunk, line: u32) {
    let b_slot = chunk.local_count;
    let a_slot = chunk.local_count + 1;
    chunk.alloc_scratch(2);
    chunk.emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a_slot, line);
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

/// clamp(x, min, max) = min(max(x, min), max). Stack: [x, min, max] → [result].
/// Pure WASM — no host import needed.
pub fn emit_clamp(chunk: &mut Chunk, line: u32) {
    let max_slot = chunk.local_count;
    let min_slot = chunk.local_count + 1;
    chunk.alloc_scratch(2);
    chunk.emit_op_u16(Op::LOCAL_SET, max_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, min_slot, line);
    // stack: [x]
    chunk.emit_op_u16(Op::LOCAL_GET, min_slot, line); // [x, min]
    chunk.emit_op(Op::F64_MAX, line); // [max(x, min)]
    chunk.emit_op_u16(Op::LOCAL_GET, max_slot, line); // [max(x,min), max]
    chunk.emit_op(Op::F64_MIN, line); // [min(max(x,min), max)]
}

// ── Host imports (standard math, same across all languages) ──
/// Pow via direct ECMA host import. Stack: [base, exponent] → [result].
pub fn emit_pow(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:math", "pow");
    chunk.emit_call(idx, 2, line);
}

/// Stack: [value] → [result]
pub fn emit_log(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:math", "log");
    chunk.emit_call(idx, 1, line);
}

pub fn emit_sin(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:math", "sin");
    chunk.emit_call(idx, 1, line);
}

pub fn emit_cos(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:math", "cos");
    chunk.emit_call(idx, 1, line);
}

pub fn emit_tan(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:math", "tan");
    chunk.emit_call(idx, 1, line);
}

pub fn emit_exp(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:math", "exp");
    chunk.emit_call(idx, 1, line);
}

/// Stack: [] → [f64 random 0..1]
pub fn emit_random(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:math", "random");
    chunk.emit_call(idx, 0, line);
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
        chunk.emit_call(idx, 2, line);
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
    chunk.emit_call(idx, 1, line);
}

// ── IEEE-754 float semantics (copysign / sign bit / bit reinterpret) ──────
//
// Generic WASM compositions shared across languages (Go `math`, C `math.h`,
// Python `math`). Stack contract matches the profile builtin/value-method
// convention: operands are already pushed left-to-right.

/// `copysign(x, y)` — magnitude of `x` with the sign of `y`. Stack: `[x, y]`.
pub fn emit_copysign(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_COPYSIGN, line);
}

/// `signbit(x)` — true when the IEEE sign bit is set (including `-0`). Detected
/// via `copysign(1, x) < 0`. Stack: `[x]` → boolean.
pub fn emit_signbit(chunk: &mut Chunk, line: u32) {
    let base = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, base, line); // stash x
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op(Op::F64_COPYSIGN, line); // ±1
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_LT, line); // < 0 → i32
    crate::primitives::ops::emit_i32_to_bool(chunk, line);
}

/// `dim(x, y)` — positive difference `max(x - y, 0)` (C `fdim`, Go `math.Dim`).
/// Stack: `[x, y]`.
pub fn emit_dim(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_MAX, line);
}

/// A quiet NaN constant. Stack: `[]` → NaN.
pub fn emit_nan(chunk: &mut Chunk, line: u32) {
    chunk.emit_f64_const(f64::NAN, line);
}

/// Go `math.Inf(sign)` — `+Inf` when `sign >= 0`, else `-Inf`.
/// `copysign(+Inf, sign)` yields exactly that (sign 0 → positive). Stack: `[sign]`.
pub fn emit_inf(chunk: &mut Chunk, line: u32) {
    let base = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, base, line); // stash sign
    chunk.emit_f64_const(f64::INFINITY, line);
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op(Op::F64_COPYSIGN, line);
}

/// Go `math.IsInf(x, sign)` — `(x == +Inf && sign >= 0) || (x == -Inf && sign <= 0)`.
/// Stack: `[x, sign]` → boolean.
pub fn emit_is_inf(chunk: &mut Chunk, line: u32) {
    let base = chunk.alloc_scratch(2);
    chunk.emit_op_u16(Op::LOCAL_SET, base + 1, line); // sign
    chunk.emit_op_u16(Op::LOCAL_SET, base, line); // x
    // x == +Inf
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_f64_const(f64::INFINITY, line);
    chunk.emit_op(Op::F64_EQ, line);
    // sign >= 0
    chunk.emit_op_u16(Op::LOCAL_GET, base + 1, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_GE, line);
    chunk.emit_op(Op::I32_AND, line);
    // x == -Inf
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_f64_const(f64::NEG_INFINITY, line);
    chunk.emit_op(Op::F64_EQ, line);
    // sign <= 0
    chunk.emit_op_u16(Op::LOCAL_GET, base + 1, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_LE, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_OR, line);
    crate::primitives::ops::emit_i32_to_bool(chunk, line);
}

/// Reinterpret an `f64` as its raw `u64` bits (Go `math.Float64bits`).
pub fn emit_f64_bits(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I64_REINTERPRET_F64, line);
}

/// Reinterpret raw `u64` bits as an `f64` (Go `math.Float64frombits`).
pub fn emit_f64_from_bits(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_REINTERPRET_I64, line);
}

/// Reinterpret an `f32` (narrowed from `f64`) as its raw `u32` bits
/// (Go `math.Float32bits`).
pub fn emit_f32_bits(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F32_DEMOTE_F64, line);
    chunk.emit_op(Op::I32_REINTERPRET_F32, line);
}

/// Reinterpret raw `u32` bits as an `f32`, widened back to `f64`
/// (Go `math.Float32frombits`).
pub fn emit_f32_from_bits(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F32_REINTERPRET_I32, line);
    chunk.emit_op(Op::F64_PROMOTE_F32, line);
}
