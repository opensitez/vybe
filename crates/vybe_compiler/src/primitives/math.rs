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

// ── Bessel functions ─────────────────────────────────────────────────────────
//
// Shared here rather than in a platform crate because they are ordinary
// mathematics: Fortran spells them `bessel_j0`, C's `<math.h>` spells them
// `j0`, and both want the same code. Only the SPELLING is per-language, which
// is what a profile row is for.
//
// Kernels are Abramowitz & Stegun 9.4. Each order has its OWN amplitude and
// phase coefficients for the large-argument form — reusing J0's for J1, as a
// first attempt here did, puts `j1(5)` out by 5.7% and makes `y1` come back
// with the wrong SIGN. J1 is not a phase-shifted J0.
//
// Accuracy is ~1e-7 absolute (A&S's stated bound for these polynomials), which
// is single-precision-clean and NOT correctly-rounded double. A `real(kind=8)`
// comparison against libm will differ around the 8th digit.

/// A&S 9.4.1 — J0 for |x| <= 3, in t = (x/3)^2, highest power first.
pub const J0_SMALL: &[f64] = &[
    0.000_210_0, -0.003_944_4, 0.044_447_9, -0.316_386_6, 1.265_620_8, -2.249_999_7, 1.0,
];
/// A&S 9.4.4 — J1(x)/x for |x| <= 3, in t = (x/3)^2.
pub const J1_SMALL: &[f64] = &[
    0.000_011_09, -0.000_317_61, 0.004_433_19, -0.039_542_89, 0.210_935_73, -0.562_499_85, 0.5,
];
/// A&S 9.4.2 — the additive part of Y0 for |x| <= 3, in t = (x/3)^2.
pub const Y0_SMALL: &[f64] = &[
    -0.000_248_46, 0.004_279_16, -0.042_612_14, 0.253_001_17, -0.743_503_84, 0.605_593_66,
    0.367_466_91,
];
/// A&S 9.4.5 — the additive part of x*Y1(x) for |x| <= 3, in t = (x/3)^2.
pub const Y1_SMALL: &[f64] = &[
    0.002_787_3, -0.040_097_6, 0.312_395_1, -1.316_482_7, 2.168_270_9, 0.221_209_1, -0.636_619_8,
];
/// A&S 9.4.3 — J0/Y0 amplitude f0 for x > 3, in t = 3/x.
pub const F0: &[f64] = &[
    0.000_144_76, -0.000_728_05, 0.001_372_37, -0.000_095_12, -0.005_527_40, -0.000_000_77,
    0.797_884_56,
];
/// A&S 9.4.3 — J0/Y0 phase offset theta0 for x > 3, in t = 3/x.
pub const T0: &[f64] = &[
    0.000_135_58, -0.000_293_33, -0.000_541_25, 0.002_625_73, -0.000_039_54, -0.041_663_97,
    -0.785_398_16,
];
/// A&S 9.4.6 — J1/Y1 amplitude f1 for x > 3, in t = 3/x. NOT f0.
pub const F1: &[f64] = &[
    -0.000_200_33, 0.001_136_53, -0.002_495_11, 0.000_171_05, 0.016_596_67, 0.000_001_56,
    0.797_884_56,
];
/// A&S 9.4.6 — J1/Y1 phase offset theta1 for x > 3, in t = 3/x. NOT t0.
pub const T1: &[f64] = &[
    -0.000_291_66, 0.000_798_24, 0.000_743_48, -0.006_378_79, 0.000_056_50, 0.124_996_12,
    -2.356_194_49,
];


// ── Linkable chunk builders ──────────────────────────────────────────────────
//
// The `emit_*` functions above splice instructions into the CALLER. These build
// standalone chunks that compiled code reaches through a `__vybe_*` global —
// same mathematics, different packaging, and they belong beside the inline
// versions rather than in a general helper bag. `channels::build_chan_*` is the
// same arrangement.

// ── floor(n) → int — wraps f64_floor opcode ────────────────
#[allow(dead_code)]
pub fn build_floor(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_floor");
    c.arity = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::F64_FLOOR, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── fmod(a, b) → a % b (floating-point remainder) ──────────
// WASM has no f64.rem. Pure bytecode: a - trunc(a/b) * b.
// Host can override __vybe_fmod with native fmod for performance.
pub fn build_fmod(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_fmod");
    c.arity = 2; // a, b
    c.local_count = 2; // a(0) + b(1)
    let a = 0u16;
    let b = 1u16;

    // result = a - trunc(a / b) * b
    c.emit_op_u16(Op::LOCAL_GET, a, 0); // a
    c.emit_op_u16(Op::LOCAL_GET, a, 0); // a
    c.emit_op_u16(Op::LOCAL_GET, b, 0); // b
    c.emit_op(Op::F64_DIV, 0); // a / b
    c.emit_op(Op::F64_TRUNC, 0); // trunc(a / b)
    c.emit_op_u16(Op::LOCAL_GET, b, 0); // b
    c.emit_op(Op::F64_MUL, 0); // trunc(a / b) * b
    c.emit_op(Op::F64_SUB, 0); // a - trunc(a / b) * b
    c.emit_op(Op::RETURN, 0);
    c
}

// ── isinf(n) → bool — Python `math.isinf`: ±Infinity check ──────────
//
// Composition: `!isFinite(n) && !isNaN(n)` ≡ "infinite, not NaN".
pub fn build_isinf(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_isinf");
    c.arity = 1;
    c.local_count = 1;
    let isfin = c.add_import("ecma:number", "isFinite");
    let isnan = c.add_import("ecma:number", "isNaN");

    // !isFinite(n) && !isNaN(n)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(isfin, 1, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(isnan, 1, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_op(Op::I32_AND, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
