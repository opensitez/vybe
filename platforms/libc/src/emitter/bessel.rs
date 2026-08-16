//! Bessel functions of the first and second kind — `j0`/`j1`/`jn`/`y0`/`y1`/`yn`.
//!
//! These live in `platforms/libc` beside `erf`/`tgamma` because they are libm's
//! (`math.h` declares all six) and because the tree registers them under
//! `libc.math.*`, so any language resolves them without depending on the C
//! frontend: C gets them from `<math.h>`, Fortran from `bessel_j0` and friends,
//! and anything else through the same namespace.
//!
//! ACCURACY — read before relying on these. The kernels are the Abramowitz &
//! Stegun 9.4 polynomial approximations, which carry an absolute error near
//! 1e-8 for J0/J1 and 1e-7 for Y0/Y1. That is far short of libm's correctly
//! rounded double precision. It is enough for single-precision `real` and for
//! comparisons to ~7 significant digits, and NOT enough to compare a
//! `real(kind=8)` result digit-for-digit against gfortran. Anything needing
//! true double precision needs a different kernel (Cephes-style rational
//! minimax), not a tweak to these coefficients.
//!
//! `jn`/`yn` use the standard recurrences from the order-0 and order-1 values.
//! Upward recurrence is stable for Y and for J only while n < x; J's downward
//! (Miller) recurrence is what a production implementation uses for n >= x, and
//! is deliberately not attempted here — see the note on `emit_jn`.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn alloc(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn call_math(chunks: &mut [Chunk], current: usize, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import("ecma:math", name);
    chunks[current].emit_call(idx, argc, line);
}

/// Evaluate a polynomial in `t` by Horner's method, leaving the result on the
/// stack. `coeffs` is highest power FIRST.
///
/// Written once because the six kernels below are eleven polynomials between
/// them; emitting each coefficient by hand the way `erf` does would be both
/// unreadable and easy to get subtly wrong.
fn horner(chunks: &mut [Chunk], current: usize, t: u16, coeffs: &[f64], line: u32) {
    let mut first = true;
    for c in coeffs {
        if first {
            chunks[current].emit_f64_const(*c, line);
            first = false;
        } else {
            lget(&mut chunks[current], t, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_f64_const(*c, line);
            chunks[current].emit_op(Op::F64_ADD, line);
        }
    }
}

// Coefficients live in `vybe_compiler::primitives::math` so C and Fortran —
// and anything else — share one set. The first version of this file kept its
// own, with J1's small-argument terms mistyped and J1/Y1 reusing J0's
// large-argument pair; that produced a 5.7% error at `j1(5)` and a
// sign-inverted `y1`.
use vybe_compiler::primitives::math::{F0, F1, J0_SMALL, J1_SMALL, T0, T1, Y0_SMALL, Y1_SMALL};

/// Amplitude/phase form shared by J0 and Y0 for |x| > 3:
/// `sqrt(2/(pi*x)) * f(3/x) * trig(x - pi/4 + theta(3/x))`.
fn emit_large_arg(
    chunks: &mut [Chunk],
    current: usize,
    x: u16,
    amp: &[f64],
    phase: &[f64],
    use_sin: bool,
    line: u32,
) {
    let t = alloc(&mut chunks[current]);
    chunks[current].emit_f64_const(3.0, line);
    lget(&mut chunks[current], x, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    lset(&mut chunks[current], t, line);

    // sqrt(1/x) * f(t)
    chunks[current].emit_f64_const(1.0, line);
    lget(&mut chunks[current], x, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_SQRT, line);
    horner(chunks, current, t, amp, line);
    chunks[current].emit_op(Op::F64_MUL, line);

    // angle = x + theta(t); theta already carries the -pi/4 (J0/Y0) or
    // -3pi/4 (J1/Y1) constant term as its lowest-order coefficient.
    lget(&mut chunks[current], x, line);
    horner(chunks, current, t, phase, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    call_math(chunks, current, if use_sin { "sin" } else { "cos" }, 1, line);
    chunks[current].emit_op(Op::F64_MUL, line);
}

/// `j0(x)` — Bessel function of the first kind, order 0.
pub fn emit_j0(chunks: &mut [Chunk], current: usize, line: u32) {
    let x = alloc(&mut chunks[current]);
    lset(&mut chunks[current], x, line);
    // Even function: fold the sign away so one kernel covers both halves.
    lget(&mut chunks[current], x, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    lset(&mut chunks[current], x, line);

    lget(&mut chunks[current], x, line);
    chunks[current].emit_f64_const(3.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_if(line);
    let t = alloc(&mut chunks[current]);
    lget(&mut chunks[current], x, line);
    chunks[current].emit_f64_const(3.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    let over3 = alloc(&mut chunks[current]);
    lset(&mut chunks[current], over3, line);
    lget(&mut chunks[current], over3, line);
    lget(&mut chunks[current], over3, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    lset(&mut chunks[current], t, line);
    horner(chunks, current, t, J0_SMALL, line);
    chunks[current].emit_else(line);
    emit_large_arg(chunks, current, x, F0, T0, false, line);
    chunks[current].emit_end(line);
}

/// `j1(x)` — order 1. ODD, so the sign is restored after the even kernel.
pub fn emit_j1(chunks: &mut [Chunk], current: usize, line: u32) {
    let signed = alloc(&mut chunks[current]);
    lset(&mut chunks[current], signed, line);
    let x = alloc(&mut chunks[current]);
    lget(&mut chunks[current], signed, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    lset(&mut chunks[current], x, line);

    lget(&mut chunks[current], x, line);
    chunks[current].emit_f64_const(3.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_if(line);
    // J1(x) = x * poly((x/3)^2)
    let t = alloc(&mut chunks[current]);
    let over3 = alloc(&mut chunks[current]);
    lget(&mut chunks[current], x, line);
    chunks[current].emit_f64_const(3.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    lset(&mut chunks[current], over3, line);
    lget(&mut chunks[current], over3, line);
    lget(&mut chunks[current], over3, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    lset(&mut chunks[current], t, line);
    lget(&mut chunks[current], signed, line);
    horner(chunks, current, t, J1_SMALL, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_else(line);
    emit_large_arg(chunks, current, x, F1, T1, false, line);
    // Restore the sign for negative x: J1(-x) = -J1(x).
    lget(&mut chunks[current], signed, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(-1.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `y0(x)` — Bessel function of the SECOND kind, order 0.
///
/// Undefined for x <= 0 (it diverges to -inf at 0), which is what the guard
/// returns rather than a silently wrong finite number.
pub fn emit_y0(chunks: &mut [Chunk], current: usize, line: u32) {
    let x = alloc(&mut chunks[current]);
    lset(&mut chunks[current], x, line);

    lget(&mut chunks[current], x, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(f64::NEG_INFINITY, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], x, line);
    chunks[current].emit_f64_const(3.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_if(line);
    // Y0 = (2/pi) ln(x/2) J0(x) + poly((x/3)^2)
    chunks[current].emit_f64_const(std::f64::consts::FRAC_2_PI, line);
    lget(&mut chunks[current], x, line);
    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    call_math(chunks, current, "log", 1, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    lget(&mut chunks[current], x, line);
    emit_j0(chunks, current, line);
    chunks[current].emit_op(Op::F64_MUL, line);

    let t = alloc(&mut chunks[current]);
    let over3 = alloc(&mut chunks[current]);
    lget(&mut chunks[current], x, line);
    chunks[current].emit_f64_const(3.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    lset(&mut chunks[current], over3, line);
    lget(&mut chunks[current], over3, line);
    lget(&mut chunks[current], over3, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    lset(&mut chunks[current], t, line);
    horner(chunks, current, t, Y0_SMALL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_else(line);
    emit_large_arg(chunks, current, x, F0, T0, true, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `y1(x)` — Bessel function of the second kind, order 1.
///
/// A&S 9.4.5 for x <= 3: `x*Y1(x) = (2/pi)*x*ln(x/2)*J1(x) + poly((x/3)^2)`.
/// The first version of this reused J0's kernel with `sin` instead of `cos`
/// and called it Y1; that is not a phase shift of anything and returned the
/// wrong sign.
pub fn emit_y1(chunks: &mut [Chunk], current: usize, line: u32) {
    let x = alloc(&mut chunks[current]);
    lset(&mut chunks[current], x, line);

    lget(&mut chunks[current], x, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(f64::NEG_INFINITY, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], x, line);
    chunks[current].emit_f64_const(3.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_if(line);
    // (2/pi) * ln(x/2) * J1(x)
    chunks[current].emit_f64_const(std::f64::consts::FRAC_2_PI, line);
    lget(&mut chunks[current], x, line);
    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    call_math(chunks, current, "log", 1, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    lget(&mut chunks[current], x, line);
    emit_j1(chunks, current, line);
    chunks[current].emit_op(Op::F64_MUL, line);

    // + poly((x/3)^2) / x   — A&S gives the series for x*Y1(x).
    let t = alloc(&mut chunks[current]);
    let over3 = alloc(&mut chunks[current]);
    lget(&mut chunks[current], x, line);
    chunks[current].emit_f64_const(3.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    lset(&mut chunks[current], over3, line);
    lget(&mut chunks[current], over3, line);
    lget(&mut chunks[current], over3, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    lset(&mut chunks[current], t, line);
    horner(chunks, current, t, Y1_SMALL, line);
    lget(&mut chunks[current], x, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_else(line);
    emit_large_arg(chunks, current, x, F1, T1, true, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `jn(n, x)` / `yn(n, x)` — order `n` by the three-term recurrence
/// `Z_{k+1} = (2k/x) Z_k - Z_{k-1}`, seeded from the order-0 and order-1
/// kernels.
///
/// STABILITY: upward recurrence is stable for Y at every order, and for J only
/// while `n < x`. Above that J's upward direction amplifies rounding, and a
/// production implementation switches to Miller's downward recurrence. That
/// switch is NOT implemented, so `jn` with `n` well above `x` degrades
/// smoothly rather than failing loudly — check against a reference before
/// relying on it there.
///
/// Stack: `[n, x]` → `[value]`.
fn emit_recurrence(chunks: &mut [Chunk], current: usize, use_y: bool, line: u32) {
    let x = alloc(&mut chunks[current]);
    let n = alloc(&mut chunks[current]);
    lset(&mut chunks[current], x, line);
    lset(&mut chunks[current], n, line);

    let prev = alloc(&mut chunks[current]);
    let cur = alloc(&mut chunks[current]);
    let k = alloc(&mut chunks[current]);
    let tmp = alloc(&mut chunks[current]);

    lget(&mut chunks[current], x, line);
    if use_y {
        emit_y0(chunks, current, line);
    } else {
        emit_j0(chunks, current, line);
    }
    lset(&mut chunks[current], prev, line);
    lget(&mut chunks[current], x, line);
    if use_y {
        emit_y1(chunks, current, line);
    } else {
        emit_j1(chunks, current, line);
    }
    lset(&mut chunks[current], cur, line);

    lget(&mut chunks[current], n, line);
    chunks[current].emit_f64_const(0.5, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], prev, line);
    chunks[current].emit_else(line);

    chunks[current].emit_f64_const(1.0, line);
    lset(&mut chunks[current], k, line);
    // The shared block+loop pair, NOT a hand-rolled `emit_loop_s` + `emit_loop`.
    // The hand-rolled version never closed the loop's own block and never
    // patched it, so every `end` after it closed the wrong scope — a
    // `bessel_jn` call inside an `if` BODY then made that body run
    // unconditionally, with the condition evaluating correctly and being
    // ignored. Unbalanced blocks in an adapter corrupt the ENCLOSING
    // structure, and the symptom appears nowhere near the cause.
    let loop_state =
        vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], k, line);
    lget(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::F64_LT, line);
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    chunks[current].emit_f64_const(2.0, line);
    lget(&mut chunks[current], k, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    lget(&mut chunks[current], x, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    lget(&mut chunks[current], cur, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    lget(&mut chunks[current], prev, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    lset(&mut chunks[current], tmp, line);
    lget(&mut chunks[current], cur, line);
    lset(&mut chunks[current], prev, line);
    lget(&mut chunks[current], tmp, line);
    lset(&mut chunks[current], cur, line);
    lget(&mut chunks[current], k, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    lset(&mut chunks[current], k, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    lget(&mut chunks[current], cur, line);
    chunks[current].emit_end(line);
}

pub fn emit_jn(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_recurrence(chunks, current, false, line);
}

pub fn emit_yn(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_recurrence(chunks, current, true, line);
}

/// `erfcx(x)` — the SCALED complementary error function, `exp(x^2)*erfc(x)`.
///
/// Computing it as written overflows: at x=10, `exp(100)` is 2.7e43 and
/// `erfc(10)` is 1.4e-45, so the product is 0.056 while each factor is at the
/// edge of the range — and our `erfc` underflows to exactly 0 there, making the
/// whole thing 0. The scaled form exists precisely so that neither factor is
/// ever formed.
///
/// Small x keeps the direct product, which is accurate and cheap. Large x uses
/// the asymptotic series `1/(x*sqrt(pi)) * (1 - 1/(2x^2) + 3/(4x^4) - ...)`,
/// which is where the underflow would otherwise be.
pub fn emit_erfcx(chunks: &mut [Chunk], current: usize, line: u32) {
    let x = alloc(&mut chunks[current]);
    lset(&mut chunks[current], x, line);

    lget(&mut chunks[current], x, line);
    chunks[current].emit_f64_const(4.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    // exp(x*x) * erfc(x)
    lget(&mut chunks[current], x, line);
    lget(&mut chunks[current], x, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    call_math(chunks, current, "exp", 1, line);
    lget(&mut chunks[current], x, line);
    super::dispatch::emit_erfc_public(chunks, current, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_else(line);
    // 1/(x*sqrt(pi)) * series(1/x^2)
    let t = alloc(&mut chunks[current]);
    chunks[current].emit_f64_const(1.0, line);
    lget(&mut chunks[current], x, line);
    lget(&mut chunks[current], x, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    lset(&mut chunks[current], t, line);

    chunks[current].emit_f64_const(1.0, line);
    lget(&mut chunks[current], x, line);
    chunks[current].emit_f64_const(std::f64::consts::PI.sqrt(), line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    // 1 - t/2 + 3t^2/4 - 15t^3/8 + 105t^4/16
    horner(chunks, current, t, &[6.562_5, -1.875, 0.75, -0.5, 1.0], line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_end(line);
}

// ── name dispatch ────────────────────────────────────────────────────────────

pub fn emit_bessel(name: &str, chunks: &mut [Chunk], current: usize, line: u32) -> bool {
    match name {
        "libc.math.j0" => emit_j0(chunks, current, line),
        "libc.math.j1" => emit_j1(chunks, current, line),
        "libc.math.y0" => emit_y0(chunks, current, line),
        "libc.math.y1" => emit_y1(chunks, current, line),
        "libc.math.jn" => emit_jn(chunks, current, line),
        "libc.math.yn" => emit_yn(chunks, current, line),
        "libc.math.erfcx" => emit_erfcx(chunks, current, line),
        _ => return false,
    }
    true
}
