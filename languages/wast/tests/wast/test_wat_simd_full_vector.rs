//! Full-vector SIMD result assertions.
//!
//! The per-shape SIMD suites verify one lane at a time via `extract_lane`,
//! which cannot catch a bug that puts the right value in the WRONG lane, or
//! that only mishandles a lane the extract never reads. These tests assert the
//! ENTIRE resulting `v128` against a `(v128.const …)` expected vector through
//! `assert_return`, so every lane is checked at once — the coverage that
//! matters for the lane-CROSSING ops (shuffle, swizzle, narrow, widen, dot,
//! extmul, extadd_pairwise) where lane position is the whole point.
//!
//! Expected vectors are the WebAssembly SIMD spec results.
use crate::helpers::run_wast_asserts;

/// Assert a module whose exported `f` returns a `v128` equals the expected
/// vector. Panics with the mismatch (got/expected) on failure.
fn v128_eq(func_body: &str, expected: &str) {
    let src = format!(
        "(module (func (export \"f\") (result v128)\n{func_body}))\n\
         (assert_return (invoke \"f\") ({expected}))\n"
    );
    run_wast_asserts(&src).unwrap_or_else(|e| panic!("full-vector assert failed: {e}\n{src}"));
}

/// Assert a module whose exported `f` returns an `i32` (SIMD reductions like
/// `all_true` / `bitmask` collapse the whole vector to a scalar) equals the
/// expected integer — these still exercise EVERY lane, just folded into one
/// result.
fn i32_eq(func_body: &str, expected: i32) {
    let src = format!(
        "(module (func (export \"f\") (result i32)\n{func_body}))\n\
         (assert_return (invoke \"f\") (i32.const {expected}))\n"
    );
    run_wast_asserts(&src).unwrap_or_else(|e| panic!("reduction assert failed: {e}\n{src}"));
}

/// Whole-vector check for an f32x4-PRODUCING body. The `assert_return` expected
/// grammar can't express float lanes inside a `v128.const`, so instead emit one
/// exported function per lane (`<body> f32x4.extract_lane i`) and assert each
/// lane's `f32.const`. Every lane is verified — including NaN lanes (pass e.g.
/// `"f32.const nan:canonical"`, which the harness normalises). This is the
/// float-side equivalent of `v128_eq` without the integer-lane grammar limit.
fn f32x4_eq(body: &str, lanes: [&str; 4]) {
    let mut funcs = String::new();
    let mut asserts = String::new();
    for (i, lane) in lanes.iter().enumerate() {
        funcs.push_str(&format!(
            "  (func (export \"f{i}\") (result f32)\n{body}\n  f32x4.extract_lane {i})\n"
        ));
        asserts.push_str(&format!("(assert_return (invoke \"f{i}\") ({lane}))\n"));
    }
    let src = format!("(module\n{funcs})\n{asserts}");
    run_wast_asserts(&src).unwrap_or_else(|e| panic!("f32x4 lane assert failed: {e}\n{src}"));
}

/// Whole-vector check for an f64x2-producing body (2 lanes). See [`f32x4_eq`].
fn f64x2_eq(body: &str, lanes: [&str; 2]) {
    let mut funcs = String::new();
    let mut asserts = String::new();
    for (i, lane) in lanes.iter().enumerate() {
        funcs.push_str(&format!(
            "  (func (export \"f{i}\") (result f64)\n{body}\n  f64x2.extract_lane {i})\n"
        ));
        asserts.push_str(&format!("(assert_return (invoke \"f{i}\") ({lane}))\n"));
    }
    let src = format!("(module\n{funcs})\n{asserts}");
    run_wast_asserts(&src).unwrap_or_else(|e| panic!("f64x2 lane assert failed: {e}\n{src}"));
}

// ── i8x16.shuffle — arbitrary lane permutation from two vectors ──────────────
// Indices 0..15 select `a`'s lanes, 16..31 select `b`'s. Interleave a/b.
#[test]
fn shuffle_interleave_low() {
    v128_eq(
        "  v128.const i8x16 0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15\n\
         \x20 v128.const i8x16 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31\n\
         \x20 i8x16.shuffle 0 16 1 17 2 18 3 19 4 20 5 21 6 22 7 23",
        "v128.const i8x16 0 16 1 17 2 18 3 19 4 20 5 21 6 22 7 23",
    );
}

// Reverse a single vector's 16 lanes (b ignored).
#[test]
fn shuffle_reverse() {
    v128_eq(
        "  v128.const i8x16 0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15\n\
         \x20 v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0\n\
         \x20 i8x16.shuffle 15 14 13 12 11 10 9 8 7 6 5 4 3 2 1 0",
        "v128.const i8x16 15 14 13 12 11 10 9 8 7 6 5 4 3 2 1 0",
    );
}

// ── i8x16.swizzle — dynamic per-lane select; out-of-range index → 0 ─────────
#[test]
fn swizzle_gather_and_zero() {
    // lane i takes a[index[i]]; indices ≥ 16 (here 16 and 200) yield 0.
    v128_eq(
        "  v128.const i8x16 100 101 102 103 104 105 106 107 108 109 110 111 112 113 114 115\n\
         \x20 v128.const i8x16 15 0 3 16 1 1 200 7 8 9 10 11 12 13 14 2\n\
         \x20 i8x16.swizzle",
        "v128.const i8x16 115 100 103 0 101 101 0 107 108 109 110 111 112 113 114 102",
    );
}

// ── i16x8.narrow_i32x4_s — two i32x4 → one i16x8, signed saturation ─────────
// Result low 4 lanes from the first operand, high 4 from the second; each i32
// saturates to i16 range [-32768, 32767].
#[test]
fn narrow_i32x4_signed_saturates() {
    v128_eq(
        "  v128.const i32x4 0 32767 32768 -32769\n\
         \x20 v128.const i32x4 -40000 40000 -1 1\n\
         \x20 i16x8.narrow_i32x4_s",
        "v128.const i16x8 0 32767 32767 -32768 -32768 32767 -1 1",
    );
}

// ── i16x8.narrow_i32x4_u — unsigned saturation to [0, 65535] ────────────────
#[test]
fn narrow_i32x4_unsigned_saturates() {
    v128_eq(
        "  v128.const i32x4 -1 0 65535 65536\n\
         \x20 v128.const i32x4 100000 -100000 1 2\n\
         \x20 i16x8.narrow_i32x4_u",
        "v128.const i16x8 0 0 65535 65535 65535 0 1 2",
    );
}

// ── i32x4.dot_i16x8_s — multiply adjacent i16 pairs, add into i32 lanes ─────
// out[j] = a[2j]*b[2j] + a[2j+1]*b[2j+1].
#[test]
fn dot_product_pairs() {
    v128_eq(
        "  v128.const i16x8 1 2 3 4 5 6 7 8\n\
         \x20 v128.const i16x8 1 1 2 2 3 3 4 4\n\
         \x20 i32x4.dot_i16x8_s",
        // (1+2, 6+8, 15+18, 28+32) = (3, 14, 33, 60)
        "v128.const i32x4 3 14 33 60",
    );
}

// ── i16x8.extend_low/high_i8x16_s — widen one half, sign-extended ───────────
#[test]
fn extend_low_signed() {
    v128_eq(
        "  v128.const i8x16 -1 -2 3 4 5 6 7 8 100 100 100 100 100 100 100 100\n\
         \x20 i16x8.extend_low_i8x16_s",
        "v128.const i16x8 -1 -2 3 4 5 6 7 8",
    );
}

#[test]
fn extend_high_signed() {
    v128_eq(
        "  v128.const i8x16 0 0 0 0 0 0 0 0 -1 -2 3 4 5 6 7 8\n\
         \x20 i16x8.extend_high_i8x16_s",
        "v128.const i16x8 -1 -2 3 4 5 6 7 8",
    );
}

// ── i16x8.extadd_pairwise_i8x16_u — sum adjacent unsigned byte pairs ─────────
#[test]
fn extadd_pairwise_unsigned() {
    v128_eq(
        "  v128.const i8x16 255 1 10 20 0 0 100 100 1 2 3 4 5 6 7 8\n\
         \x20 i16x8.extadd_pairwise_i8x16_u",
        // (255+1, 10+20, 0, 200, 3, 7, 11, 15)
        "v128.const i16x8 256 30 0 200 3 7 11 15",
    );
}

// ── i16x8.extmul_low_i8x16_s — widen-multiply the low 8 lanes ────────────────
#[test]
fn extmul_low_signed() {
    v128_eq(
        "  v128.const i8x16 -2 3 -4 5 6 7 8 9 0 0 0 0 0 0 0 0\n\
         \x20 v128.const i8x16 10 10 10 10 10 10 10 10 0 0 0 0 0 0 0 0\n\
         \x20 i16x8.extmul_low_i8x16_s",
        "v128.const i16x8 -20 30 -40 50 60 70 80 90",
    );
}

// ── all-lane integer arithmetic verified across EVERY lane at once ───────────
#[test]
fn i32x4_add_all_lanes() {
    v128_eq(
        "  v128.const i32x4 1 2 3 4\n\
         \x20 v128.const i32x4 10 20 30 40\n\
         \x20 i32x4.add",
        "v128.const i32x4 11 22 33 44",
    );
}

#[test]
fn i8x16_add_saturate_signed_all_lanes() {
    v128_eq(
        "  v128.const i8x16 127 127 -128 -128 1 2 3 4 5 6 7 8 9 10 11 12\n\
         \x20 v128.const i8x16 1 127 -1 -128 1 1 1 1 1 1 1 1 1 1 1 1\n\
         \x20 i8x16.add_sat_s",
        "v128.const i8x16 127 127 -128 -128 2 3 4 5 6 7 8 9 10 11 12 13",
    );
}

// ── i16x8.mul wraps per-lane at 16 bits, verified for all 8 lanes ────────────
#[test]
fn i16x8_mul_wraps_all_lanes() {
    v128_eq(
        "  v128.const i16x8 256 256 -1 32767 2 3 4 5\n\
         \x20 v128.const i16x8 256 128 -1 2 2 3 4 5\n\
         \x20 i16x8.mul",
        // 256*256=65536→0, 256*128=32768→-32768, (-1)*(-1)=1, 32767*2=65534→-2
        "v128.const i16x8 0 -32768 1 -2 4 9 16 25",
    );
}

// ── v128.bitselect — per-bit choose from a/b by mask; whole-vector check ─────
#[test]
fn bitselect_whole_vector() {
    v128_eq(
        "  v128.const i32x4 -1 -1 -1 -1\n\
         \x20 v128.const i32x4 0 0 0 0\n\
         \x20 v128.const i32x4 -1 0 -1 0\n\
         \x20 v128.bitselect",
        // mask lane all-ones → a (-1); all-zero → b (0)
        "v128.const i32x4 -1 0 -1 0",
    );
}

// ── i8x16.avgr_u — unsigned rounding average (a+b+1)>>1, all 16 lanes ────────
#[test]
fn avgr_u_rounds_up_all_lanes() {
    v128_eq(
        "  v128.const i8x16 255 1 10 100 0 2 4 6 8 10 12 14 16 18 20 22\n\
         \x20 v128.const i8x16 255 2 13 101 0 2 4 6 8 10 12 14 16 18 20 22\n\
         \x20 i8x16.avgr_u",
        "v128.const i8x16 255 2 12 101 0 2 4 6 8 10 12 14 16 18 20 22",
    );
}

// ── i8x16.narrow_i16x8_s — two i16x8 → i8x16, signed saturation ─────────────
#[test]
fn narrow_i16x8_signed_saturates() {
    v128_eq(
        "  v128.const i16x8 0 127 128 -129 -1 1 200 -200\n\
         \x20 v128.const i16x8 100 -100 127 -128 0 300 -300 5\n\
         \x20 i8x16.narrow_i16x8_s",
        "v128.const i8x16 0 127 127 -128 -1 1 127 -128 100 -100 127 -128 0 127 -128 5",
    );
}

// ── i32x4.extend_low_i16x8_s — low 4 i16 lanes sign-extended to i32 ──────────
#[test]
fn extend_low_i16x8_signed() {
    v128_eq(
        "  v128.const i16x8 -1 -32768 32767 5 9 9 9 9\n\
         \x20 i32x4.extend_low_i16x8_s",
        "v128.const i32x4 -1 -32768 32767 5",
    );
}

// ── i16x8.extmul_high_i8x16_s — widen-multiply the HIGH 8 lanes ──────────────
#[test]
fn extmul_high_signed() {
    v128_eq(
        "  v128.const i8x16 0 0 0 0 0 0 0 0 -2 3 -4 5 6 7 8 9\n\
         \x20 v128.const i8x16 0 0 0 0 0 0 0 0 10 10 10 10 10 10 10 10\n\
         \x20 i16x8.extmul_high_i8x16_s",
        "v128.const i16x8 -20 30 -40 50 60 70 80 90",
    );
}

// ── i32x4.abs — INT_MIN abs wraps to itself; all lanes ──────────────────────
#[test]
fn i32x4_abs_all_lanes() {
    v128_eq(
        "  v128.const i32x4 -5 5 -2147483648 0\n\
         \x20 i32x4.abs",
        "v128.const i32x4 5 5 -2147483648 0",
    );
}

// ── i8x16.min_s — signed per-lane minimum, all lanes ────────────────────────
#[test]
fn i8x16_min_s_all_lanes() {
    v128_eq(
        "  v128.const i8x16 -1 5 -128 127 9 9 9 9 9 9 9 9 9 9 9 9\n\
         \x20 v128.const i8x16 1 -5 127 -128 0 0 0 0 0 0 0 0 0 0 0 0\n\
         \x20 i8x16.min_s",
        "v128.const i8x16 -1 -5 -128 -128 0 0 0 0 0 0 0 0 0 0 0 0",
    );
}

// ── i32x4.shl — logical left shift by scalar, all lanes ─────────────────────
#[test]
fn i32x4_shl_all_lanes() {
    v128_eq(
        "  v128.const i32x4 1 2 3 -1\n\
         \x20 i32.const 4\n\
         \x20 i32x4.shl",
        "v128.const i32x4 16 32 48 -16",
    );
}

// ── i32x4.shr_s — arithmetic right shift by scalar, all lanes ────────────────
#[test]
fn i32x4_shr_s_all_lanes() {
    v128_eq(
        "  v128.const i32x4 16 -16 7 -7\n\
         \x20 i32.const 1\n\
         \x20 i32x4.shr_s",
        "v128.const i32x4 8 -8 3 -4",
    );
}

// ── i64x2.add — both 64-bit lanes, wide values ──────────────────────────────
#[test]
fn i64x2_add_all_lanes() {
    v128_eq(
        "  v128.const i64x2 1000000000000 -5\n\
         \x20 v128.const i64x2 1 5\n\
         \x20 i64x2.add",
        "v128.const i64x2 1000000000001 0",
    );
}

// ── comparisons produce a full per-lane mask (all-ones / all-zero) ──────────
// Single-lane extraction can't see a mask bug in a lane it never reads; assert
// the WHOLE mask vector.
#[test]
fn i32x4_eq_mask() {
    v128_eq(
        "  v128.const i32x4 1 2 3 3\n\
         \x20 v128.const i32x4 1 9 3 3\n\
         \x20 i32x4.eq",
        "v128.const i32x4 -1 0 -1 -1",
    );
}

#[test]
fn i32x4_lt_s_mask() {
    v128_eq(
        "  v128.const i32x4 -5 5 0 100\n\
         \x20 v128.const i32x4 0 0 0 100\n\
         \x20 i32x4.lt_s",
        "v128.const i32x4 -1 0 0 0",
    );
}

#[test]
fn i32x4_ge_s_mask() {
    v128_eq(
        "  v128.const i32x4 5 5 5 5\n\
         \x20 v128.const i32x4 5 6 4 -1\n\
         \x20 i32x4.ge_s",
        "v128.const i32x4 -1 0 -1 -1",
    );
}

#[test]
fn i8x16_eq_mask_all_16_lanes() {
    v128_eq(
        "  v128.const i8x16 1 2 3 4 5 6 7 8 1 2 3 4 5 6 7 8\n\
         \x20 v128.const i8x16 1 0 3 0 5 0 7 0 1 0 3 0 5 0 7 0\n\
         \x20 i8x16.eq",
        "v128.const i8x16 -1 0 -1 0 -1 0 -1 0 -1 0 -1 0 -1 0 -1 0",
    );
}

// f32x4 comparison yields an INTEGER mask vector (verifiable full-vector).
#[test]
fn f32x4_eq_yields_integer_mask() {
    v128_eq(
        "  v128.const f32x4 1.0 2.0 3.0 3.0\n\
         \x20 v128.const f32x4 1.0 9.0 3.0 3.0\n\
         \x20 f32x4.eq",
        "v128.const i32x4 -1 0 -1 -1",
    );
}

#[test]
fn f32x4_lt_yields_integer_mask() {
    v128_eq(
        "  v128.const f32x4 -1.5 5.0 2.0 2.0\n\
         \x20 v128.const f32x4 0.0 0.0 2.0 3.0\n\
         \x20 f32x4.lt",
        "v128.const i32x4 -1 0 0 -1",
    );
}

// ── trunc_sat: float input (built in the body) → saturating integer result ──
// The `assert_return` expected grammar only allows integer lanes, but the
// result IS an i32x4 — so saturation of out-of-range/negative floats is fully
// checkable across every lane.
#[test]
fn i32x4_trunc_sat_f32x4_s_saturates() {
    v128_eq(
        "  v128.const f32x4 3.9 -3.9 3000000000.0 -3000000000.0\n\
         \x20 i32x4.trunc_sat_f32x4_s",
        // 3.9→3, -3.9→-3, +3e9>INT_MAX→2147483647, -3e9<INT_MIN→-2147483648
        "v128.const i32x4 3 -3 2147483647 -2147483648",
    );
}

#[test]
fn i32x4_trunc_sat_f32x4_u_saturates() {
    v128_eq(
        "  v128.const f32x4 3.9 -1.0 5000000000.0 100.5\n\
         \x20 i32x4.trunc_sat_f32x4_u",
        // 3.9→3, negative→0, 5e9>UINT_MAX→0xFFFFFFFF (-1 as i32), 100.5→100
        "v128.const i32x4 3 0 -1 100",
    );
}

// ── reductions fold the whole vector to a scalar; every lane participates ────
#[test]
fn i8x16_all_true_false_when_any_zero() {
    i32_eq(
        "  v128.const i8x16 1 1 1 1 1 1 1 0 1 1 1 1 1 1 1 1\n\
         \x20 i8x16.all_true",
        0,
    );
}

#[test]
fn i8x16_all_true_true_when_all_nonzero() {
    i32_eq(
        "  v128.const i8x16 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16\n\
         \x20 i8x16.all_true",
        1,
    );
}

#[test]
fn i8x16_bitmask_gathers_high_bits() {
    i32_eq(
        // High bit set on lanes 0,2,15 (negative bytes) → bits 0,2,15.
        "  v128.const i8x16 -1 0 -128 0 0 0 0 0 0 0 0 0 0 0 0 -1\n\
         \x20 i8x16.bitmask",
        0b1000_0000_0000_0101,
    );
}

// ── float-RESULT ops verified across ALL lanes (lane-extracting helper) ──────
// f32x4.sqrt: a negative lane yields NaN while its neighbours stay finite —
// single-lane extraction of lane 0 would never see the NaN in lane 3.
#[test]
fn f32x4_sqrt_all_lanes_with_nan() {
    f32x4_eq(
        "  v128.const f32x4 4.0 9.0 16.0 -1.0\n  f32x4.sqrt",
        ["f32.const 2.0", "f32.const 3.0", "f32.const 4.0", "f32.const nan:canonical"],
    );
}

// f32x4.min: NaN propagates from EITHER operand, in a DIFFERENT lane each time.
#[test]
fn f32x4_min_nan_propagates_per_lane() {
    f32x4_eq(
        "  v128.const f32x4 nan 5.0 3.0 6.0\n\
         \x20 v128.const f32x4 1.0 nan 8.0 2.0\n\
         \x20 f32x4.min",
        // lane0 left-NaN, lane1 right-NaN, lane2 min(3,8)=3, lane3 min(6,2)=2
        ["f32.const nan:canonical", "f32.const nan:canonical", "f32.const 3.0", "f32.const 2.0"],
    );
}

#[test]
fn f32x4_max_nan_propagates_per_lane() {
    f32x4_eq(
        "  v128.const f32x4 nan 5.0 3.0 7.0\n\
         \x20 v128.const f32x4 1.0 nan 8.0 2.0\n\
         \x20 f32x4.max",
        ["f32.const nan:canonical", "f32.const nan:canonical", "f32.const 8.0", "f32.const 7.0"],
    );
}

// f32x4.pmin/pmax are the NON-propagating pseudo-ops: `b<a?b:a` / `a<b?b:a`, so
// a NaN operand makes the comparison false and the FIRST operand is returned.
#[test]
fn f32x4_pmin_returns_first_on_nan() {
    f32x4_eq(
        "  v128.const f32x4 3.0 8.0 nan 5.0\n\
         \x20 v128.const f32x4 8.0 3.0 2.0 nan\n\
         \x20 f32x4.pmin",
        // lane0 8<3?no→3, lane1 3<8?yes→3, lane2 2<nan?no→nan(a), lane3 nan<5?no→5(a)
        ["f32.const 3.0", "f32.const 3.0", "f32.const nan:canonical", "f32.const 5.0"],
    );
}

// Rounding modes differ per lane (sign + fraction) — check all four at once.
#[test]
fn f32x4_ceil_all_lanes() {
    f32x4_eq(
        "  v128.const f32x4 1.1 -1.1 2.9 -2.9\n  f32x4.ceil",
        ["f32.const 2.0", "f32.const -1.0", "f32.const 3.0", "f32.const -2.0"],
    );
}

#[test]
fn f32x4_floor_all_lanes() {
    f32x4_eq(
        "  v128.const f32x4 1.1 -1.1 2.9 -2.9\n  f32x4.floor",
        ["f32.const 1.0", "f32.const -2.0", "f32.const 2.0", "f32.const -3.0"],
    );
}

#[test]
fn f32x4_trunc_all_lanes() {
    f32x4_eq(
        "  v128.const f32x4 1.9 -1.9 2.1 -2.1\n  f32x4.trunc",
        ["f32.const 1.0", "f32.const -1.0", "f32.const 2.0", "f32.const -2.0"],
    );
}

// f32x4.nearest rounds ties to EVEN — 0.5→0, 1.5→2, 2.5→2, 3.5→4.
#[test]
fn f32x4_nearest_ties_to_even() {
    f32x4_eq(
        "  v128.const f32x4 0.5 1.5 2.5 3.5\n  f32x4.nearest",
        ["f32.const 0.0", "f32.const 2.0", "f32.const 2.0", "f32.const 4.0"],
    );
}

// abs/neg preserve magnitude/sign per lane, including a zero lane.
#[test]
fn f32x4_abs_all_lanes() {
    f32x4_eq(
        "  v128.const f32x4 -1.5 2.5 -3.5 0.0\n  f32x4.abs",
        ["f32.const 1.5", "f32.const 2.5", "f32.const 3.5", "f32.const 0.0"],
    );
}

#[test]
fn f32x4_neg_all_lanes() {
    f32x4_eq(
        "  v128.const f32x4 1.5 -2.5 3.5 -4.5\n  f32x4.neg",
        ["f32.const -1.5", "f32.const 2.5", "f32.const -3.5", "f32.const 4.5"],
    );
}

#[test]
fn f32x4_div_all_lanes() {
    f32x4_eq(
        "  v128.const f32x4 10.0 9.0 8.0 7.0\n\
         \x20 v128.const f32x4 2.0 3.0 4.0 7.0\n\
         \x20 f32x4.div",
        ["f32.const 5.0", "f32.const 3.0", "f32.const 2.0", "f32.const 1.0"],
    );
}

// int → float conversion, every lane (signed).
#[test]
fn f32x4_convert_i32x4_s_all_lanes() {
    f32x4_eq(
        "  v128.const i32x4 -5 0 100 1000\n  f32x4.convert_i32x4_s",
        ["f32.const -5.0", "f32.const 0.0", "f32.const 100.0", "f32.const 1000.0"],
    );
}

// ── f64x2 (2 lanes) ─────────────────────────────────────────────────────────
#[test]
fn f64x2_add_both_lanes() {
    f64x2_eq(
        "  v128.const f64x2 1.5 2.5\n  v128.const f64x2 0.5 0.5\n  f64x2.add",
        ["f64.const 2.0", "f64.const 3.0"],
    );
}

#[test]
fn f64x2_sqrt_with_nan_lane() {
    f64x2_eq(
        "  v128.const f64x2 16.0 -4.0\n  f64x2.sqrt",
        ["f64.const 4.0", "f64.const nan:canonical"],
    );
}

#[test]
fn f64x2_min_nan_propagates() {
    f64x2_eq(
        "  v128.const f64x2 nan 8.0\n  v128.const f64x2 5.0 3.0\n  f64x2.min",
        ["f64.const nan:canonical", "f64.const 3.0"],
    );
}

// f64x2.promote_low_f32x4 widens the LOW two f32 lanes to f64.
#[test]
fn f64x2_promote_low_f32x4_both_lanes() {
    f64x2_eq(
        "  v128.const f32x4 1.5 2.5 9.0 9.0\n  f64x2.promote_low_f32x4",
        ["f64.const 1.5", "f64.const 2.5"],
    );
}

// ── unsigned ops: pin BOTH the true/non-saturating AND false/floor results ──
// The per-shape suites only assert the `0` (unsigned-false / saturate-to-floor)
// case for these, which a constant-0 stub would pass. These add the other half.
//
// i8x16.lt_u: -1 as a byte is 0xFF (255) — unsigned it is NOT < small values,
// but small values ARE < it. Mix true and false lanes.
#[test]
fn i8x16_lt_u_true_and_false_lanes() {
    v128_eq(
        "  v128.const i8x16 0 255 16 128 1 1 1 1 1 1 1 1 1 1 1 1\n\
         \x20 v128.const i8x16 1 0 32 127 1 1 1 1 1 1 1 1 1 1 1 1\n\
         \x20 i8x16.lt_u",
        // 0<1 T, 255<0 F, 16<32 T, 128<127 F, then all 1<1 F
        "v128.const i8x16 -1 0 -1 0 0 0 0 0 0 0 0 0 0 0 0 0",
    );
}

#[test]
fn i32x4_lt_u_true_and_false_lanes() {
    v128_eq(
        // -1 lane = 0xFFFFFFFF (unsigned max), -2 = 0xFFFFFFFE.
        "  v128.const i32x4 1 -1 10 50\n\
         \x20 v128.const i32x4 2 -2 10 999\n\
         \x20 i32x4.lt_u",
        // 1<2 T, MAX<MAX-1 F, 10<10 F, 50<999 T
        "v128.const i32x4 -1 0 0 -1",
    );
}

// i8x16.sub_sat_u: unsigned saturating subtract — verify a NON-saturating lane
// (real difference) alongside the saturate-to-0 floor, per lane.
#[test]
fn i8x16_sub_sat_u_saturating_and_not() {
    v128_eq(
        "  v128.const i8x16 10 5 200 100 50 0 255 8 1 2 3 4 5 6 7 8\n\
         \x20 v128.const i8x16 3 10 50 100 0 5 1 8 1 2 3 4 5 6 7 8\n\
         \x20 i8x16.sub_sat_u",
        // 10-3=7, 5-10→0(floor), 200-50=150, 100-100=0, 50-0=50, 0-5→0, 255-1=254,
        // 8-8=0, then n-n=0
        "v128.const i8x16 7 0 150 0 50 0 254 0 0 0 0 0 0 0 0 0",
    );
}

#[test]
fn i16x8_sub_sat_u_saturating_and_not() {
    v128_eq(
        "  v128.const i16x8 1000 5 300 0 100 200 65535 8\n\
         \x20 v128.const i16x8 1 10 300 5 50 100 1 8\n\
         \x20 i16x8.sub_sat_u",
        // 1000-1=999, 5-10→0, 300-300=0, 0-5→0, 100-50=50, 200-100=100, 65535-1=65534, 8-8=0
        "v128.const i16x8 999 0 0 0 50 100 65534 0",
    );
}

// ── load*_zero: verify BOTH halves at once — the loaded lane AND the zero-fill.
// The per-shape suites check only ONE (load32_zero reads lane 1=0 but never the
// loaded lane 0; load64_zero the reverse), so a stub could pass. Assert the
// whole vector from memory.
#[test]
fn v128_load32_zero_loads_and_zeroes() {
    run_wast_asserts(
        "(module\n\
         \x20 (memory 1) (data (i32.const 0) \"\\07\\00\\00\\00\")\n\
         \x20 (func (export \"f\") (result v128)\n\
         \x20   i32.const 0 v128.load32_zero))\n\
         (assert_return (invoke \"f\") (v128.const i32x4 7 0 0 0))\n",
    )
    .unwrap_or_else(|e| panic!("load32_zero whole-vector assert failed: {e}"));
}

#[test]
fn v128_load64_zero_loads_and_zeroes() {
    run_wast_asserts(
        "(module\n\
         \x20 (memory 1) (data (i32.const 0) \"\\09\\00\\00\\00\\00\\00\\00\\00\")\n\
         \x20 (func (export \"f\") (result v128)\n\
         \x20   i32.const 0 v128.load64_zero))\n\
         (assert_return (invoke \"f\") (v128.const i64x2 9 0))\n",
    )
    .unwrap_or_else(|e| panic!("load64_zero whole-vector assert failed: {e}"));
}
