;; vybe-test: wast/wat_simd_full_vector/float_rounding_conversion_and_pminmax
;; vybe-test-mode: run
;;
;; The float-lane SIMD group: rounding (`ceil`/`floor`/`trunc`/`nearest`), the
;; pseudo-min/max pair, width conversion (`convert_*`, `demote`, `promote`), and
;; the saturating float→int truncations. Each of these occurred once in the
;; corpus, on an operand where every one of them agrees.
;;
;; What separates them:
;;
;;   * rounding: a value at .5 with BOTH signs. `ceil`, `floor` and `trunc`
;;     differ on -1.5 alone, and `nearest` is ties-to-EVEN, so 0.5 → 0 and
;;     2.5 → 2 — not "round half up", which is the reading that passes every
;;     .25/.75 test ever written.
;;   * `pmin`/`pmax`: defined as a bare comparison (`pmin(x,y) = y<x ? y : x`),
;;     so they are NOT `min`/`max`. With NaN as the second operand the first is
;;     returned, and on ±0 they return the SECOND-listed operand rather than the
;;     signed minimum.
;;   * signed vs unsigned convert: a NEGATIVE lane. -1 converts to -1.0 signed
;;     and 4294967295.0 unsigned.
;;   * `convert_i32x4_u` at 16777217, which is not representable in f32 and
;;     rounds to even.
;;   * `trunc_sat_*_zero`: operands beyond the integer range in both directions,
;;     plus the upper lanes that the instruction must ZERO.
;;
;; Spec-format so `wasmtime wast` arbitrates every expectation.

(module
  ;; ── f64x2 rounding: the four modes differ on -1.5 ────────────────────
  (func (export "f64x2_ceil") (result v128)
    (f64x2.ceil (v128.const f64x2 -1.5 2.5)))
  (func (export "f64x2_floor") (result v128)
    (f64x2.floor (v128.const f64x2 -1.5 2.5)))
  (func (export "f64x2_trunc") (result v128)
    (f64x2.trunc (v128.const f64x2 -1.5 2.5)))
  (func (export "f64x2_nearest") (result v128)
    (f64x2.nearest (v128.const f64x2 -1.5 2.5)))
  ;; Ties-to-even, where "round half away from zero" gives 1.0 and 3.0.
  (func (export "f64x2_nearest_ties_even") (result v128)
    (f64x2.nearest (v128.const f64x2 0.5 2.5)))
  (func (export "f64x2_nearest_ties_even_neg") (result v128)
    (f64x2.nearest (v128.const f64x2 -0.5 -2.5)))

  ;; ── pmin / pmax are comparisons, not min / max ───────────────────────
  (func (export "f64x2_pmin") (result v128)
    (f64x2.pmin (v128.const f64x2 1.0 4.0) (v128.const f64x2 2.0 3.0)))
  (func (export "f64x2_pmax") (result v128)
    (f64x2.pmax (v128.const f64x2 1.0 4.0) (v128.const f64x2 2.0 3.0)))
  ;; NaN as the SECOND operand: the comparison is false, so the FIRST is
  ;; returned by both. And on +0/-0 the comparison is also false.
  (func (export "f64x2_pmin_nan_second") (result v128)
    (f64x2.pmin (v128.const f64x2 3.0 -0.0) (v128.const f64x2 nan 0.0)))
  (func (export "f64x2_pmax_nan_second") (result v128)
    (f64x2.pmax (v128.const f64x2 3.0 -0.0) (v128.const f64x2 nan 0.0)))

  ;; ── int → float conversion: the negative lane is the whole test ──────
  (func (export "f64x2_convert_low_s") (result v128)
    (f64x2.convert_low_i32x4_s (v128.const i32x4 -1 2147483647 99 99)))
  (func (export "f64x2_convert_low_u") (result v128)
    (f64x2.convert_low_i32x4_u (v128.const i32x4 -1 2147483647 99 99)))
  (func (export "f32x4_convert_s") (result v128)
    (f32x4.convert_i32x4_s (v128.const i32x4 -1 1 16777217 0)))
  (func (export "f32x4_convert_u") (result v128)
    (f32x4.convert_i32x4_u (v128.const i32x4 -1 1 16777217 0)))

  ;; ── width conversion: the upper lanes are zeroed / dropped ───────────
  (func (export "f32x4_demote") (result v128)
    (f32x4.demote_f64x2_zero (v128.const f64x2 1.5 -2.5)))
  (func (export "f64x2_promote_low") (result v128)
    (f64x2.promote_low_f32x4 (v128.const f32x4 1.5 -2.5 99.0 99.0)))

  ;; ── saturating float → int, with the upper lanes zeroed ─────────────
  (func (export "trunc_sat_f64x2_s_zero") (result v128)
    (i32x4.trunc_sat_f64x2_s_zero (v128.const f64x2 1e300 -1e300)))
  (func (export "trunc_sat_f64x2_u_zero") (result v128)
    (i32x4.trunc_sat_f64x2_u_zero (v128.const f64x2 -0.5 1e300)))
  ;; -0.5 truncates toward zero and lands INSIDE the unsigned range, so it is
  ;; 0 rather than a saturation; -1.5 is below the range and clamps.
  (func (export "trunc_sat_f64x2_u_zero_below") (result v128)
    (i32x4.trunc_sat_f64x2_u_zero (v128.const f64x2 -1.5 4294967295.0)))

  ;; ── the remaining thin integer lanes ─────────────────────────────────
  (func (export "i16x8_all_true_yes") (result i32)
    (i16x8.all_true (v128.const i16x8 1 -1 32767 -32768 2 3 4 5)))
  (func (export "i16x8_all_true_no") (result i32)
    (i16x8.all_true (v128.const i16x8 1 -1 32767 0 2 3 4 5)))
  (func (export "i16x8_max_s") (result v128)
    (i16x8.max_s (v128.const i16x8 -1 5 -32768 3 0 0 0 0)
                 (v128.const i16x8 1 -5 32767 3 0 0 0 0)))
  (func (export "i16x8_max_u") (result v128)
    (i16x8.max_u (v128.const i16x8 -1 5 -32768 3 0 0 0 0)
                 (v128.const i16x8 1 -5 32767 3 0 0 0 0)))
  (func (export "i8x16_max_u") (result v128)
    (i8x16.max_u (v128.const i8x16 -1 5 -128 3 0 0 0 0 0 0 0 0 0 0 0 0)
                 (v128.const i8x16 1 -5 127 3 0 0 0 0 0 0 0 0 0 0 0 0)))
  (func (export "i16x8_replace_lane") (result v128)
    (i16x8.replace_lane 3 (v128.const i16x8 0 0 0 0 0 0 0 0) (i32.const 1234)))
  ;; A replaced lane takes the LOW 16 bits of the i32 operand.
  (func (export "i16x8_replace_lane_truncates") (result v128)
    (i16x8.replace_lane 0 (v128.const i16x8 0 0 0 0 0 0 0 0) (i32.const 65538)))
)

;; ── rounding ─────────────────────────────────────────────────────────
(assert_return (invoke "f64x2_ceil")   (v128.const f64x2 -1.0 3.0))
(assert_return (invoke "f64x2_floor")  (v128.const f64x2 -2.0 2.0))
(assert_return (invoke "f64x2_trunc")  (v128.const f64x2 -1.0 2.0))
(assert_return (invoke "f64x2_nearest") (v128.const f64x2 -2.0 2.0))
(assert_return (invoke "f64x2_nearest_ties_even") (v128.const f64x2 0.0 2.0))
(assert_return (invoke "f64x2_nearest_ties_even_neg") (v128.const f64x2 -0.0 -2.0))

;; ── pmin / pmax ──────────────────────────────────────────────────────
(assert_return (invoke "f64x2_pmin") (v128.const f64x2 1.0 3.0))
(assert_return (invoke "f64x2_pmax") (v128.const f64x2 2.0 4.0))
(assert_return (invoke "f64x2_pmin_nan_second") (v128.const f64x2 3.0 -0.0))
(assert_return (invoke "f64x2_pmax_nan_second") (v128.const f64x2 3.0 -0.0))

;; ── conversion ───────────────────────────────────────────────────────
(assert_return (invoke "f64x2_convert_low_s") (v128.const f64x2 -1.0 2147483647.0))
(assert_return (invoke "f64x2_convert_low_u") (v128.const f64x2 4294967295.0 2147483647.0))
(assert_return (invoke "f32x4_convert_s") (v128.const f32x4 -1.0 1.0 16777216.0 0.0))
(assert_return (invoke "f32x4_convert_u") (v128.const f32x4 4294967296.0 1.0 16777216.0 0.0))

(assert_return (invoke "f32x4_demote") (v128.const f32x4 1.5 -2.5 0.0 0.0))
(assert_return (invoke "f64x2_promote_low") (v128.const f64x2 1.5 -2.5))

;; ── saturating truncation ────────────────────────────────────────────
(assert_return (invoke "trunc_sat_f64x2_s_zero") (v128.const i32x4 2147483647 -2147483648 0 0))
(assert_return (invoke "trunc_sat_f64x2_u_zero") (v128.const i32x4 0 -1 0 0))
(assert_return (invoke "trunc_sat_f64x2_u_zero_below") (v128.const i32x4 0 -1 0 0))

;; ── integer lanes ────────────────────────────────────────────────────
(assert_return (invoke "i16x8_all_true_yes") (i32.const 1))
(assert_return (invoke "i16x8_all_true_no") (i32.const 0))
(assert_return (invoke "i16x8_max_s") (v128.const i16x8 1 5 32767 3 0 0 0 0))
(assert_return (invoke "i16x8_max_u") (v128.const i16x8 -1 -5 -32768 3 0 0 0 0))
(assert_return (invoke "i8x16_max_u")
  (v128.const i8x16 -1 -5 -128 3 0 0 0 0 0 0 0 0 0 0 0 0))
(assert_return (invoke "i16x8_replace_lane") (v128.const i16x8 0 0 0 1234 0 0 0 0))
(assert_return (invoke "i16x8_replace_lane_truncates") (v128.const i16x8 2 0 0 0 0 0 0 0))
