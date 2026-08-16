;; vybe-test: wast/wat_simd_full_vector/lane_comparisons_signed_vs_unsigned
;; vybe-test-mode: run
;;
;; Every SIMD comparison, at the operands that make it DIFFERENT from its
;; neighbours. Forty of these mnemonics occurred exactly once in the whole
;; run corpus, each in a file asserting one ordinary true case — which
;; `i8x16.eq` also satisfies, so nothing in the corpus could tell `lt_s` from
;; `lt_u`, or either from `gt_s`.
;;
;; The operand pattern is chosen so signed and unsigned readings DISAGREE on
;; every lane that is not equal:
;;
;;   lane 0:  0 vs 0            — the equal case
;;   lane 1:  1 vs -1           — unsigned -1 is the LARGEST value, signed the
;;                                smallest, so lt_u and lt_s disagree
;;   lane 2: -1 vs 1            — the mirror
;;   lane 3: MIN vs MAX         — the extreme where the two orderings invert
;;
;; So `lt_s` and `gt_u` must produce the SAME mask, `lt_u` and `gt_s` the same
;; mask, and an implementation that computes any of them by the wrong signedness
;; swaps two of the four lanes rather than failing outright.
;;
;; Floats add the two facts integers cannot express: every comparison with NaN
;; is false (including `ne`, which is true — the negation of `eq`, not an
;; ordered predicate), and `+0.0 == -0.0`.
;;
;; A true lane is all-ones (-1), a false lane 0; the result of a float
;; comparison is an INTEGER mask, not a float.
;;
;; Spec-format so `wasmtime wast` arbitrates every expectation.

(module
  ;; ── i8x16: 0/1/-1/-128 vs 0/-1/1/127, repeated over all 16 lanes ─────
  (func (export "i8x16_eq") (result v128)
    (i8x16.eq (v128.const i8x16 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128)
              (v128.const i8x16 0 -1 1 127 0 -1 1 127 0 -1 1 127 0 -1 1 127)))
  (func (export "i8x16_ne") (result v128)
    (i8x16.ne (v128.const i8x16 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128)
              (v128.const i8x16 0 -1 1 127 0 -1 1 127 0 -1 1 127 0 -1 1 127)))
  (func (export "i8x16_lt_s") (result v128)
    (i8x16.lt_s (v128.const i8x16 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128)
                (v128.const i8x16 0 -1 1 127 0 -1 1 127 0 -1 1 127 0 -1 1 127)))
  (func (export "i8x16_lt_u") (result v128)
    (i8x16.lt_u (v128.const i8x16 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128)
                (v128.const i8x16 0 -1 1 127 0 -1 1 127 0 -1 1 127 0 -1 1 127)))
  (func (export "i8x16_gt_s") (result v128)
    (i8x16.gt_s (v128.const i8x16 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128)
                (v128.const i8x16 0 -1 1 127 0 -1 1 127 0 -1 1 127 0 -1 1 127)))
  (func (export "i8x16_gt_u") (result v128)
    (i8x16.gt_u (v128.const i8x16 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128)
                (v128.const i8x16 0 -1 1 127 0 -1 1 127 0 -1 1 127 0 -1 1 127)))
  (func (export "i8x16_le_s") (result v128)
    (i8x16.le_s (v128.const i8x16 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128)
                (v128.const i8x16 0 -1 1 127 0 -1 1 127 0 -1 1 127 0 -1 1 127)))
  (func (export "i8x16_le_u") (result v128)
    (i8x16.le_u (v128.const i8x16 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128)
                (v128.const i8x16 0 -1 1 127 0 -1 1 127 0 -1 1 127 0 -1 1 127)))
  (func (export "i8x16_ge_s") (result v128)
    (i8x16.ge_s (v128.const i8x16 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128)
                (v128.const i8x16 0 -1 1 127 0 -1 1 127 0 -1 1 127 0 -1 1 127)))
  (func (export "i8x16_ge_u") (result v128)
    (i8x16.ge_u (v128.const i8x16 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128 0 1 -1 -128)
                (v128.const i8x16 0 -1 1 127 0 -1 1 127 0 -1 1 127 0 -1 1 127)))

  ;; ── i16x8: same shape at 16-bit MIN/MAX ──────────────────────────────
  (func (export "i16x8_eq") (result v128)
    (i16x8.eq (v128.const i16x8 0 1 -1 -32768 0 1 -1 -32768)
              (v128.const i16x8 0 -1 1 32767 0 -1 1 32767)))
  (func (export "i16x8_ne") (result v128)
    (i16x8.ne (v128.const i16x8 0 1 -1 -32768 0 1 -1 -32768)
              (v128.const i16x8 0 -1 1 32767 0 -1 1 32767)))
  (func (export "i16x8_lt_s") (result v128)
    (i16x8.lt_s (v128.const i16x8 0 1 -1 -32768 0 1 -1 -32768)
                (v128.const i16x8 0 -1 1 32767 0 -1 1 32767)))
  (func (export "i16x8_lt_u") (result v128)
    (i16x8.lt_u (v128.const i16x8 0 1 -1 -32768 0 1 -1 -32768)
                (v128.const i16x8 0 -1 1 32767 0 -1 1 32767)))
  (func (export "i16x8_gt_s") (result v128)
    (i16x8.gt_s (v128.const i16x8 0 1 -1 -32768 0 1 -1 -32768)
                (v128.const i16x8 0 -1 1 32767 0 -1 1 32767)))
  (func (export "i16x8_gt_u") (result v128)
    (i16x8.gt_u (v128.const i16x8 0 1 -1 -32768 0 1 -1 -32768)
                (v128.const i16x8 0 -1 1 32767 0 -1 1 32767)))
  (func (export "i16x8_le_s") (result v128)
    (i16x8.le_s (v128.const i16x8 0 1 -1 -32768 0 1 -1 -32768)
                (v128.const i16x8 0 -1 1 32767 0 -1 1 32767)))
  (func (export "i16x8_le_u") (result v128)
    (i16x8.le_u (v128.const i16x8 0 1 -1 -32768 0 1 -1 -32768)
                (v128.const i16x8 0 -1 1 32767 0 -1 1 32767)))
  (func (export "i16x8_ge_s") (result v128)
    (i16x8.ge_s (v128.const i16x8 0 1 -1 -32768 0 1 -1 -32768)
                (v128.const i16x8 0 -1 1 32767 0 -1 1 32767)))
  (func (export "i16x8_ge_u") (result v128)
    (i16x8.ge_u (v128.const i16x8 0 1 -1 -32768 0 1 -1 -32768)
                (v128.const i16x8 0 -1 1 32767 0 -1 1 32767)))

  ;; ── i32x4: same shape at 32-bit MIN/MAX ──────────────────────────────
  (func (export "i32x4_eq") (result v128)
    (i32x4.eq (v128.const i32x4 0 1 -1 -2147483648)
              (v128.const i32x4 0 -1 1 2147483647)))
  (func (export "i32x4_ne") (result v128)
    (i32x4.ne (v128.const i32x4 0 1 -1 -2147483648)
              (v128.const i32x4 0 -1 1 2147483647)))
  (func (export "i32x4_lt_s") (result v128)
    (i32x4.lt_s (v128.const i32x4 0 1 -1 -2147483648)
                (v128.const i32x4 0 -1 1 2147483647)))
  (func (export "i32x4_lt_u") (result v128)
    (i32x4.lt_u (v128.const i32x4 0 1 -1 -2147483648)
                (v128.const i32x4 0 -1 1 2147483647)))
  (func (export "i32x4_gt_s") (result v128)
    (i32x4.gt_s (v128.const i32x4 0 1 -1 -2147483648)
                (v128.const i32x4 0 -1 1 2147483647)))
  (func (export "i32x4_gt_u") (result v128)
    (i32x4.gt_u (v128.const i32x4 0 1 -1 -2147483648)
                (v128.const i32x4 0 -1 1 2147483647)))
  (func (export "i32x4_le_s") (result v128)
    (i32x4.le_s (v128.const i32x4 0 1 -1 -2147483648)
                (v128.const i32x4 0 -1 1 2147483647)))
  (func (export "i32x4_le_u") (result v128)
    (i32x4.le_u (v128.const i32x4 0 1 -1 -2147483648)
                (v128.const i32x4 0 -1 1 2147483647)))
  (func (export "i32x4_ge_s") (result v128)
    (i32x4.ge_s (v128.const i32x4 0 1 -1 -2147483648)
                (v128.const i32x4 0 -1 1 2147483647)))
  (func (export "i32x4_ge_u") (result v128)
    (i32x4.ge_u (v128.const i32x4 0 1 -1 -2147483648)
                (v128.const i32x4 0 -1 1 2147483647)))

  ;; ── i64x2: the spec gives it SIGNED comparisons only ─────────────────
  (func (export "i64x2_eq") (result v128)
    (i64x2.eq (v128.const i64x2 0 -1) (v128.const i64x2 0 1)))
  (func (export "i64x2_ne") (result v128)
    (i64x2.ne (v128.const i64x2 0 -1) (v128.const i64x2 0 1)))
  (func (export "i64x2_lt_s") (result v128)
    (i64x2.lt_s (v128.const i64x2 0 -1) (v128.const i64x2 0 1)))
  (func (export "i64x2_gt_s") (result v128)
    (i64x2.gt_s (v128.const i64x2 0 -1) (v128.const i64x2 0 1)))
  (func (export "i64x2_le_s") (result v128)
    (i64x2.le_s (v128.const i64x2 0 -1) (v128.const i64x2 0 1)))
  (func (export "i64x2_ge_s") (result v128)
    (i64x2.ge_s (v128.const i64x2 0 -1) (v128.const i64x2 0 1)))
  ;; The 64-bit extreme, where a 32-bit-wide comparison would read the wrong
  ;; half and call the two lanes equal.
  (func (export "i64x2_lt_s_wide") (result v128)
    (i64x2.lt_s (v128.const i64x2 -9223372036854775808 9223372036854775807)
                (v128.const i64x2 9223372036854775807 -9223372036854775808)))

  ;; ── f32x4: NaN is unordered, and +0.0 equals -0.0 ────────────────────
  (func (export "f32x4_eq_nan_zero") (result v128)
    (f32x4.eq (v128.const f32x4 0.0 -0.0 nan 1.0)
              (v128.const f32x4 -0.0 0.0 1.0 nan)))
  (func (export "f32x4_ne_nan_zero") (result v128)
    (f32x4.ne (v128.const f32x4 0.0 -0.0 nan 1.0)
              (v128.const f32x4 -0.0 0.0 1.0 nan)))
  (func (export "f32x4_lt_nan_zero") (result v128)
    (f32x4.lt (v128.const f32x4 0.0 -0.0 nan 1.0)
              (v128.const f32x4 -0.0 0.0 1.0 nan)))
  (func (export "f32x4_gt_nan_zero") (result v128)
    (f32x4.gt (v128.const f32x4 0.0 -0.0 nan 1.0)
              (v128.const f32x4 -0.0 0.0 1.0 nan)))
  (func (export "f32x4_le_nan_zero") (result v128)
    (f32x4.le (v128.const f32x4 0.0 -0.0 nan 1.0)
              (v128.const f32x4 -0.0 0.0 1.0 nan)))
  (func (export "f32x4_ge_nan_zero") (result v128)
    (f32x4.ge (v128.const f32x4 0.0 -0.0 nan 1.0)
              (v128.const f32x4 -0.0 0.0 1.0 nan)))

  ;; Ordering on ordinary values, including across zero.
  (func (export "f32x4_lt_ordered") (result v128)
    (f32x4.lt (v128.const f32x4 1.0 2.0 -1.0 -2.0)
              (v128.const f32x4 2.0 1.0 -2.0 -1.0)))
  (func (export "f32x4_gt_ordered") (result v128)
    (f32x4.gt (v128.const f32x4 1.0 2.0 -1.0 -2.0)
              (v128.const f32x4 2.0 1.0 -2.0 -1.0)))
  (func (export "f32x4_le_ordered") (result v128)
    (f32x4.le (v128.const f32x4 1.0 2.0 -1.0 -2.0)
              (v128.const f32x4 2.0 1.0 -2.0 -1.0)))
  (func (export "f32x4_ge_ordered") (result v128)
    (f32x4.ge (v128.const f32x4 1.0 2.0 -1.0 -2.0)
              (v128.const f32x4 2.0 1.0 -2.0 -1.0)))

  ;; ── f64x2: the same two facts at double width ────────────────────────
  (func (export "f64x2_eq_nan_zero") (result v128)
    (f64x2.eq (v128.const f64x2 0.0 nan) (v128.const f64x2 -0.0 1.0)))
  (func (export "f64x2_ne_nan_zero") (result v128)
    (f64x2.ne (v128.const f64x2 0.0 nan) (v128.const f64x2 -0.0 1.0)))
  (func (export "f64x2_lt_nan_zero") (result v128)
    (f64x2.lt (v128.const f64x2 0.0 nan) (v128.const f64x2 -0.0 1.0)))
  (func (export "f64x2_gt_nan_zero") (result v128)
    (f64x2.gt (v128.const f64x2 0.0 nan) (v128.const f64x2 -0.0 1.0)))
  (func (export "f64x2_le_ordered") (result v128)
    (f64x2.le (v128.const f64x2 1.0 -1.0) (v128.const f64x2 2.0 -2.0)))
  (func (export "f64x2_ge_ordered") (result v128)
    (f64x2.ge (v128.const f64x2 1.0 -1.0) (v128.const f64x2 2.0 -2.0)))
)

;; i8x16 — lane pattern (equal, 1v-1, -1v1, MINvMAX) repeated four times.
(assert_return (invoke "i8x16_eq")   (v128.const i8x16 -1 0 0 0 -1 0 0 0 -1 0 0 0 -1 0 0 0))
(assert_return (invoke "i8x16_ne")   (v128.const i8x16 0 -1 -1 -1 0 -1 -1 -1 0 -1 -1 -1 0 -1 -1 -1))
(assert_return (invoke "i8x16_lt_s") (v128.const i8x16 0 0 -1 -1 0 0 -1 -1 0 0 -1 -1 0 0 -1 -1))
(assert_return (invoke "i8x16_lt_u") (v128.const i8x16 0 -1 0 0 0 -1 0 0 0 -1 0 0 0 -1 0 0))
(assert_return (invoke "i8x16_gt_s") (v128.const i8x16 0 -1 0 0 0 -1 0 0 0 -1 0 0 0 -1 0 0))
(assert_return (invoke "i8x16_gt_u") (v128.const i8x16 0 0 -1 -1 0 0 -1 -1 0 0 -1 -1 0 0 -1 -1))
(assert_return (invoke "i8x16_le_s") (v128.const i8x16 -1 0 -1 -1 -1 0 -1 -1 -1 0 -1 -1 -1 0 -1 -1))
(assert_return (invoke "i8x16_le_u") (v128.const i8x16 -1 -1 0 0 -1 -1 0 0 -1 -1 0 0 -1 -1 0 0))
(assert_return (invoke "i8x16_ge_s") (v128.const i8x16 -1 -1 0 0 -1 -1 0 0 -1 -1 0 0 -1 -1 0 0))
(assert_return (invoke "i8x16_ge_u") (v128.const i8x16 -1 0 -1 -1 -1 0 -1 -1 -1 0 -1 -1 -1 0 -1 -1))

;; i16x8 — the same four-lane pattern, twice.
(assert_return (invoke "i16x8_eq")   (v128.const i16x8 -1 0 0 0 -1 0 0 0))
(assert_return (invoke "i16x8_ne")   (v128.const i16x8 0 -1 -1 -1 0 -1 -1 -1))
(assert_return (invoke "i16x8_lt_s") (v128.const i16x8 0 0 -1 -1 0 0 -1 -1))
(assert_return (invoke "i16x8_lt_u") (v128.const i16x8 0 -1 0 0 0 -1 0 0))
(assert_return (invoke "i16x8_gt_s") (v128.const i16x8 0 -1 0 0 0 -1 0 0))
(assert_return (invoke "i16x8_gt_u") (v128.const i16x8 0 0 -1 -1 0 0 -1 -1))
(assert_return (invoke "i16x8_le_s") (v128.const i16x8 -1 0 -1 -1 -1 0 -1 -1))
(assert_return (invoke "i16x8_le_u") (v128.const i16x8 -1 -1 0 0 -1 -1 0 0))
(assert_return (invoke "i16x8_ge_s") (v128.const i16x8 -1 -1 0 0 -1 -1 0 0))
(assert_return (invoke "i16x8_ge_u") (v128.const i16x8 -1 0 -1 -1 -1 0 -1 -1))

;; i32x4 — one pass of the pattern.
(assert_return (invoke "i32x4_eq")   (v128.const i32x4 -1 0 0 0))
(assert_return (invoke "i32x4_ne")   (v128.const i32x4 0 -1 -1 -1))
(assert_return (invoke "i32x4_lt_s") (v128.const i32x4 0 0 -1 -1))
(assert_return (invoke "i32x4_lt_u") (v128.const i32x4 0 -1 0 0))
(assert_return (invoke "i32x4_gt_s") (v128.const i32x4 0 -1 0 0))
(assert_return (invoke "i32x4_gt_u") (v128.const i32x4 0 0 -1 -1))
(assert_return (invoke "i32x4_le_s") (v128.const i32x4 -1 0 -1 -1))
(assert_return (invoke "i32x4_le_u") (v128.const i32x4 -1 -1 0 0))
(assert_return (invoke "i32x4_ge_s") (v128.const i32x4 -1 -1 0 0))
(assert_return (invoke "i32x4_ge_u") (v128.const i32x4 -1 0 -1 -1))

;; i64x2 — lanes (0 vs 0) and (-1 vs 1).
(assert_return (invoke "i64x2_eq")   (v128.const i64x2 -1 0))
(assert_return (invoke "i64x2_ne")   (v128.const i64x2 0 -1))
(assert_return (invoke "i64x2_lt_s") (v128.const i64x2 0 -1))
(assert_return (invoke "i64x2_gt_s") (v128.const i64x2 0 0))
(assert_return (invoke "i64x2_le_s") (v128.const i64x2 -1 -1))
(assert_return (invoke "i64x2_ge_s") (v128.const i64x2 -1 0))
(assert_return (invoke "i64x2_lt_s_wide") (v128.const i64x2 -1 0))

;; f32x4 — NaN in lanes 2 and 3; lanes 0 and 1 are +0.0 against -0.0.
(assert_return (invoke "f32x4_eq_nan_zero") (v128.const i32x4 -1 -1 0 0))
(assert_return (invoke "f32x4_ne_nan_zero") (v128.const i32x4 0 0 -1 -1))
(assert_return (invoke "f32x4_lt_nan_zero") (v128.const i32x4 0 0 0 0))
(assert_return (invoke "f32x4_gt_nan_zero") (v128.const i32x4 0 0 0 0))
(assert_return (invoke "f32x4_le_nan_zero") (v128.const i32x4 -1 -1 0 0))
(assert_return (invoke "f32x4_ge_nan_zero") (v128.const i32x4 -1 -1 0 0))

(assert_return (invoke "f32x4_lt_ordered") (v128.const i32x4 -1 0 0 -1))
(assert_return (invoke "f32x4_gt_ordered") (v128.const i32x4 0 -1 -1 0))
(assert_return (invoke "f32x4_le_ordered") (v128.const i32x4 -1 0 0 -1))
(assert_return (invoke "f32x4_ge_ordered") (v128.const i32x4 0 -1 -1 0))

;; f64x2 — lane 0 is +0.0 vs -0.0, lane 1 involves NaN.
(assert_return (invoke "f64x2_eq_nan_zero") (v128.const i64x2 -1 0))
(assert_return (invoke "f64x2_ne_nan_zero") (v128.const i64x2 0 -1))
(assert_return (invoke "f64x2_lt_nan_zero") (v128.const i64x2 0 0))
(assert_return (invoke "f64x2_gt_nan_zero") (v128.const i64x2 0 0))
(assert_return (invoke "f64x2_le_ordered") (v128.const i64x2 -1 0))
(assert_return (invoke "f64x2_ge_ordered") (v128.const i64x2 0 -1))
