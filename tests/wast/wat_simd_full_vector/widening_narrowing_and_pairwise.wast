;; vybe-test: wast/wat_simd_full_vector/widening_narrowing_and_pairwise
;; vybe-test-mode: run
;;
;; The width-changing SIMD family: `extend_*`, `extmul_*`, `extadd_pairwise_*`,
;; `narrow_*`, and `q15mulr_sat_s`. Thirty-odd of these occurred once each in the
;; corpus, always on small positive operands — where signed and unsigned
;; extension give the SAME answer and no saturation is reached, so nothing
;; distinguished `_s` from `_u` or a saturating narrow from a truncating one.
;;
;; What actually separates them:
;;
;;   * extension: a NEGATIVE source lane. Sign-extended -1 stays -1; zero-
;;     extended it becomes 255 / 65535 / 4294967295.
;;   * low vs high: the two halves carry different values here, so reading the
;;     wrong half is a wrong answer rather than the same answer.
;;   * extmul: products that overflow the SOURCE width but fit the destination
;;     (-128 * -128 = 16384), which is the whole reason the instruction exists.
;;   * narrow: operands outside the destination range on BOTH sides, so the
;;     result shows clamping rather than truncation — and the unsigned form
;;     clamps a negative lane to 0, not to its two's-complement bits.
;;   * q15mulr: the fixed-point rounding term and the one product that
;;     saturates (-32768 * -32768).
;;
;; Spec-format so `wasmtime wast` arbitrates every expectation.

(module
  ;; ── i16x8.extend_*_i8x16_{s,u} ───────────────────────────────────────
  ;; low half and high half hold different values, so the two are not
  ;; interchangeable.
  (func (export "i16x8_extend_low_s") (result v128)
    (i16x8.extend_low_i8x16_s
      (v128.const i8x16 0 1 -1 -128 127 -2 2 -127  3 -3 -128 127 0 -1 1 -4)))
  (func (export "i16x8_extend_low_u") (result v128)
    (i16x8.extend_low_i8x16_u
      (v128.const i8x16 0 1 -1 -128 127 -2 2 -127  3 -3 -128 127 0 -1 1 -4)))
  (func (export "i16x8_extend_high_s") (result v128)
    (i16x8.extend_high_i8x16_s
      (v128.const i8x16 0 1 -1 -128 127 -2 2 -127  3 -3 -128 127 0 -1 1 -4)))
  (func (export "i16x8_extend_high_u") (result v128)
    (i16x8.extend_high_i8x16_u
      (v128.const i8x16 0 1 -1 -128 127 -2 2 -127  3 -3 -128 127 0 -1 1 -4)))

  ;; ── i32x4.extend_*_i16x8_{s,u} ───────────────────────────────────────
  (func (export "i32x4_extend_low_s") (result v128)
    (i32x4.extend_low_i16x8_s
      (v128.const i16x8 0 1 -1 -32768  32767 -2 2 -32767)))
  (func (export "i32x4_extend_low_u") (result v128)
    (i32x4.extend_low_i16x8_u
      (v128.const i16x8 0 1 -1 -32768  32767 -2 2 -32767)))
  (func (export "i32x4_extend_high_s") (result v128)
    (i32x4.extend_high_i16x8_s
      (v128.const i16x8 0 1 -1 -32768  32767 -2 2 -32767)))
  (func (export "i32x4_extend_high_u") (result v128)
    (i32x4.extend_high_i16x8_u
      (v128.const i16x8 0 1 -1 -32768  32767 -2 2 -32767)))

  ;; ── i64x2.extend_*_i32x4_{s,u} ───────────────────────────────────────
  (func (export "i64x2_extend_low_s") (result v128)
    (i64x2.extend_low_i32x4_s (v128.const i32x4 -1 2147483647 -2147483648 1)))
  (func (export "i64x2_extend_low_u") (result v128)
    (i64x2.extend_low_i32x4_u (v128.const i32x4 -1 2147483647 -2147483648 1)))
  (func (export "i64x2_extend_high_s") (result v128)
    (i64x2.extend_high_i32x4_s (v128.const i32x4 -1 2147483647 -2147483648 1)))
  (func (export "i64x2_extend_high_u") (result v128)
    (i64x2.extend_high_i32x4_u (v128.const i32x4 -1 2147483647 -2147483648 1)))

  ;; ── i16x8.extmul_*_i8x16_{s,u} ───────────────────────────────────────
  ;; -128 * -128 = 16384 overflows i8 and fits i16 — the point of extmul.
  (func (export "i16x8_extmul_low_s") (result v128)
    (i16x8.extmul_low_i8x16_s
      (v128.const i8x16 2 -1 -128 127 -2 1 -3 3   4 -4 -128 127 5 -5 6 -6)
      (v128.const i8x16 3 -1 -1 2 -2 -1 -1 -1     2 2 -128 127 -1 -1 3 3)))
  (func (export "i16x8_extmul_low_u") (result v128)
    (i16x8.extmul_low_i8x16_u
      (v128.const i8x16 2 -1 -128 127 -2 1 -3 3   4 -4 -128 127 5 -5 6 -6)
      (v128.const i8x16 3 -1 -1 2 -2 -1 -1 -1     2 2 -128 127 -1 -1 3 3)))
  (func (export "i16x8_extmul_high_s") (result v128)
    (i16x8.extmul_high_i8x16_s
      (v128.const i8x16 2 -1 -128 127 -2 1 -3 3   4 -4 -128 127 5 -5 6 -6)
      (v128.const i8x16 3 -1 -1 2 -2 -1 -1 -1     2 2 -128 127 -1 -1 3 3)))
  (func (export "i16x8_extmul_high_u") (result v128)
    (i16x8.extmul_high_i8x16_u
      (v128.const i8x16 2 -1 -128 127 -2 1 -3 3   4 -4 -128 127 5 -5 6 -6)
      (v128.const i8x16 3 -1 -1 2 -2 -1 -1 -1     2 2 -128 127 -1 -1 3 3)))

  ;; ── i32x4.extmul_*_i16x8_{s,u} ───────────────────────────────────────
  (func (export "i32x4_extmul_low_s") (result v128)
    (i32x4.extmul_low_i16x8_s
      (v128.const i16x8 -1 -32768 3 -4   7 -8 32767 -32768)
      (v128.const i16x8 -1 -32768 -5 6   -9 10 32767 32767)))
  (func (export "i32x4_extmul_low_u") (result v128)
    (i32x4.extmul_low_i16x8_u
      (v128.const i16x8 -1 -32768 3 -4   7 -8 32767 -32768)
      (v128.const i16x8 -1 -32768 -5 6   -9 10 32767 32767)))
  (func (export "i32x4_extmul_high_s") (result v128)
    (i32x4.extmul_high_i16x8_s
      (v128.const i16x8 -1 -32768 3 -4   7 -8 32767 -32768)
      (v128.const i16x8 -1 -32768 -5 6   -9 10 32767 32767)))
  (func (export "i32x4_extmul_high_u") (result v128)
    (i32x4.extmul_high_i16x8_u
      (v128.const i16x8 -1 -32768 3 -4   7 -8 32767 -32768)
      (v128.const i16x8 -1 -32768 -5 6   -9 10 32767 32767)))

  ;; ── i64x2.extmul_*_i32x4_{s,u} ───────────────────────────────────────
  (func (export "i64x2_extmul_low_s") (result v128)
    (i64x2.extmul_low_i32x4_s
      (v128.const i32x4 -1 -2147483648   3 -2147483648)
      (v128.const i32x4 -1 -2147483648   -5 2147483647)))
  (func (export "i64x2_extmul_low_u") (result v128)
    (i64x2.extmul_low_i32x4_u
      (v128.const i32x4 -1 -2147483648   3 -2147483648)
      (v128.const i32x4 -1 -2147483648   -5 2147483647)))
  (func (export "i64x2_extmul_high_s") (result v128)
    (i64x2.extmul_high_i32x4_s
      (v128.const i32x4 -1 -2147483648   3 -2147483648)
      (v128.const i32x4 -1 -2147483648   -5 2147483647)))
  (func (export "i64x2_extmul_high_u") (result v128)
    (i64x2.extmul_high_i32x4_u
      (v128.const i32x4 -1 -2147483648   3 -2147483648)
      (v128.const i32x4 -1 -2147483648   -5 2147483647)))

  ;; ── extadd_pairwise: adjacent lanes summed at the WIDER type ─────────
  ;; 127 + 127 = 254 and -128 + -128 = -256 both leave i8 range.
  (func (export "i16x8_extadd_pairwise_s") (result v128)
    (i16x8.extadd_pairwise_i8x16_s
      (v128.const i8x16 1 2 -1 -2 127 127 -128 -128  3 -3 0 0 -1 1 5 5)))
  (func (export "i16x8_extadd_pairwise_u") (result v128)
    (i16x8.extadd_pairwise_i8x16_u
      (v128.const i8x16 1 2 -1 -2 127 127 -128 -128  3 -3 0 0 -1 1 5 5)))
  (func (export "i32x4_extadd_pairwise_s") (result v128)
    (i32x4.extadd_pairwise_i16x8_s
      (v128.const i16x8 1 2 -1 -2 32767 32767 -32768 -32768)))
  (func (export "i32x4_extadd_pairwise_u") (result v128)
    (i32x4.extadd_pairwise_i16x8_u
      (v128.const i16x8 1 2 -1 -2 32767 32767 -32768 -32768)))

  ;; ── narrow: SATURATES, and the unsigned form clamps negatives to 0 ───
  (func (export "i8x16_narrow_s") (result v128)
    (i8x16.narrow_i16x8_s
      (v128.const i16x8 0 127 128 -128 -129 32767 -32768 1)
      (v128.const i16x8 -1 255 -256 100 -100 200 -200 2)))
  (func (export "i8x16_narrow_u") (result v128)
    (i8x16.narrow_i16x8_u
      (v128.const i16x8 0 127 128 -128 -129 32767 -32768 1)
      (v128.const i16x8 -1 255 -256 100 -100 200 -200 2)))
  (func (export "i16x8_narrow_s") (result v128)
    (i16x8.narrow_i32x4_s
      (v128.const i32x4 0 32767 32768 -32769)
      (v128.const i32x4 -1 100000 -100000 5)))
  (func (export "i16x8_narrow_u") (result v128)
    (i16x8.narrow_i32x4_u
      (v128.const i32x4 0 32767 32768 -32769)
      (v128.const i32x4 -1 100000 -100000 5)))

  ;; ── q15mulr_sat_s: fixed-point multiply with rounding, one saturation ─
  (func (export "i16x8_q15mulr_sat_s") (result v128)
    (i16x8.q15mulr_sat_s
      (v128.const i16x8 32767 -32768 16384 -16384 0 1 -1 32767)
      (v128.const i16x8 32767 -32768 16384 16384 100 1 1 -32768)))
)

;; ── extend ───────────────────────────────────────────────────────────
(assert_return (invoke "i16x8_extend_low_s")  (v128.const i16x8 0 1 -1 -128 127 -2 2 -127))
(assert_return (invoke "i16x8_extend_low_u")  (v128.const i16x8 0 1 255 128 127 254 2 129))
(assert_return (invoke "i16x8_extend_high_s") (v128.const i16x8 3 -3 -128 127 0 -1 1 -4))
(assert_return (invoke "i16x8_extend_high_u") (v128.const i16x8 3 253 128 127 0 255 1 252))

(assert_return (invoke "i32x4_extend_low_s")  (v128.const i32x4 0 1 -1 -32768))
(assert_return (invoke "i32x4_extend_low_u")  (v128.const i32x4 0 1 65535 32768))
(assert_return (invoke "i32x4_extend_high_s") (v128.const i32x4 32767 -2 2 -32767))
(assert_return (invoke "i32x4_extend_high_u") (v128.const i32x4 32767 65534 2 32769))

(assert_return (invoke "i64x2_extend_low_s")  (v128.const i64x2 -1 2147483647))
(assert_return (invoke "i64x2_extend_low_u")  (v128.const i64x2 4294967295 2147483647))
(assert_return (invoke "i64x2_extend_high_s") (v128.const i64x2 -2147483648 1))
(assert_return (invoke "i64x2_extend_high_u") (v128.const i64x2 2147483648 1))

;; ── extmul ───────────────────────────────────────────────────────────
;; low_s: 2*3, -1*-1, -128*-1, 127*2, -2*-2, 1*-1, -3*-1, 3*-1
(assert_return (invoke "i16x8_extmul_low_s")  (v128.const i16x8 6 1 128 254 4 -1 3 -3))
;; low_u: 2*3, 255*255, 128*255, 127*2, 254*254, 1*255, 253*255, 3*255
(assert_return (invoke "i16x8_extmul_low_u")  (v128.const i16x8 6 -511 32640 254 -1020 255 -1021 765))
;; high_s: 4*2, -4*2, -128*-128, 127*127, 5*-1, -5*-1, 6*3, -6*3
(assert_return (invoke "i16x8_extmul_high_s") (v128.const i16x8 8 -8 16384 16129 -5 5 18 -18))
;; high_u: 4*2, 252*2, 128*128, 127*127, 5*255, 251*255, 6*3, 250*3
(assert_return (invoke "i16x8_extmul_high_u") (v128.const i16x8 8 504 16384 16129 1275 -1531 18 750))

(assert_return (invoke "i32x4_extmul_low_s")  (v128.const i32x4 1 1073741824 -15 -24))
(assert_return (invoke "i32x4_extmul_low_u")  (v128.const i32x4 4294836225 1073741824 196593 393192))
(assert_return (invoke "i32x4_extmul_high_s") (v128.const i32x4 -63 -80 1073676289 -1073709056))
(assert_return (invoke "i32x4_extmul_high_u") (v128.const i32x4 458689 655280 1073676289 1073709056))

(assert_return (invoke "i64x2_extmul_low_s")  (v128.const i64x2 1 4611686018427387904))
(assert_return (invoke "i64x2_extmul_low_u")  (v128.const i64x2 18446744065119617025 4611686018427387904))
(assert_return (invoke "i64x2_extmul_high_s") (v128.const i64x2 -15 -4611686016279904256))
(assert_return (invoke "i64x2_extmul_high_u") (v128.const i64x2 12884901873 4611686016279904256))

;; ── extadd_pairwise ──────────────────────────────────────────────────
(assert_return (invoke "i16x8_extadd_pairwise_s") (v128.const i16x8 3 -3 254 -256 0 0 0 10))
(assert_return (invoke "i16x8_extadd_pairwise_u") (v128.const i16x8 3 509 254 256 256 0 256 10))
(assert_return (invoke "i32x4_extadd_pairwise_s") (v128.const i32x4 3 -3 65534 -65536))
(assert_return (invoke "i32x4_extadd_pairwise_u") (v128.const i32x4 3 131069 65534 65536))

;; ── narrow ───────────────────────────────────────────────────────────
(assert_return (invoke "i8x16_narrow_s")
  (v128.const i8x16 0 127 127 -128 -128 127 -128 1   -1 127 -128 100 -100 127 -128 2))
(assert_return (invoke "i8x16_narrow_u")
  (v128.const i8x16 0 127 -128 0 0 -1 0 1   0 -1 0 100 0 -56 0 2))
(assert_return (invoke "i16x8_narrow_s") (v128.const i16x8 0 32767 32767 -32768  -1 32767 -32768 5))
(assert_return (invoke "i16x8_narrow_u") (v128.const i16x8 0 32767 -32768 0  0 -1 0 5))

;; ── q15mulr_sat_s ────────────────────────────────────────────────────
(assert_return (invoke "i16x8_q15mulr_sat_s") (v128.const i16x8 32766 32767 8192 -8192 0 0 0 -32767))
