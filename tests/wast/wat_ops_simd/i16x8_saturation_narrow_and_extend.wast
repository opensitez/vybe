;; vybe-test: wast/wat_ops_simd/i16x8_saturation_narrow_and_extend
;; origin: coverage gap — 34 i16x8 mnemonics occurred at most ONCE in the run corpus
;; vybe-test-mode: run
;;
;; i16x8, plus the ops that CHANGE shape — the ones with no scalar analogue at
;; all, and so the ones least likely to be right by accident:
;;
;;   * `narrow_i32x4_s`/`_u` pack two 4-lane vectors into eight 16-bit lanes
;;     with SATURATION, and the two halves land in a defined order (the first
;;     operand's lanes first). Order and saturation are independent mistakes.
;;   * `extend_low`/`extend_high_i8x16_s`/`_u` widen half a vector, so the
;;     _low_/_high_ choice and the sign choice are two more independent bits.
;;   * `extadd_pairwise` sums ADJACENT lanes, so a lane-pairing mistake shows
;;     up only when neighbouring lanes differ.
;;   * `avgr_u` is a ROUNDING average — `(a+b+1)>>1` — so it differs from a
;;     plain average exactly when `a+b` is odd, and it must compute the sum at
;;     17 bits: 0xffff and 0xffff average to 0xffff, not to 0x7fff.
;;
;; `extract_lane_s`/`_u` are here for the same reason as i8x16: they differ
;; only on a lane with the top bit set.

(module
  (func (export "splat") (param i32) (result v128) (i16x8.splat (local.get 0)))
  (func (export "ext_s0") (param v128) (result i32) (i16x8.extract_lane_s 0 (local.get 0)))
  (func (export "ext_u0") (param v128) (result i32) (i16x8.extract_lane_u 0 (local.get 0)))
  (func (export "ext_s7") (param v128) (result i32) (i16x8.extract_lane_s 7 (local.get 0)))
  (func (export "add") (param v128 v128) (result v128) (i16x8.add (local.get 0) (local.get 1)))
  (func (export "sub") (param v128 v128) (result v128) (i16x8.sub (local.get 0) (local.get 1)))
  (func (export "mul") (param v128 v128) (result v128) (i16x8.mul (local.get 0) (local.get 1)))
  (func (export "add_sat_s") (param v128 v128) (result v128) (i16x8.add_sat_s (local.get 0) (local.get 1)))
  (func (export "add_sat_u") (param v128 v128) (result v128) (i16x8.add_sat_u (local.get 0) (local.get 1)))
  (func (export "sub_sat_s") (param v128 v128) (result v128) (i16x8.sub_sat_s (local.get 0) (local.get 1)))
  (func (export "sub_sat_u") (param v128 v128) (result v128) (i16x8.sub_sat_u (local.get 0) (local.get 1)))
  (func (export "shl") (param v128 i32) (result v128) (i16x8.shl (local.get 0) (local.get 1)))
  (func (export "shr_s") (param v128 i32) (result v128) (i16x8.shr_s (local.get 0) (local.get 1)))
  (func (export "shr_u") (param v128 i32) (result v128) (i16x8.shr_u (local.get 0) (local.get 1)))
  (func (export "neg") (param v128) (result v128) (i16x8.neg (local.get 0)))
  (func (export "abs") (param v128) (result v128) (i16x8.abs (local.get 0)))
  (func (export "min_s") (param v128 v128) (result v128) (i16x8.min_s (local.get 0) (local.get 1)))
  (func (export "min_u") (param v128 v128) (result v128) (i16x8.min_u (local.get 0) (local.get 1)))
  (func (export "avgr_u") (param v128 v128) (result v128) (i16x8.avgr_u (local.get 0) (local.get 1)))
  (func (export "bitmask") (param v128) (result i32) (i16x8.bitmask (local.get 0)))
  (func (export "lt_s") (param v128 v128) (result v128) (i16x8.lt_s (local.get 0) (local.get 1)))
  (func (export "lt_u") (param v128 v128) (result v128) (i16x8.lt_u (local.get 0) (local.get 1)))
  ;; shape-changing
  (func (export "narrow_s") (param v128 v128) (result v128) (i16x8.narrow_i32x4_s (local.get 0) (local.get 1)))
  (func (export "narrow_u") (param v128 v128) (result v128) (i16x8.narrow_i32x4_u (local.get 0) (local.get 1)))
  (func (export "extend_low_s") (param v128) (result v128) (i16x8.extend_low_i8x16_s (local.get 0)))
  (func (export "extend_high_s") (param v128) (result v128) (i16x8.extend_high_i8x16_s (local.get 0)))
  (func (export "extend_low_u") (param v128) (result v128) (i16x8.extend_low_i8x16_u (local.get 0)))
  (func (export "extadd_pair_s") (param v128) (result v128) (i16x8.extadd_pairwise_i8x16_s (local.get 0)))
  (func (export "extadd_pair_u") (param v128) (result v128) (i16x8.extadd_pairwise_i8x16_u (local.get 0)))
)

;; ── lanes: splat truncates to 16 bits, extract_s/_u split on the top bit ─
(assert_return (invoke "splat" (i32.const 0x12345)) (v128.const i16x8 0x2345 0x2345 0x2345 0x2345 0x2345 0x2345 0x2345 0x2345))
(assert_return (invoke "ext_s0" (v128.const i16x8 0x8000 0 0 0 0 0 0 0)) (i32.const -32768))
(assert_return (invoke "ext_u0" (v128.const i16x8 0x8000 0 0 0 0 0 0 0)) (i32.const 32768))
(assert_return (invoke "ext_s0" (v128.const i16x8 0x7fff 0 0 0 0 0 0 0)) (i32.const 32767))
(assert_return (invoke "ext_u0" (v128.const i16x8 0x7fff 0 0 0 0 0 0 0)) (i32.const 32767))
;; The LAST lane, so a lane-0 broadcast is caught.
(assert_return (invoke "ext_s7" (v128.const i16x8 0 0 0 0 0 0 0 -1)) (i32.const -1))

;; ── add wraps; add_sat clamps, at a different ceiling per signedness ────
(assert_return (invoke "add" (v128.const i16x8 0x7fff 0x8000 -1 0 0 0 0 0)
                             (v128.const i16x8 1 -1 1 0 0 0 0 0))
               (v128.const i16x8 0x8000 0x7fff 0 0 0 0 0 0))
(assert_return (invoke "add_sat_s" (v128.const i16x8 0x7fff 0x8000 5 0 0 0 0 0)
                                   (v128.const i16x8 1 -1 5 0 0 0 0 0))
               (v128.const i16x8 0x7fff 0x8000 10 0 0 0 0 0))
(assert_return (invoke "add_sat_u" (v128.const i16x8 0xffff 0x8000 5 0 0 0 0 0)
                                   (v128.const i16x8 1 0x8000 5 0 0 0 0 0))
               (v128.const i16x8 0xffff 0xffff 10 0 0 0 0 0))
(assert_return (invoke "sub_sat_s" (v128.const i16x8 0x8000 0x7fff 0 0 0 0 0 0)
                                   (v128.const i16x8 1 -1 0 0 0 0 0 0))
               (v128.const i16x8 0x8000 0x7fff 0 0 0 0 0 0))
;; sub_sat_u floors at 0 rather than wrapping to 0xffff.
(assert_return (invoke "sub_sat_u" (v128.const i16x8 0 5 0xffff 0 0 0 0 0)
                                   (v128.const i16x8 1 10 1 0 0 0 0 0))
               (v128.const i16x8 0 0 0xfffe 0 0 0 0 0))
(assert_return (invoke "sub" (v128.const i16x8 0 0 0 0 0 0 0 0)
                             (v128.const i16x8 1 0 0 0 0 0 0 0))
               (v128.const i16x8 -1 0 0 0 0 0 0 0))
;; The product exceeds 16 bits: the high half is discarded, not carried.
(assert_return (invoke "mul" (v128.const i16x8 0x100 0x7fff -1 3 0 0 0 0)
                             (v128.const i16x8 0x100 2 -1 4 0 0 0 0))
               (v128.const i16x8 0 -2 1 12 0 0 0 0))

;; ── shift counts are modulo the LANE width (16) ─────────────────────────
(assert_return (invoke "shl" (v128.const i16x8 1 1 1 1 1 1 1 1) (i32.const 1))
               (v128.const i16x8 2 2 2 2 2 2 2 2))
(assert_return (invoke "shl" (v128.const i16x8 1 1 1 1 1 1 1 1) (i32.const 16))
               (v128.const i16x8 1 1 1 1 1 1 1 1))
(assert_return (invoke "shl" (v128.const i16x8 1 1 1 1 1 1 1 1) (i32.const 17))
               (v128.const i16x8 2 2 2 2 2 2 2 2))
;; Bits leaving a lane are discarded, not carried into the neighbour.
(assert_return (invoke "shl" (v128.const i16x8 0x8000 0 0x4000 0 0 0 0 0) (i32.const 1))
               (v128.const i16x8 0 0 0x8000 0 0 0 0 0))
(assert_return (invoke "shr_s" (v128.const i16x8 -1 0x8000 0x4000 0 0 0 0 0) (i32.const 1))
               (v128.const i16x8 -1 0xc000 0x2000 0 0 0 0 0))
(assert_return (invoke "shr_u" (v128.const i16x8 -1 0x8000 0x4000 0 0 0 0 0) (i32.const 1))
               (v128.const i16x8 0x7fff 0x4000 0x2000 0 0 0 0 0))

;; ── neg / abs at the asymmetric minimum ─────────────────────────────────
(assert_return (invoke "neg" (v128.const i16x8 0x8000 1 -1 0 0 0 0 0))
               (v128.const i16x8 0x8000 -1 1 0 0 0 0 0))
(assert_return (invoke "abs" (v128.const i16x8 0x8000 -1 1 0 0 0 0 0))
               (v128.const i16x8 0x8000 1 1 0 0 0 0 0))

;; ── signed vs unsigned min, and the comparison masks ────────────────────
(assert_return (invoke "min_s" (v128.const i16x8 -1 0 0 0 0 0 0 0) (v128.const i16x8 1 0 0 0 0 0 0 0))
               (v128.const i16x8 -1 0 0 0 0 0 0 0))
(assert_return (invoke "min_u" (v128.const i16x8 -1 0 0 0 0 0 0 0) (v128.const i16x8 1 0 0 0 0 0 0 0))
               (v128.const i16x8 1 0 0 0 0 0 0 0))
(assert_return (invoke "lt_s" (v128.const i16x8 -1 1 0 0 0 0 0 0) (v128.const i16x8 1 -1 0 0 0 0 0 0))
               (v128.const i16x8 -1 0 0 0 0 0 0 0))
(assert_return (invoke "lt_u" (v128.const i16x8 -1 1 0 0 0 0 0 0) (v128.const i16x8 1 -1 0 0 0 0 0 0))
               (v128.const i16x8 0 -1 0 0 0 0 0 0))

;; ── avgr_u ROUNDS UP and sums at 17 bits ────────────────────────────────
;; (a+b+1)>>1: odd sums round up, so 1 and 2 average to 2, not 1.
(assert_return (invoke "avgr_u" (v128.const i16x8 1 2 3 0 0 0 0 0)
                                (v128.const i16x8 2 2 4 0 0 0 0 0))
               (v128.const i16x8 2 2 4 0 0 0 0 0))
;; 0xffff+0xffff+1 overflows 16 bits — computed at 17 bits it is 0xffff.
(assert_return (invoke "avgr_u" (v128.const i16x8 0xffff 0xffff 0 0 0 0 0 0)
                                (v128.const i16x8 0xffff 0 0 0 0 0 0 0))
               (v128.const i16x8 0xffff 0x8000 0 0 0 0 0 0))

;; ── bitmask: lane 0 in bit 0 ────────────────────────────────────────────
(assert_return (invoke "bitmask" (v128.const i16x8 -1 0 0 0 0 0 0 0)) (i32.const 1))
(assert_return (invoke "bitmask" (v128.const i16x8 0 0 0 0 0 0 0 -1)) (i32.const 0x80))
(assert_return (invoke "bitmask" (v128.const i16x8 -1 -1 -1 -1 -1 -1 -1 -1)) (i32.const 0xff))
(assert_return (invoke "bitmask" (v128.const i16x8 0x7fff 0x8000 0 0 0 0 0 0)) (i32.const 2))

;; ── narrow: saturating pack, first operand's lanes FIRST ────────────────
;; Lanes 0-3 come from operand A, lanes 4-7 from operand B. A reversed
;; implementation produces the mirror image of this.
(assert_return (invoke "narrow_s" (v128.const i32x4 1 2 3 4) (v128.const i32x4 5 6 7 8))
               (v128.const i16x8 1 2 3 4 5 6 7 8))
;; Out-of-range lanes SATURATE to the i16 extremes, they do not wrap.
(assert_return (invoke "narrow_s" (v128.const i32x4 0x7fffffff -0x80000000 32767 -32768)
                                  (v128.const i32x4 32768 -32769 0 0))
               (v128.const i16x8 0x7fff 0x8000 32767 -32768 0x7fff 0x8000 0 0))
;; The unsigned form saturates to [0, 0xffff]: a negative lane becomes 0.
(assert_return (invoke "narrow_u" (v128.const i32x4 -1 0x10000 0xffff 0)
                                  (v128.const i32x4 0 0 0 0))
               (v128.const i16x8 0 0xffff 0xffff 0 0 0 0 0))

;; ── extend: which HALF, and which sign ──────────────────────────────────
;; The two halves are given different values so a low/high mix-up is caught.
(assert_return (invoke "extend_low_s" (v128.const i8x16 -1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16))
               (v128.const i16x8 -1 2 3 4 5 6 7 8))
(assert_return (invoke "extend_high_s" (v128.const i8x16 -1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16))
               (v128.const i16x8 9 10 11 12 13 14 15 16))
;; Same bytes, unsigned: 0xff widens to 255 rather than -1.
(assert_return (invoke "extend_low_u" (v128.const i8x16 -1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16))
               (v128.const i16x8 255 2 3 4 5 6 7 8))

;; ── extadd_pairwise sums ADJACENT lanes ─────────────────────────────────
;; Neighbours differ, so a wrong pairing gives a different answer.
(assert_return (invoke "extadd_pair_s" (v128.const i8x16 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16))
               (v128.const i16x8 3 7 11 15 19 23 27 31))
;; Signed vs unsigned on the same bytes: -1 + -1 is -2, but 255 + 255 is 510.
(assert_return (invoke "extadd_pair_s" (v128.const i8x16 -1 -1 0 0 0 0 0 0 0 0 0 0 0 0 0 0))
               (v128.const i16x8 -2 0 0 0 0 0 0 0))
(assert_return (invoke "extadd_pair_u" (v128.const i8x16 -1 -1 0 0 0 0 0 0 0 0 0 0 0 0 0 0))
               (v128.const i16x8 510 0 0 0 0 0 0 0))
