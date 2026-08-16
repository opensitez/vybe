;; vybe-test: wast/wat_ops_simd/i32x4_i64x2_arith_and_compare
;; origin: coverage gap — 30 i32x4 and 25 i64x2 mnemonics occurred at most ONCE in the run corpus
;; vybe-test-mode: run
;;
;; The two wide integer shapes together, because the mistakes are the same and
;; the boundary values are what differ.
;;
;;   * `extract_lane` on i32x4 has NO signed/unsigned pair (an i32 lane already
;;     fills the i32 result) — unlike i8x16/i16x8. i64x2 likewise. So the
;;     interesting lane property here is POSITION, and every vector below gives
;;     its lanes distinct values so a lane-0 broadcast is caught.
;;   * Arithmetic WRAPS at the lane width, and the lane width is the thing most
;;     easily got wrong: `i32x4.add` of 0x7fffffff + 1 must be -2147483648 and
;;     must NOT carry into the next lane.
;;   * Shift counts are modulo the LANE width — 32 and 64 respectively, so a
;;     shift by 32 is the identity for i32x4 but a real shift for i64x2. That
;;     single pair of assertions separates the two shapes' shift paths.
;;   * `lt_s` vs `lt_u` differ only on a lane with the top bit set.
;;
;; i64x2 has no unsigned comparisons in the spec (only `lt_s`/`gt_s`/`le_s`/
;; `ge_s`), so the unsigned cases below are i32x4 only — deliberately, not an
;; omission.

(module
  ;; i32x4
  (func (export "i32x4.splat") (param i32) (result v128) (i32x4.splat (local.get 0)))
  (func (export "i32x4.ext3") (param v128) (result i32) (i32x4.extract_lane 3 (local.get 0)))
  (func (export "i32x4.replace1") (param v128 i32) (result v128) (i32x4.replace_lane 1 (local.get 0) (local.get 1)))
  (func (export "i32x4.add") (param v128 v128) (result v128) (i32x4.add (local.get 0) (local.get 1)))
  (func (export "i32x4.sub") (param v128 v128) (result v128) (i32x4.sub (local.get 0) (local.get 1)))
  (func (export "i32x4.mul") (param v128 v128) (result v128) (i32x4.mul (local.get 0) (local.get 1)))
  (func (export "i32x4.neg") (param v128) (result v128) (i32x4.neg (local.get 0)))
  (func (export "i32x4.abs") (param v128) (result v128) (i32x4.abs (local.get 0)))
  (func (export "i32x4.shl") (param v128 i32) (result v128) (i32x4.shl (local.get 0) (local.get 1)))
  (func (export "i32x4.shr_s") (param v128 i32) (result v128) (i32x4.shr_s (local.get 0) (local.get 1)))
  (func (export "i32x4.shr_u") (param v128 i32) (result v128) (i32x4.shr_u (local.get 0) (local.get 1)))
  (func (export "i32x4.min_s") (param v128 v128) (result v128) (i32x4.min_s (local.get 0) (local.get 1)))
  (func (export "i32x4.min_u") (param v128 v128) (result v128) (i32x4.min_u (local.get 0) (local.get 1)))
  (func (export "i32x4.max_s") (param v128 v128) (result v128) (i32x4.max_s (local.get 0) (local.get 1)))
  (func (export "i32x4.max_u") (param v128 v128) (result v128) (i32x4.max_u (local.get 0) (local.get 1)))
  (func (export "i32x4.eq") (param v128 v128) (result v128) (i32x4.eq (local.get 0) (local.get 1)))
  (func (export "i32x4.lt_s") (param v128 v128) (result v128) (i32x4.lt_s (local.get 0) (local.get 1)))
  (func (export "i32x4.lt_u") (param v128 v128) (result v128) (i32x4.lt_u (local.get 0) (local.get 1)))
  (func (export "i32x4.bitmask") (param v128) (result i32) (i32x4.bitmask (local.get 0)))
  ;; i64x2
  (func (export "i64x2.splat") (param i64) (result v128) (i64x2.splat (local.get 0)))
  (func (export "i64x2.ext1") (param v128) (result i64) (i64x2.extract_lane 1 (local.get 0)))
  (func (export "i64x2.replace0") (param v128 i64) (result v128) (i64x2.replace_lane 0 (local.get 0) (local.get 1)))
  (func (export "i64x2.add") (param v128 v128) (result v128) (i64x2.add (local.get 0) (local.get 1)))
  (func (export "i64x2.sub") (param v128 v128) (result v128) (i64x2.sub (local.get 0) (local.get 1)))
  (func (export "i64x2.mul") (param v128 v128) (result v128) (i64x2.mul (local.get 0) (local.get 1)))
  (func (export "i64x2.neg") (param v128) (result v128) (i64x2.neg (local.get 0)))
  (func (export "i64x2.abs") (param v128) (result v128) (i64x2.abs (local.get 0)))
  (func (export "i64x2.shl") (param v128 i32) (result v128) (i64x2.shl (local.get 0) (local.get 1)))
  (func (export "i64x2.shr_s") (param v128 i32) (result v128) (i64x2.shr_s (local.get 0) (local.get 1)))
  (func (export "i64x2.shr_u") (param v128 i32) (result v128) (i64x2.shr_u (local.get 0) (local.get 1)))
  (func (export "i64x2.eq") (param v128 v128) (result v128) (i64x2.eq (local.get 0) (local.get 1)))
  (func (export "i64x2.lt_s") (param v128 v128) (result v128) (i64x2.lt_s (local.get 0) (local.get 1)))
  (func (export "i64x2.bitmask") (param v128) (result i32) (i64x2.bitmask (local.get 0)))
)

;; ── lane position: every lane distinct ──────────────────────────────────
(assert_return (invoke "i32x4.splat" (i32.const -1)) (v128.const i32x4 -1 -1 -1 -1))
(assert_return (invoke "i32x4.ext3" (v128.const i32x4 1 2 3 4)) (i32.const 4))
(assert_return (invoke "i32x4.replace1" (v128.const i32x4 1 2 3 4) (i32.const 9))
               (v128.const i32x4 1 9 3 4))
(assert_return (invoke "i64x2.splat" (i64.const -1)) (v128.const i64x2 -1 -1))
(assert_return (invoke "i64x2.ext1" (v128.const i64x2 1 2)) (i64.const 2))
(assert_return (invoke "i64x2.replace0" (v128.const i64x2 1 2) (i64.const 9))
               (v128.const i64x2 9 2))
;; splat TRUNCATES to the lane width rather than saturating.
(assert_return (invoke "i32x4.splat" (i32.const 0x80000000))
               (v128.const i32x4 0x80000000 0x80000000 0x80000000 0x80000000))

;; ── arithmetic wraps at the LANE width and never carries across lanes ───
(assert_return (invoke "i32x4.add" (v128.const i32x4 0x7fffffff 0x80000000 -1 1)
                                   (v128.const i32x4 1 -1 1 -1))
               (v128.const i32x4 0x80000000 0x7fffffff 0 0))
(assert_return (invoke "i32x4.sub" (v128.const i32x4 0x80000000 0 1 0)
                                   (v128.const i32x4 1 1 1 0))
               (v128.const i32x4 0x7fffffff -1 0 0))
;; The product exceeds 32 bits: the high half is DISCARDED, not carried.
(assert_return (invoke "i32x4.mul" (v128.const i32x4 0x10000 0x7fffffff -1 3)
                                   (v128.const i32x4 0x10000 2 -1 4))
               (v128.const i32x4 0 -2 1 12))
(assert_return (invoke "i64x2.add" (v128.const i64x2 0x7fffffffffffffff -1)
                                   (v128.const i64x2 1 1))
               (v128.const i64x2 0x8000000000000000 0))
(assert_return (invoke "i64x2.sub" (v128.const i64x2 0x8000000000000000 0)
                                   (v128.const i64x2 1 1))
               (v128.const i64x2 0x7fffffffffffffff -1))
;; 2^32 * 2^32 = 2^64 wraps to 0 — a lane narrowed to 32 bits cannot show this.
(assert_return (invoke "i64x2.mul" (v128.const i64x2 0x100000000 0x7fffffffffffffff)
                                   (v128.const i64x2 0x100000000 2))
               (v128.const i64x2 0 -2))

;; ── neg / abs at the asymmetric minimum ─────────────────────────────────
;; i32.min has no positive counterpart: neg and abs both return it unchanged.
(assert_return (invoke "i32x4.neg" (v128.const i32x4 0x80000000 1 -1 0))
               (v128.const i32x4 0x80000000 -1 1 0))
(assert_return (invoke "i32x4.abs" (v128.const i32x4 0x80000000 -1 1 0))
               (v128.const i32x4 0x80000000 1 1 0))
(assert_return (invoke "i64x2.neg" (v128.const i64x2 0x8000000000000000 1))
               (v128.const i64x2 0x8000000000000000 -1))
(assert_return (invoke "i64x2.abs" (v128.const i64x2 0x8000000000000000 -1))
               (v128.const i64x2 0x8000000000000000 1))

;; ── shift counts are modulo the LANE width: 32 here, 64 there ───────────
(assert_return (invoke "i32x4.shl" (v128.const i32x4 1 1 1 1) (i32.const 1))
               (v128.const i32x4 2 2 2 2))
;; Shift by 32 is the IDENTITY for i32x4 — a clamping implementation gives 0.
(assert_return (invoke "i32x4.shl" (v128.const i32x4 1 1 1 1) (i32.const 32))
               (v128.const i32x4 1 1 1 1))
(assert_return (invoke "i32x4.shl" (v128.const i32x4 1 1 1 1) (i32.const 33))
               (v128.const i32x4 2 2 2 2))
;; ...but 32 is a REAL shift for i64x2, which has twice the lane width.
(assert_return (invoke "i64x2.shl" (v128.const i64x2 1 1) (i32.const 32))
               (v128.const i64x2 0x100000000 0x100000000))
(assert_return (invoke "i64x2.shl" (v128.const i64x2 1 1) (i32.const 64))
               (v128.const i64x2 1 1))
;; shr_s copies the sign bit in; shr_u brings zeros.
(assert_return (invoke "i32x4.shr_s" (v128.const i32x4 -1 0x80000000 0x40000000 0) (i32.const 1))
               (v128.const i32x4 -1 0xc0000000 0x20000000 0))
(assert_return (invoke "i32x4.shr_u" (v128.const i32x4 -1 0x80000000 0x40000000 0) (i32.const 1))
               (v128.const i32x4 0x7fffffff 0x40000000 0x20000000 0))
(assert_return (invoke "i64x2.shr_s" (v128.const i64x2 -1 0x8000000000000000) (i32.const 1))
               (v128.const i64x2 -1 0xc000000000000000))
(assert_return (invoke "i64x2.shr_u" (v128.const i64x2 -1 0x8000000000000000) (i32.const 1))
               (v128.const i64x2 0x7fffffffffffffff 0x4000000000000000))

;; ── signed vs unsigned min/max: only a top-bit lane separates them ──────
(assert_return (invoke "i32x4.min_s" (v128.const i32x4 -1 1 0x80000000 5)
                                     (v128.const i32x4 1 -1 0x7fffffff 5))
               (v128.const i32x4 -1 -1 0x80000000 5))
(assert_return (invoke "i32x4.min_u" (v128.const i32x4 -1 1 0x80000000 5)
                                     (v128.const i32x4 1 -1 0x7fffffff 5))
               (v128.const i32x4 1 1 0x7fffffff 5))
(assert_return (invoke "i32x4.max_s" (v128.const i32x4 -1 0x80000000 0 0)
                                     (v128.const i32x4 1 0x7fffffff 0 0))
               (v128.const i32x4 1 0x7fffffff 0 0))
(assert_return (invoke "i32x4.max_u" (v128.const i32x4 -1 0x80000000 0 0)
                                     (v128.const i32x4 1 0x7fffffff 0 0))
               (v128.const i32x4 -1 0x80000000 0 0))

;; ── comparisons: all-ones / all-zeros lane masks ────────────────────────
(assert_return (invoke "i32x4.eq" (v128.const i32x4 1 2 3 4) (v128.const i32x4 1 0 3 0))
               (v128.const i32x4 -1 0 -1 0))
;; -1 < 1 signed but 0xffffffff > 1 unsigned: same lanes, opposite masks.
(assert_return (invoke "i32x4.lt_s" (v128.const i32x4 -1 1 0 0) (v128.const i32x4 1 -1 0 0))
               (v128.const i32x4 -1 0 0 0))
(assert_return (invoke "i32x4.lt_u" (v128.const i32x4 -1 1 0 0) (v128.const i32x4 1 -1 0 0))
               (v128.const i32x4 0 -1 0 0))
(assert_return (invoke "i64x2.eq" (v128.const i64x2 1 2) (v128.const i64x2 1 0))
               (v128.const i64x2 -1 0))
(assert_return (invoke "i64x2.lt_s" (v128.const i64x2 -1 1) (v128.const i64x2 1 -1))
               (v128.const i64x2 -1 0))

;; ── bitmask reads each lane's sign bit, lane 0 in bit 0 ─────────────────
(assert_return (invoke "i32x4.bitmask" (v128.const i32x4 -1 0 0 0)) (i32.const 1))
(assert_return (invoke "i32x4.bitmask" (v128.const i32x4 0 0 0 -1)) (i32.const 8))
(assert_return (invoke "i32x4.bitmask" (v128.const i32x4 -1 -1 -1 -1)) (i32.const 15))
;; 0x7fffffff is non-zero but its sign bit is CLEAR.
(assert_return (invoke "i32x4.bitmask" (v128.const i32x4 0x7fffffff 0x80000000 0 0)) (i32.const 2))
(assert_return (invoke "i64x2.bitmask" (v128.const i64x2 -1 0)) (i32.const 1))
(assert_return (invoke "i64x2.bitmask" (v128.const i64x2 0 -1)) (i32.const 2))
(assert_return (invoke "i64x2.bitmask" (v128.const i64x2 -1 -1)) (i32.const 3))
