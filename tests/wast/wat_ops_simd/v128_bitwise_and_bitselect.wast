;; vybe-test: wast/wat_ops_simd/v128_bitwise_and_bitselect
;; origin: coverage gap — 17 v128 mnemonics occurred at most ONCE in the run corpus
;; vybe-test-mode: run
;;
;; The v128 bitwise ops are lane-agnostic: they see 128 bits, not 4×i32. That
;; makes them the right place to check the vector representation itself, before
;; any lane-typed arithmetic is involved.
;;
;; The two that carry real semantics:
;;
;;   * `bitselect` is per-BIT, not per-lane: `(a & c) | (b & ~c)`. A mask that
;;     is all-ones or all-zeros per lane cannot tell it apart from a lane-wise
;;     select, so the mixed-mask cases below are the ones that matter.
;;   * `any_true` is a reduction over ALL 128 bits, while `i32x4.all_true` is
;;     per-lane — so a vector with one non-zero BIT in a lane is "any true" and
;;     that lane is non-zero, but a vector with a zero lane is not "all true".
;;
;; `v128.const` is written as i32x4 throughout so every expectation is an exact
;; bit pattern.

(module
  (func (export "not") (param v128) (result v128) (v128.not (local.get 0)))
  (func (export "and") (param v128 v128) (result v128) (v128.and (local.get 0) (local.get 1)))
  (func (export "or") (param v128 v128) (result v128) (v128.or (local.get 0) (local.get 1)))
  (func (export "xor") (param v128 v128) (result v128) (v128.xor (local.get 0) (local.get 1)))
  (func (export "andnot") (param v128 v128) (result v128) (v128.andnot (local.get 0) (local.get 1)))
  (func (export "bitselect") (param v128 v128 v128) (result v128)
    (v128.bitselect (local.get 0) (local.get 1) (local.get 2)))
  (func (export "any_true") (param v128) (result i32) (v128.any_true (local.get 0)))
  (func (export "all_true_i32x4") (param v128) (result i32) (i32x4.all_true (local.get 0)))
)

;; ── not / and / or / xor across the full 128 bits ───────────────────────
(assert_return (invoke "not" (v128.const i32x4 0 0 0 0))
               (v128.const i32x4 -1 -1 -1 -1))
(assert_return (invoke "not" (v128.const i32x4 0x0f0f0f0f 0 -1 0x55555555))
               (v128.const i32x4 0xf0f0f0f0 -1 0 0xaaaaaaaa))
(assert_return (invoke "and" (v128.const i32x4 -1 -1 -1 -1) (v128.const i32x4 0x0f0f0f0f 0 -1 0x12345678))
               (v128.const i32x4 0x0f0f0f0f 0 -1 0x12345678))
(assert_return (invoke "or" (v128.const i32x4 0xffff0000 0 0 0) (v128.const i32x4 0x0000ffff 0 0 0))
               (v128.const i32x4 -1 0 0 0))
(assert_return (invoke "xor" (v128.const i32x4 0xaaaaaaaa 0 -1 5) (v128.const i32x4 0x55555555 0 -1 5))
               (v128.const i32x4 -1 0 0 0))

;; ── andnot is a & ~b, and is NOT commutative ────────────────────────────
;; The two orders give different answers — a symmetric implementation passes
;; only if both directions are checked.
(assert_return (invoke "andnot" (v128.const i32x4 0x0f0f0f0f 0 0 0) (v128.const i32x4 0x00ff00ff 0 0 0))
               (v128.const i32x4 0x0f000f00 0 0 0))
(assert_return (invoke "andnot" (v128.const i32x4 0x00ff00ff 0 0 0) (v128.const i32x4 0x0f0f0f0f 0 0 0))
               (v128.const i32x4 0x00f000f0 0 0 0))
(assert_return (invoke "andnot" (v128.const i32x4 -1 -1 -1 -1) (v128.const i32x4 -1 -1 -1 -1))
               (v128.const i32x4 0 0 0 0))

;; ── bitselect is per-BIT ────────────────────────────────────────────────
;; Whole-lane masks: indistinguishable from a lane-wise select.
(assert_return (invoke "bitselect" (v128.const i32x4 0xaaaaaaaa 0xaaaaaaaa 0xaaaaaaaa 0xaaaaaaaa)
                                   (v128.const i32x4 0x55555555 0x55555555 0x55555555 0x55555555)
                                   (v128.const i32x4 -1 0 -1 0))
               (v128.const i32x4 0xaaaaaaaa 0x55555555 0xaaaaaaaa 0x55555555))
;; MIXED mask inside a single lane: only a per-bit implementation gets this.
(assert_return (invoke "bitselect" (v128.const i32x4 0xffffffff 0 0 0)
                                   (v128.const i32x4 0x00000000 0 0 0)
                                   (v128.const i32x4 0x0f0f0f0f 0 0 0))
               (v128.const i32x4 0x0f0f0f0f 0 0 0))
(assert_return (invoke "bitselect" (v128.const i32x4 0xaaaaaaaa 0 0 0)
                                   (v128.const i32x4 0x55555555 0 0 0)
                                   (v128.const i32x4 0x0f0f0f0f 0 0 0))
               ;; (0xaaaaaaaa & 0x0f0f0f0f) | (0x55555555 & 0xf0f0f0f0)
               ;;  = 0x0a0a0a0a | 0x50505050
               (v128.const i32x4 0x5a5a5a5a 0 0 0))
;; A mask of all ones / all zeros selects one operand entirely.
(assert_return (invoke "bitselect" (v128.const i32x4 1 2 3 4)
                                   (v128.const i32x4 5 6 7 8)
                                   (v128.const i32x4 -1 -1 -1 -1))
               (v128.const i32x4 1 2 3 4))
(assert_return (invoke "bitselect" (v128.const i32x4 1 2 3 4)
                                   (v128.const i32x4 5 6 7 8)
                                   (v128.const i32x4 0 0 0 0))
               (v128.const i32x4 5 6 7 8))

;; ── any_true is over all 128 bits; all_true is per lane ─────────────────
(assert_return (invoke "any_true" (v128.const i32x4 0 0 0 0)) (i32.const 0))
;; A single set bit, in the LAST lane — a reduction that only reads lane 0 fails.
(assert_return (invoke "any_true" (v128.const i32x4 0 0 0 1)) (i32.const 1))
(assert_return (invoke "any_true" (v128.const i32x4 1 0 0 0)) (i32.const 1))
(assert_return (invoke "any_true" (v128.const i32x4 -1 -1 -1 -1)) (i32.const 1))
(assert_return (invoke "all_true_i32x4" (v128.const i32x4 1 1 1 1)) (i32.const 1))
;; One zero lane is enough to make all_true false — and any_true still true.
(assert_return (invoke "all_true_i32x4" (v128.const i32x4 1 1 1 0)) (i32.const 0))
(assert_return (invoke "all_true_i32x4" (v128.const i32x4 0 0 0 0)) (i32.const 0))
(assert_return (invoke "all_true_i32x4" (v128.const i32x4 -1 -1 -1 -1)) (i32.const 1))
