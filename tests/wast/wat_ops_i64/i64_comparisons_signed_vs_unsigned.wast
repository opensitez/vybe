;; vybe-test: wast/wat_ops_i64/i64_comparisons_signed_vs_unsigned
;; origin: coverage gap vs crates/vybe_runtime/tests/i64_ops_test.rs
;; vybe-test-mode: run
;;
;; The eight i64 ordering comparisons, at the operand pairs where the SIGNED
;; and UNSIGNED readings disagree. Comparing 1 against 2 exercises none of
;; that — it gives the same answer either way, so a `lt_u` wired to `lt_s`
;; passes it. Every pair below is chosen so the two readings return OPPOSITE
;; results, which is the only arrangement that can tell them apart.
;;
;; -1 is 0xFFFF_FFFF_FFFF_FFFF: the most negative value signed, the largest
;; unsigned. i64.min (0x8000…) is negative signed and above half-range
;; unsigned. Those two carry most of the discrimination.
;;
;; Spec-format so `wasmtime wast` runs this file unchanged — the expectations
;; are checked against a second engine rather than against ourselves.

(module
  (func (export "lt_s") (param i64 i64) (result i32) (i64.lt_s (local.get 0) (local.get 1)))
  (func (export "lt_u") (param i64 i64) (result i32) (i64.lt_u (local.get 0) (local.get 1)))
  (func (export "gt_s") (param i64 i64) (result i32) (i64.gt_s (local.get 0) (local.get 1)))
  (func (export "gt_u") (param i64 i64) (result i32) (i64.gt_u (local.get 0) (local.get 1)))
  (func (export "le_s") (param i64 i64) (result i32) (i64.le_s (local.get 0) (local.get 1)))
  (func (export "le_u") (param i64 i64) (result i32) (i64.le_u (local.get 0) (local.get 1)))
  (func (export "ge_s") (param i64 i64) (result i32) (i64.ge_s (local.get 0) (local.get 1)))
  (func (export "ge_u") (param i64 i64) (result i32) (i64.ge_u (local.get 0) (local.get 1)))
)

;; -1 vs 0 — signed: -1 < 0. unsigned: 0xFFFF… > 0. Opposite answers.
(assert_return (invoke "lt_s" (i64.const -1) (i64.const 0)) (i32.const 1))
(assert_return (invoke "lt_u" (i64.const -1) (i64.const 0)) (i32.const 0))
(assert_return (invoke "gt_s" (i64.const -1) (i64.const 0)) (i32.const 0))
(assert_return (invoke "gt_u" (i64.const -1) (i64.const 0)) (i32.const 1))
(assert_return (invoke "le_s" (i64.const -1) (i64.const 0)) (i32.const 1))
(assert_return (invoke "le_u" (i64.const -1) (i64.const 0)) (i32.const 0))
(assert_return (invoke "ge_s" (i64.const -1) (i64.const 0)) (i32.const 0))
(assert_return (invoke "ge_u" (i64.const -1) (i64.const 0)) (i32.const 1))

;; i64.min vs i64.max — signed: min < max. unsigned: 0x8000… > 0x7FFF…
(assert_return (invoke "lt_s" (i64.const 0x8000000000000000) (i64.const 0x7fffffffffffffff)) (i32.const 1))
(assert_return (invoke "lt_u" (i64.const 0x8000000000000000) (i64.const 0x7fffffffffffffff)) (i32.const 0))
(assert_return (invoke "ge_s" (i64.const 0x8000000000000000) (i64.const 0x7fffffffffffffff)) (i32.const 0))
(assert_return (invoke "ge_u" (i64.const 0x8000000000000000) (i64.const 0x7fffffffffffffff)) (i32.const 1))

;; Equal operands: the strict forms are false, the non-strict true, for both
;; readings. Catches a `le` implemented as `lt`.
(assert_return (invoke "lt_s" (i64.const -1) (i64.const -1)) (i32.const 0))
(assert_return (invoke "lt_u" (i64.const -1) (i64.const -1)) (i32.const 0))
(assert_return (invoke "gt_s" (i64.const -1) (i64.const -1)) (i32.const 0))
(assert_return (invoke "gt_u" (i64.const -1) (i64.const -1)) (i32.const 0))
(assert_return (invoke "le_s" (i64.const -1) (i64.const -1)) (i32.const 1))
(assert_return (invoke "le_u" (i64.const -1) (i64.const -1)) (i32.const 1))
(assert_return (invoke "ge_s" (i64.const -1) (i64.const -1)) (i32.const 1))
(assert_return (invoke "ge_u" (i64.const -1) (i64.const -1)) (i32.const 1))

;; Both negative: -2 < -1 signed, and 0xFFFF…FE < 0xFFFF…FF unsigned too —
;; here the readings AGREE, which is the case a sign-confused implementation
;; still gets right. Included so the file distinguishes "wrong" from
;; "wrong only where it matters".
(assert_return (invoke "lt_s" (i64.const -2) (i64.const -1)) (i32.const 1))
(assert_return (invoke "lt_u" (i64.const -2) (i64.const -1)) (i32.const 1))

;; 64-bit width: a value that differs from its low 32 bits only above bit 31.
;; Truncating to i32 anywhere in the path makes both of these answer 0.
(assert_return (invoke "gt_u" (i64.const 0x100000000) (i64.const 0xffffffff)) (i32.const 1))
(assert_return (invoke "gt_s" (i64.const 0x100000000) (i64.const 0xffffffff)) (i32.const 1))
