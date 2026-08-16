;; vybe-test: wast/wat_integer_conversions/unsigned_compare_and_float_to_unsigned
;; vybe-test-mode: run
;;
;; The last two scalar mnemonics still at one occurrence: `i32.gt_u` and
;; `i32.trunc_f64_u`. Both were asserted on an operand where the unsigned
;; reading and the signed one agree, which is the one place they say nothing.
;;
;;   * `i32.gt_u` differs from `i32.gt_s` exactly when a sign bit is set: -1 is
;;     the LARGEST unsigned value and the smallest signed one.
;;   * `i32.trunc_f64_u` is defined for `-1 < trunc(z) < 2^32`. So -0.5
;;     truncates to 0 and is NOT a trap, while -1.0 IS one — the boundary sits
;;     between them, and an implementation that rejects everything negative
;;     passes any test that only checks -1.0. 2^32-1 is the last representable
;;     value; 2^32 traps. NaN and the infinities always trap.
;;
;; Spec-format so `wasmtime wast` arbitrates.

(module
  (func (export "gt_u") (param i32 i32) (result i32) (i32.gt_u (local.get 0) (local.get 1)))
  (func (export "gt_s") (param i32 i32) (result i32) (i32.gt_s (local.get 0) (local.get 1)))
  (func (export "ge_u") (param i32 i32) (result i32) (i32.ge_u (local.get 0) (local.get 1)))
  (func (export "trunc_f64_u") (param f64) (result i32) (i32.trunc_f64_u (local.get 0)))
  (func (export "trunc_f64_s") (param f64) (result i32) (i32.trunc_f64_s (local.get 0)))
)

;; ── gt_u vs gt_s: they disagree on every pair with a sign bit ────────
(assert_return (invoke "gt_u" (i32.const -1) (i32.const 1)) (i32.const 1))
(assert_return (invoke "gt_s" (i32.const -1) (i32.const 1)) (i32.const 0))
(assert_return (invoke "gt_u" (i32.const 1) (i32.const -1)) (i32.const 0))
(assert_return (invoke "gt_s" (i32.const 1) (i32.const -1)) (i32.const 1))
(assert_return (invoke "gt_u" (i32.const -2147483648) (i32.const 2147483647)) (i32.const 1))
(assert_return (invoke "gt_s" (i32.const -2147483648) (i32.const 2147483647)) (i32.const 0))
;; Equal operands are not greater, under either reading.
(assert_return (invoke "gt_u" (i32.const -1) (i32.const -1)) (i32.const 0))
(assert_return (invoke "ge_u" (i32.const -1) (i32.const -1)) (i32.const 1))
(assert_return (invoke "gt_u" (i32.const 0) (i32.const 0)) (i32.const 0))
;; 0 is the unsigned minimum, so nothing is below it.
(assert_return (invoke "gt_u" (i32.const 0) (i32.const -1)) (i32.const 0))

;; ── trunc_f64_u: the negative boundary is -1, not 0 ─────────────────
(assert_return (invoke "trunc_f64_u" (f64.const 0.0)) (i32.const 0))
(assert_return (invoke "trunc_f64_u" (f64.const -0.0)) (i32.const 0))
(assert_return (invoke "trunc_f64_u" (f64.const -0.5)) (i32.const 0))
(assert_return (invoke "trunc_f64_u" (f64.const -0.9999999999999999)) (i32.const 0))
(assert_return (invoke "trunc_f64_u" (f64.const 1.9)) (i32.const 1))
(assert_return (invoke "trunc_f64_u" (f64.const 2147483648.0)) (i32.const -2147483648))
(assert_return (invoke "trunc_f64_u" (f64.const 4294967295.0)) (i32.const -1))
;; …and the same operands read as signed, where 2^31 is already out of range.
(assert_return (invoke "trunc_f64_s" (f64.const -0.5)) (i32.const 0))
(assert_return (invoke "trunc_f64_s" (f64.const -1.5)) (i32.const -1))
(assert_return (invoke "trunc_f64_s" (f64.const 2147483647.0)) (i32.const 2147483647))

(assert_trap (invoke "trunc_f64_u" (f64.const -1.0)) "integer overflow")
(assert_trap (invoke "trunc_f64_u" (f64.const 4294967296.0)) "integer overflow")
(assert_trap (invoke "trunc_f64_u" (f64.const nan)) "invalid conversion to integer")
(assert_trap (invoke "trunc_f64_u" (f64.const inf)) "integer overflow")
(assert_trap (invoke "trunc_f64_u" (f64.const -inf)) "integer overflow")
(assert_trap (invoke "trunc_f64_s" (f64.const 2147483648.0)) "integer overflow")
(assert_trap (invoke "trunc_f64_s" (f64.const -2147483649.0)) "integer overflow")
