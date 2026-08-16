;; vybe-test: wast/wat_ops_i64/i64_unsigned_div_rem
;; origin: coverage gap vs crates/vybe_runtime/tests/i64_ops_test.rs
;; vybe-test-mode: run
;;
;; `i64.div_u` / `i64.rem_u` read BOTH operands as unsigned. On non-negative
;; inputs they agree with the signed forms exactly, so only a negative
;; operand — i.e. one with the top bit set — can tell them apart.
;;
;; The sharpest case is -1 / 2: signed that is 0, unsigned it is
;; 0x7FFF_FFFF_FFFF_FFFF. And -1 as a DIVISOR is the largest unsigned value,
;; so `x div_u -1` is 0 for every x below it — where `x div_s -1` is -x.
;;
;; Division by zero traps for both, and `rem_u` has no overflow case (unlike
;; `rem_s`, where i64.min % -1 is the special one).

(module
  (func (export "div_u") (param i64 i64) (result i64) (i64.div_u (local.get 0) (local.get 1)))
  (func (export "rem_u") (param i64 i64) (result i64) (i64.rem_u (local.get 0) (local.get 1)))
)

;; ── where unsigned and signed give different answers ────────────────────
(assert_return (invoke "div_u" (i64.const -1) (i64.const 2)) (i64.const 0x7fffffffffffffff))
(assert_return (invoke "rem_u" (i64.const -1) (i64.const 2)) (i64.const 1))
(assert_return (invoke "div_u" (i64.const 0x8000000000000000) (i64.const 2)) (i64.const 0x4000000000000000))
;; -1 as divisor is the maximum unsigned value: everything below it divides to 0.
(assert_return (invoke "div_u" (i64.const 5) (i64.const -1)) (i64.const 0))
(assert_return (invoke "rem_u" (i64.const 5) (i64.const -1)) (i64.const 5))
(assert_return (invoke "div_u" (i64.const -1) (i64.const -1)) (i64.const 1))
(assert_return (invoke "rem_u" (i64.const -1) (i64.const -1)) (i64.const 0))
;; i64.min is NOT a special case unsigned — no overflow, unlike div_s.
(assert_return (invoke "div_u" (i64.const 0x8000000000000000) (i64.const -1)) (i64.const 0))
(assert_return (invoke "rem_u" (i64.const 0x8000000000000000) (i64.const -1)) (i64.const 0x8000000000000000))

;; ── ordinary arithmetic, and truncation toward zero ─────────────────────
(assert_return (invoke "div_u" (i64.const 7) (i64.const 2)) (i64.const 3))
(assert_return (invoke "rem_u" (i64.const 7) (i64.const 2)) (i64.const 1))
(assert_return (invoke "div_u" (i64.const 0) (i64.const 7)) (i64.const 0))
(assert_return (invoke "rem_u" (i64.const 0) (i64.const 7)) (i64.const 0))
;; Exceeds 32 bits — a path that narrowed to i32 cannot produce this.
(assert_return (invoke "div_u" (i64.const 0x100000000) (i64.const 2)) (i64.const 0x80000000))

;; ── division by zero traps, for both ────────────────────────────────────
(assert_trap (invoke "div_u" (i64.const 1) (i64.const 0)) "integer divide by zero")
(assert_trap (invoke "rem_u" (i64.const 1) (i64.const 0)) "integer divide by zero")
(assert_trap (invoke "div_u" (i64.const 0) (i64.const 0)) "integer divide by zero")
