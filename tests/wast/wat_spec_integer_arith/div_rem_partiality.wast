;; vybe-test: wast/wat_spec_integer_arith/div_rem_partiality
;; vybe-test-mode: run
;;
;; Derived from the spec's own definitions of the integer division and
;; remainder operators (core/exec/numerics.rst, `op-idiv` / `op-irem`).
;; These four are the only PARTIAL integer arithmetic operators, and the spec
;; states each undefined case explicitly. What matters is which cases are
;; undefined and which are not — the asymmetry between div_s and rem_s is the
;; whole content of the rule and is easy to get wrong in both directions.
;;
;;   idiv_u(i1, 0)  = {}                        -- undefined
;;   idiv_s(i1, 0)  = {}                        -- undefined
;;   idiv_s(i1, i2) = {}   iff  j1 / j2 = 2^(N-1)   -- the ONE overflow case
;;   irem_u(i1, 0)  = {}                        -- undefined
;;   irem_s(i1, 0)  = {}                        -- undefined
;;
;; Note the spec gives rem_s NO overflow case: `irem_s(INT_MIN, -1)` is
;; defined and is 0, even though `idiv_s(INT_MIN, -1)` traps. A implementation
;; that guards both with one shared check is wrong.
;;
;; Truncation is toward zero (`truncz`), not floor, so the sign of the operands
;; decides the result and the remainder carries the sign of the DIVIDEND.

(module
  (func (export "div_s") (param i32 i32) (result i32) (i32.div_s (local.get 0) (local.get 1)))
  (func (export "div_u") (param i32 i32) (result i32) (i32.div_u (local.get 0) (local.get 1)))
  (func (export "rem_s") (param i32 i32) (result i32) (i32.rem_s (local.get 0) (local.get 1)))
  (func (export "rem_u") (param i32 i32) (result i32) (i32.rem_u (local.get 0) (local.get 1)))
  (func (export "div_s64") (param i64 i64) (result i64) (i64.div_s (local.get 0) (local.get 1)))
  (func (export "rem_s64") (param i64 i64) (result i64) (i64.rem_s (local.get 0) (local.get 1)))
  (func (export "div_u64") (param i64 i64) (result i64) (i64.div_u (local.get 0) (local.get 1)))
  (func (export "rem_u64") (param i64 i64) (result i64) (i64.rem_u (local.get 0) (local.get 1)))
  ;; The spec's stated identity: as long as both are defined,
  ;;   i1 = i2 * idiv(i1, i2) + irem(i1, i2)
  (func (export "identity_s") (param i32 i32) (result i32)
    (i32.add (i32.mul (local.get 1) (i32.div_s (local.get 0) (local.get 1)))
             (i32.rem_s (local.get 0) (local.get 1))))
  (func (export "identity_u") (param i32 i32) (result i32)
    (i32.add (i32.mul (local.get 1) (i32.div_u (local.get 0) (local.get 1)))
             (i32.rem_u (local.get 0) (local.get 1))))
)

;; ── The undefined cases: every one of them, and no others ──────────────────
(assert_trap (invoke "div_s" (i32.const 1) (i32.const 0)) "integer divide by zero")
(assert_trap (invoke "div_s" (i32.const 0) (i32.const 0)) "integer divide by zero")
(assert_trap (invoke "div_u" (i32.const 1) (i32.const 0)) "integer divide by zero")
(assert_trap (invoke "rem_s" (i32.const 1) (i32.const 0)) "integer divide by zero")
(assert_trap (invoke "rem_u" (i32.const 1) (i32.const 0)) "integer divide by zero")
(assert_trap (invoke "div_s64" (i64.const 1) (i64.const 0)) "integer divide by zero")
(assert_trap (invoke "rem_s64" (i64.const 1) (i64.const 0)) "integer divide by zero")

;; The single overflow case: j1/j2 = 2^(N-1), i.e. INT_MIN / -1.
(assert_trap (invoke "div_s" (i32.const -2147483648) (i32.const -1)) "integer overflow")
(assert_trap (invoke "div_s64" (i64.const -9223372036854775808) (i64.const -1)) "integer overflow")

;; ...and the case that is NOT undefined, which the same guard would wrongly
;; catch. The spec defines rem_s here; the answer is 0.
(assert_return (invoke "rem_s" (i32.const -2147483648) (i32.const -1)) (i32.const 0))
(assert_return (invoke "rem_s64" (i64.const -9223372036854775808) (i64.const -1)) (i64.const 0))

;; div_u reads the SAME bits as unsigned, where neither operand is special:
;; 0x80000000 / 0xFFFFFFFF = 2147483648 / 4294967295 = 0.
(assert_return (invoke "div_u" (i32.const -2147483648) (i32.const -1)) (i32.const 0))
(assert_return (invoke "rem_u" (i32.const -2147483648) (i32.const -1)) (i32.const -2147483648))

;; ── truncz: toward zero, in all four sign combinations ─────────────────────
(assert_return (invoke "div_s" (i32.const 7) (i32.const 2)) (i32.const 3))
(assert_return (invoke "div_s" (i32.const -7) (i32.const 2)) (i32.const -3))
(assert_return (invoke "div_s" (i32.const 7) (i32.const -2)) (i32.const -3))
(assert_return (invoke "div_s" (i32.const -7) (i32.const -2)) (i32.const 3))

;; The remainder carries the sign of the DIVIDEND, not the divisor.
(assert_return (invoke "rem_s" (i32.const 7) (i32.const 2)) (i32.const 1))
(assert_return (invoke "rem_s" (i32.const -7) (i32.const 2)) (i32.const -1))
(assert_return (invoke "rem_s" (i32.const 7) (i32.const -2)) (i32.const 1))
(assert_return (invoke "rem_s" (i32.const -7) (i32.const -2)) (i32.const -1))

;; Unsigned reads the same bit patterns as large positives.
(assert_return (invoke "div_u" (i32.const -1) (i32.const 2)) (i32.const 2147483647))
(assert_return (invoke "rem_u" (i32.const -1) (i32.const 2)) (i32.const 1))
(assert_return (invoke "div_u" (i32.const -1) (i32.const -1)) (i32.const 1))
(assert_return (invoke "div_u64" (i64.const -1) (i64.const 2)) (i64.const 9223372036854775807))
(assert_return (invoke "rem_u64" (i64.const -1) (i64.const 2)) (i64.const 1))

;; ── The identity the spec states, at the signs most likely to break it ─────
(assert_return (invoke "identity_s" (i32.const -7) (i32.const 2)) (i32.const -7))
(assert_return (invoke "identity_s" (i32.const 7) (i32.const -2)) (i32.const 7))
(assert_return (invoke "identity_s" (i32.const -2147483648) (i32.const 3)) (i32.const -2147483648))
(assert_return (invoke "identity_u" (i32.const -1) (i32.const 3)) (i32.const -1))
