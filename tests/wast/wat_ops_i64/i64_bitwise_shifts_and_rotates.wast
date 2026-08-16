;; vybe-test: wast/wat_ops_i64/i64_bitwise_shifts_and_rotates
;; origin: coverage gap vs crates/vybe_runtime/tests/i64_ops_test.rs
;; vybe-test-mode: run
;;
;; i64 and/or/xor/shr_s/shr_u/rotr at 64-bit width and across the sign bit.
;;
;; Two properties no single-assertion test reaches:
;;
;;   * Shift counts are taken MODULO 64 (spec §4.3.3), not clamped. `shr_u` by
;;     64 is a shift by 0 — the identity — and by 65 is a shift by 1. An
;;     implementation that clamps returns 0 for both.
;;   * `shr_s` copies the sign bit in, `shr_u` brings in zeros. They differ
;;     only on a negative operand, so a positive-only test cannot separate
;;     them.
;;
;; Spec-format so `wasmtime wast` checks these expectations independently.

(module
  (func (export "and") (param i64 i64) (result i64) (i64.and (local.get 0) (local.get 1)))
  (func (export "or")  (param i64 i64) (result i64) (i64.or  (local.get 0) (local.get 1)))
  (func (export "xor") (param i64 i64) (result i64) (i64.xor (local.get 0) (local.get 1)))
  (func (export "shr_s") (param i64 i64) (result i64) (i64.shr_s (local.get 0) (local.get 1)))
  (func (export "shr_u") (param i64 i64) (result i64) (i64.shr_u (local.get 0) (local.get 1)))
  (func (export "rotr")  (param i64 i64) (result i64) (i64.rotr  (local.get 0) (local.get 1)))
)

;; ── and / or / xor across the full 64-bit width ─────────────────────────
(assert_return (invoke "and" (i64.const 0xffffffffffffffff) (i64.const 0x0f0f0f0f0f0f0f0f)) (i64.const 0x0f0f0f0f0f0f0f0f))
(assert_return (invoke "and" (i64.const 0xffffffff00000000) (i64.const 0x00000000ffffffff)) (i64.const 0))
(assert_return (invoke "or"  (i64.const 0xffffffff00000000) (i64.const 0x00000000ffffffff)) (i64.const -1))
(assert_return (invoke "xor" (i64.const 0xffffffffffffffff) (i64.const 0xffffffffffffffff)) (i64.const 0))
(assert_return (invoke "xor" (i64.const 0xaaaaaaaaaaaaaaaa) (i64.const 0x5555555555555555)) (i64.const -1))
;; Identities — an operation that truncated to 32 bits fails the high half.
(assert_return (invoke "and" (i64.const 0x123456789abcdef0) (i64.const -1)) (i64.const 0x123456789abcdef0))
(assert_return (invoke "or"  (i64.const 0x123456789abcdef0) (i64.const 0)) (i64.const 0x123456789abcdef0))
(assert_return (invoke "xor" (i64.const 0x123456789abcdef0) (i64.const 0)) (i64.const 0x123456789abcdef0))

;; ── shr_s vs shr_u on a NEGATIVE operand ────────────────────────────────
;; -1 arithmetic-shifted stays -1 forever; logically shifted it becomes a
;; large positive. Same operands, opposite results.
(assert_return (invoke "shr_s" (i64.const -1) (i64.const 1))  (i64.const -1))
(assert_return (invoke "shr_u" (i64.const -1) (i64.const 1))  (i64.const 0x7fffffffffffffff))
(assert_return (invoke "shr_s" (i64.const -1) (i64.const 63)) (i64.const -1))
(assert_return (invoke "shr_u" (i64.const -1) (i64.const 63)) (i64.const 1))
(assert_return (invoke "shr_s" (i64.const 0x8000000000000000) (i64.const 63)) (i64.const -1))
(assert_return (invoke "shr_u" (i64.const 0x8000000000000000) (i64.const 63)) (i64.const 1))
;; Positive operand: the two agree, so this is the case a sign-confused
;; implementation still passes.
(assert_return (invoke "shr_s" (i64.const 0x40) (i64.const 2)) (i64.const 0x10))
(assert_return (invoke "shr_u" (i64.const 0x40) (i64.const 2)) (i64.const 0x10))

;; ── shift counts are taken mod 64, never clamped ────────────────────────
(assert_return (invoke "shr_u" (i64.const -1) (i64.const 64)) (i64.const -1))
(assert_return (invoke "shr_u" (i64.const -1) (i64.const 65)) (i64.const 0x7fffffffffffffff))
(assert_return (invoke "shr_s" (i64.const -1) (i64.const 64)) (i64.const -1))
;; A negative count is read as unsigned, so -1 is 63 mod 64.
(assert_return (invoke "shr_u" (i64.const -1) (i64.const -1)) (i64.const 1))

;; ── rotr wraps bits round rather than discarding them ───────────────────
(assert_return (invoke "rotr" (i64.const 1) (i64.const 1))  (i64.const 0x8000000000000000))
(assert_return (invoke "rotr" (i64.const 1) (i64.const 64)) (i64.const 1))
(assert_return (invoke "rotr" (i64.const -1) (i64.const 7)) (i64.const -1))
(assert_return (invoke "rotr" (i64.const 0x0123456789abcdef) (i64.const 4)) (i64.const 0xf0123456789abcde))
