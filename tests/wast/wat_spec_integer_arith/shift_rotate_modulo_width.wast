;; vybe-test: wast/wat_spec_integer_arith/shift_rotate_modulo_width
;; vybe-test-mode: run
;;
;; From the spec's shift and rotate definitions (core/exec/numerics.rst,
;; `op-ishl` / `op-ishr` / `op-irotl` / `op-irotr`). Every one of the five
;; begins with the SAME sentence:
;;
;;   Let k be i2 modulo N.
;;
;; So a shift count is never out of range — it wraps. A host that lets the
;; machine instruction decide gets this right on x86 by accident (which masks
;; by 31/63) and wrong on other targets; a host that clamps large counts to
;; "all bits shifted out" is wrong everywhere. Both mistakes are invisible
;; unless the count actually exceeds the width, which is what this file does.
;;
;; The count is also read as UNSIGNED for the modulo: a count of -1 is
;; 4294967295, and 4294967295 mod 32 = 31.
;;
;; shr_s extends with the most significant bit of the ORIGINAL value, so it
;; saturates toward -1 or 0 rather than reaching the other sign.

(module
  (func (export "shl") (param i32 i32) (result i32) (i32.shl (local.get 0) (local.get 1)))
  (func (export "shr_s") (param i32 i32) (result i32) (i32.shr_s (local.get 0) (local.get 1)))
  (func (export "shr_u") (param i32 i32) (result i32) (i32.shr_u (local.get 0) (local.get 1)))
  (func (export "rotl") (param i32 i32) (result i32) (i32.rotl (local.get 0) (local.get 1)))
  (func (export "rotr") (param i32 i32) (result i32) (i32.rotr (local.get 0) (local.get 1)))
  (func (export "shl64") (param i64 i64) (result i64) (i64.shl (local.get 0) (local.get 1)))
  (func (export "shr_s64") (param i64 i64) (result i64) (i64.shr_s (local.get 0) (local.get 1)))
  (func (export "shr_u64") (param i64 i64) (result i64) (i64.shr_u (local.get 0) (local.get 1)))
  (func (export "rotl64") (param i64 i64) (result i64) (i64.rotl (local.get 0) (local.get 1)))
  (func (export "rotr64") (param i64 i64) (result i64) (i64.rotr (local.get 0) (local.get 1)))
)

;; ── k = i2 mod N: a count of exactly N is a shift of ZERO ──────────────────
(assert_return (invoke "shl" (i32.const 1) (i32.const 32)) (i32.const 1))
(assert_return (invoke "shl" (i32.const 1) (i32.const 33)) (i32.const 2))
(assert_return (invoke "shl" (i32.const 1) (i32.const 64)) (i32.const 1))
(assert_return (invoke "shr_u" (i32.const -1) (i32.const 32)) (i32.const -1))
(assert_return (invoke "shr_s" (i32.const -2147483648) (i32.const 32)) (i32.const -2147483648))
(assert_return (invoke "rotl" (i32.const 1) (i32.const 32)) (i32.const 1))
(assert_return (invoke "rotr" (i32.const 1) (i32.const 32)) (i32.const 1))
(assert_return (invoke "shl64" (i64.const 1) (i64.const 64)) (i64.const 1))
(assert_return (invoke "shl64" (i64.const 1) (i64.const 65)) (i64.const 2))
(assert_return (invoke "shr_u64" (i64.const -1) (i64.const 64)) (i64.const -1))
(assert_return (invoke "rotl64" (i64.const 1) (i64.const 64)) (i64.const 1))

;; A NEGATIVE count is read unsigned before the modulo:
;; -1 is 4294967295, and 4294967295 mod 32 = 31.
(assert_return (invoke "shl" (i32.const 1) (i32.const -1)) (i32.const -2147483648))
(assert_return (invoke "shr_u" (i32.const -2147483648) (i32.const -1)) (i32.const 1))
;; 18446744073709551615 mod 64 = 63.
(assert_return (invoke "shl64" (i64.const 1) (i64.const -1)) (i64.const -9223372036854775808))

;; ── In-range shifts, at the bit that leaves the value ──────────────────────
(assert_return (invoke "shl" (i32.const 1) (i32.const 31)) (i32.const -2147483648))
(assert_return (invoke "shl" (i32.const -1) (i32.const 1)) (i32.const -2))
(assert_return (invoke "shl" (i32.const -2147483648) (i32.const 1)) (i32.const 0))

;; shr_u zero-fills, so a negative operand becomes a large positive.
(assert_return (invoke "shr_u" (i32.const -1) (i32.const 1)) (i32.const 2147483647))
(assert_return (invoke "shr_u" (i32.const -2147483648) (i32.const 31)) (i32.const 1))

;; shr_s fills with the ORIGINAL sign bit, so it saturates at -1, never 0.
(assert_return (invoke "shr_s" (i32.const -1) (i32.const 1)) (i32.const -1))
(assert_return (invoke "shr_s" (i32.const -1) (i32.const 31)) (i32.const -1))
(assert_return (invoke "shr_s" (i32.const -2147483648) (i32.const 31)) (i32.const -1))
(assert_return (invoke "shr_s" (i32.const 2147483647) (i32.const 31)) (i32.const 0))
(assert_return (invoke "shr_s64" (i64.const -1) (i64.const 63)) (i64.const -1))
(assert_return (invoke "shr_u64" (i64.const -1) (i64.const 63)) (i64.const 1))

;; ── Rotates move bits around the end rather than discarding them ───────────
(assert_return (invoke "rotl" (i32.const 0x12345678) (i32.const 4)) (i32.const 0x23456781))
(assert_return (invoke "rotr" (i32.const 0x12345678) (i32.const 4)) (i32.const 0x81234567))
;; The bit that a shift would have lost is exactly the bit a rotate keeps.
(assert_return (invoke "rotl" (i32.const -2147483648) (i32.const 1)) (i32.const 1))
(assert_return (invoke "rotr" (i32.const 1) (i32.const 1)) (i32.const -2147483648))
(assert_return (invoke "rotl" (i32.const -1) (i32.const 13)) (i32.const -1))
(assert_return (invoke "rotl64" (i64.const -9223372036854775808) (i64.const 1)) (i64.const 1))
(assert_return (invoke "rotr64" (i64.const 1) (i64.const 1)) (i64.const -9223372036854775808))

;; rotl by k and rotr by N-k are the same rotation.
(assert_return (invoke "rotl" (i32.const 0x12345678) (i32.const 12)) (i32.const 0x45678123))
(assert_return (invoke "rotr" (i32.const 0x12345678) (i32.const 20)) (i32.const 0x45678123))
