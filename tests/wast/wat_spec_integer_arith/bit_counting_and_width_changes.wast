;; vybe-test: wast/wat_spec_integer_arith/bit_counting_and_width_changes
;; vybe-test-mode: run
;;
;; From `op-iclz` / `op-ictz` / `op-ipopcnt` and the width-changing operators
;; in core/exec/numerics.rst.
;;
;; The spec pins the one case every hardware instruction disagrees about:
;;
;;   iclz(i) : "all bits are considered leading zeros if i is 0"
;;   ictz(i) : "all bits are considered trailing zeros if i is 0"
;;
;; so `clz(0)` is N, not undefined and not 0 — x86's BSR leaves the register
;; untouched for a zero input, which is where implementations pick up a wrong
;; answer that only shows on that one input.
;;
;; The width changes are the other half: `extendN_s` reinterprets the low N
;; bits as SIGNED and sign-extends, `extend_i32_u` does not, and `wrap_i64`
;; keeps the low 32 bits with no regard for what it discards.

(module
  (func (export "clz") (param i32) (result i32) (i32.clz (local.get 0)))
  (func (export "ctz") (param i32) (result i32) (i32.ctz (local.get 0)))
  (func (export "popcnt") (param i32) (result i32) (i32.popcnt (local.get 0)))
  (func (export "clz64") (param i64) (result i64) (i64.clz (local.get 0)))
  (func (export "ctz64") (param i64) (result i64) (i64.ctz (local.get 0)))
  (func (export "popcnt64") (param i64) (result i64) (i64.popcnt (local.get 0)))

  (func (export "extend8_s") (param i32) (result i32) (i32.extend8_s (local.get 0)))
  (func (export "extend16_s") (param i32) (result i32) (i32.extend16_s (local.get 0)))
  (func (export "extend8_s64") (param i64) (result i64) (i64.extend8_s (local.get 0)))
  (func (export "extend32_s64") (param i64) (result i64) (i64.extend32_s (local.get 0)))
  (func (export "i64_from_i32_s") (param i32) (result i64) (i64.extend_i32_s (local.get 0)))
  (func (export "i64_from_i32_u") (param i32) (result i64) (i64.extend_i32_u (local.get 0)))
  (func (export "wrap") (param i64) (result i32) (i32.wrap_i64 (local.get 0)))
  (func (export "wrap_then_widen") (param i64) (result i64)
    (i64.extend_i32_s (i32.wrap_i64 (local.get 0))))
)

;; ── The zero input, which is the whole point ───────────────────────────────
(assert_return (invoke "clz" (i32.const 0)) (i32.const 32))
(assert_return (invoke "ctz" (i32.const 0)) (i32.const 32))
(assert_return (invoke "popcnt" (i32.const 0)) (i32.const 0))
(assert_return (invoke "clz64" (i64.const 0)) (i64.const 64))
(assert_return (invoke "ctz64" (i64.const 0)) (i64.const 64))
(assert_return (invoke "popcnt64" (i64.const 0)) (i64.const 0))

;; ── Both ends of the range ─────────────────────────────────────────────────
(assert_return (invoke "clz" (i32.const 1)) (i32.const 31))
(assert_return (invoke "clz" (i32.const -1)) (i32.const 0))
(assert_return (invoke "clz" (i32.const -2147483648)) (i32.const 0))
(assert_return (invoke "clz" (i32.const 2147483647)) (i32.const 1))
(assert_return (invoke "ctz" (i32.const 1)) (i32.const 0))
(assert_return (invoke "ctz" (i32.const -1)) (i32.const 0))
(assert_return (invoke "ctz" (i32.const -2147483648)) (i32.const 31))
(assert_return (invoke "popcnt" (i32.const -1)) (i32.const 32))
(assert_return (invoke "popcnt" (i32.const -2147483648)) (i32.const 1))
(assert_return (invoke "popcnt" (i32.const 0x0f0f0f0f)) (i32.const 16))
(assert_return (invoke "clz64" (i64.const 1)) (i64.const 63))
(assert_return (invoke "ctz64" (i64.const -9223372036854775808)) (i64.const 63))
(assert_return (invoke "popcnt64" (i64.const -1)) (i64.const 64))

;; ── extendN_s: the low N bits are read SIGNED ──────────────────────────────
;; 127 is the largest positive 8-bit value; 128 is the smallest negative one.
(assert_return (invoke "extend8_s" (i32.const 127)) (i32.const 127))
(assert_return (invoke "extend8_s" (i32.const 128)) (i32.const -128))
(assert_return (invoke "extend8_s" (i32.const 255)) (i32.const -1))
;; Everything above bit 7 is discarded before the sign is read.
(assert_return (invoke "extend8_s" (i32.const 0x1234ff)) (i32.const -1))
(assert_return (invoke "extend8_s" (i32.const 0)) (i32.const 0))
(assert_return (invoke "extend16_s" (i32.const 32767)) (i32.const 32767))
(assert_return (invoke "extend16_s" (i32.const 32768)) (i32.const -32768))
(assert_return (invoke "extend16_s" (i32.const 65535)) (i32.const -1))
(assert_return (invoke "extend16_s" (i32.const 0x12348000)) (i32.const -32768))
(assert_return (invoke "extend8_s64" (i64.const 255)) (i64.const -1))
(assert_return (invoke "extend32_s64" (i64.const 4294967295)) (i64.const -1))
(assert_return (invoke "extend32_s64" (i64.const 2147483648)) (i64.const -2147483648))
(assert_return (invoke "extend32_s64" (i64.const 2147483647)) (i64.const 2147483647))

;; ── i32 → i64 widening: signed and unsigned differ on exactly the sign bit ─
(assert_return (invoke "i64_from_i32_s" (i32.const -1)) (i64.const -1))
(assert_return (invoke "i64_from_i32_u" (i32.const -1)) (i64.const 4294967295))
(assert_return (invoke "i64_from_i32_s" (i32.const -2147483648)) (i64.const -2147483648))
(assert_return (invoke "i64_from_i32_u" (i32.const -2147483648)) (i64.const 2147483648))
(assert_return (invoke "i64_from_i32_s" (i32.const 2147483647)) (i64.const 2147483647))
(assert_return (invoke "i64_from_i32_u" (i32.const 2147483647)) (i64.const 2147483647))

;; ── wrap_i64 keeps the low 32 bits and never traps ─────────────────────────
(assert_return (invoke "wrap" (i64.const 8589934591)) (i32.const -1))
(assert_return (invoke "wrap" (i64.const -1)) (i32.const -1))
(assert_return (invoke "wrap" (i64.const 4294967296)) (i32.const 0))
(assert_return (invoke "wrap" (i64.const 2147483648)) (i32.const -2147483648))
(assert_return (invoke "wrap" (i64.const 9223372036854775807)) (i32.const -1))
;; Wrapping then re-widening signed is NOT the identity — that is the
;; observable content of "keeps the low 32 bits".
(assert_return (invoke "wrap_then_widen" (i64.const 4294967296)) (i64.const 0))
(assert_return (invoke "wrap_then_widen" (i64.const 2147483648)) (i64.const -2147483648))
