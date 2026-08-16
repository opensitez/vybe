;; vybe-test: wast/wat_ops_memory/narrow_i64_load_store_widths
;; origin: coverage gap — i64.store8/16, i64.load16_s and f32.store occurred ONCE in the run corpus
;; vybe-test-mode: run
;;
;; The narrow i64 accessors: `store8`/`store16`/`store32` write only the LOW
;; bytes of the operand, and `load8`/`load16`/`load32` come in signed and
;; unsigned forms that differ only when the stored byte has its top bit set.
;;
;; Two properties a single round-trip of a small positive value cannot reach:
;;
;;   * a narrow store TRUNCATES rather than trapping or widening — storing
;;     0x1122334455667788 with `store8` must leave 0x88 and touch NO other
;;     byte, which the neighbouring-byte reads below check directly;
;;   * `_s` sign-extends to the full 64 bits and `_u` zero-extends, so 0x88
;;     loads as -120 or 136 depending only on the opcode.
;;
;; Everything is read back through `i64.load8_u` at explicit offsets so a store
;; that wrote too WIDE is caught, not just one that wrote the wrong value.

(module
  (memory 1)
  (func (export "store8") (param i32 i64) (i64.store8 (local.get 0) (local.get 1)))
  (func (export "store16") (param i32 i64) (i64.store16 (local.get 0) (local.get 1)))
  (func (export "store32") (param i32 i64) (i64.store32 (local.get 0) (local.get 1)))
  (func (export "store64") (param i32 i64) (i64.store (local.get 0) (local.get 1)))
  (func (export "storef32") (param i32 f32) (f32.store (local.get 0) (local.get 1)))
  (func (export "load8_s") (param i32) (result i64) (i64.load8_s (local.get 0)))
  (func (export "load8_u") (param i32) (result i64) (i64.load8_u (local.get 0)))
  (func (export "load16_s") (param i32) (result i64) (i64.load16_s (local.get 0)))
  (func (export "load16_u") (param i32) (result i64) (i64.load16_u (local.get 0)))
  (func (export "load32_s") (param i32) (result i64) (i64.load32_s (local.get 0)))
  (func (export "load32_u") (param i32) (result i64) (i64.load32_u (local.get 0)))
  (func (export "load64") (param i32) (result i64) (i64.load (local.get 0)))
  (func (export "loadf32") (param i32) (result f32) (f32.load (local.get 0)))
  (func (export "zero16") (param i32) (i64.store (local.get 0) (i64.const 0)))
)

;; ── store8 writes ONE byte and truncates the rest ───────────────────────
(invoke "store64" (i32.const 0) (i64.const 0))
(invoke "store8" (i32.const 0) (i64.const 0x1122334455667788))
(assert_return (invoke "load8_u" (i32.const 0)) (i64.const 0x88))
;; The next seven bytes were NOT touched — a store that wrote 8 bytes fails here.
(assert_return (invoke "load8_u" (i32.const 1)) (i64.const 0))
(assert_return (invoke "load8_u" (i32.const 7)) (i64.const 0))
(assert_return (invoke "load64" (i32.const 0)) (i64.const 0x88))

;; ── signed vs unsigned narrow loads ─────────────────────────────────────
;; 0x88 has its top bit set: -120 signed, 136 unsigned. Same byte, same address.
(assert_return (invoke "load8_s" (i32.const 0)) (i64.const -120))
(assert_return (invoke "load8_u" (i32.const 0)) (i64.const 136))
;; A byte WITHOUT the top bit set agrees between the two — the case that
;; cannot tell a broken sign extension from a working one.
(invoke "store8" (i32.const 0) (i64.const 0x7f))
(assert_return (invoke "load8_s" (i32.const 0)) (i64.const 127))
(assert_return (invoke "load8_u" (i32.const 0)) (i64.const 127))

;; ── store16 / load16 ────────────────────────────────────────────────────
(invoke "store64" (i32.const 8) (i64.const 0))
(invoke "store16" (i32.const 8) (i64.const 0x1122334455667788))
(assert_return (invoke "load16_u" (i32.const 8)) (i64.const 0x7788))
(assert_return (invoke "load8_u" (i32.const 8)) (i64.const 0x88))
(assert_return (invoke "load8_u" (i32.const 9)) (i64.const 0x77))
(assert_return (invoke "load8_u" (i32.const 10)) (i64.const 0))
;; 0x8000 is negative as an i16 and positive as a u16.
(invoke "store16" (i32.const 8) (i64.const 0x8000))
(assert_return (invoke "load16_s" (i32.const 8)) (i64.const -32768))
(assert_return (invoke "load16_u" (i32.const 8)) (i64.const 32768))
(invoke "store16" (i32.const 8) (i64.const 0xffff))
(assert_return (invoke "load16_s" (i32.const 8)) (i64.const -1))
(assert_return (invoke "load16_u" (i32.const 8)) (i64.const 65535))

;; ── store32 / load32 ────────────────────────────────────────────────────
(invoke "store64" (i32.const 16) (i64.const 0))
(invoke "store32" (i32.const 16) (i64.const 0x1122334455667788))
(assert_return (invoke "load32_u" (i32.const 16)) (i64.const 0x55667788))
(assert_return (invoke "load8_u" (i32.const 20)) (i64.const 0))
(invoke "store32" (i32.const 16) (i64.const 0xffffffff))
(assert_return (invoke "load32_s" (i32.const 16)) (i64.const -1))
(assert_return (invoke "load32_u" (i32.const 16)) (i64.const 4294967295))
;; The high half of memory is still zero: store32 sign-extended nothing.
(assert_return (invoke "load64" (i32.const 16)) (i64.const 0xffffffff))

;; ── f32.store writes exactly four bytes, in little-endian order ─────────
(invoke "store64" (i32.const 24) (i64.const 0))
(invoke "storef32" (i32.const 24) (f32.const 1.0))
(assert_return (invoke "loadf32" (i32.const 24)) (f32.const 1.0))
(assert_return (invoke "load32_u" (i32.const 24)) (i64.const 0x3f800000))
(assert_return (invoke "load8_u" (i32.const 24)) (i64.const 0x00))
(assert_return (invoke "load8_u" (i32.const 26)) (i64.const 0x80))
(assert_return (invoke "load8_u" (i32.const 27)) (i64.const 0x3f))
(assert_return (invoke "load8_u" (i32.const 28)) (i64.const 0))
;; A NaN goes through memory as its exact bit pattern — memory does not quiet.
(invoke "storef32" (i32.const 24) (f32.const -nan))
(assert_return (invoke "load32_u" (i32.const 24)) (i64.const 0xffc00000))
;; The QUIET NaN above cannot detect a store that round-trips through f64 —
;; widening only sets a bit that is already set. Only a SIGNALLING NaN can,
;; which is the whole point of `float_memory.wast` ("load and store do not
;; canonicalize NaNs"). Store then reload as f32 too, so a load that quiets is
;; caught as well as a store that does.
(invoke "storef32" (i32.const 24) (f32.const nan:0x200000))
(assert_return (invoke "load32_u" (i32.const 24)) (i64.const 0x7fa00000))
(assert_return (invoke "loadf32" (i32.const 24)) (f32.const nan:0x200000))

;; ── out-of-bounds traps, at the width actually accessed ─────────────────
;; The last valid i64 address in a one-page memory is 65528; a narrow store
;; reaches further because it writes fewer bytes.
(assert_trap (invoke "store64" (i32.const 65536) (i64.const 1)) "out of bounds memory access")
(assert_trap (invoke "store8" (i32.const 65536) (i64.const 1)) "out of bounds memory access")
(invoke "store8" (i32.const 65535) (i64.const 0x5a))
(assert_return (invoke "load8_u" (i32.const 65535)) (i64.const 0x5a))
(assert_trap (invoke "load16_u" (i32.const 65535)) "out of bounds memory access")
