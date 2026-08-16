;; vybe-test: wast/wat_spec_memory/bounds_endianness_and_partial_writes
;; vybe-test-mode: run
;;
;; From the spec's memory instructions (core/exec/instructions.rst, "Memory
;; Instructions") and the memory-access rules in core/exec/modules.rst.
;;
;; The spec pins four things that are easy to implement almost-correctly:
;;
;; 1. The effective address is `i + offset` computed WITHOUT wrapping, and the
;;    bound is `ea + N/8 > |mem|` — so the check is against the END of the
;;    access, not its start. An implementation that checks only the base
;;    address accepts a read that runs off the end by three bytes.
;; 2. The dynamic address operand is UNSIGNED. `i32.load` at -1 is an access
;;    at 4294967295, which traps; it is not an access at -1 or at 0.
;; 3. The alignment immediate is a HINT with no semantic effect. A misaligned
;;    access must produce exactly the same value, never a trap.
;; 4. Numbers are stored LITTLE-ENDIAN, and `loadN_s` sign-extends where
;;    `loadN_u` zero-extends the identical bytes.
;;
;; And the one that is a real correctness hazard: a store whose access is
;; partly out of bounds must write NOTHING. The trap is not permitted to leave
;; the in-bounds prefix modified.

(module
  (memory 1)
  (func (export "i32_load") (param i32) (result i32) (i32.load (local.get 0)))
  (func (export "i32_load_off") (param i32) (result i32) (i32.load offset=65536 (local.get 0)))
  (func (export "i32_store") (param i32 i32) (i32.store (local.get 0) (local.get 1)))
  (func (export "i64_load") (param i32) (result i64) (i64.load (local.get 0)))
  (func (export "load8_u") (param i32) (result i32) (i32.load8_u (local.get 0)))
  (func (export "load8_s") (param i32) (result i32) (i32.load8_s (local.get 0)))
  (func (export "load16_u") (param i32) (result i32) (i32.load16_u (local.get 0)))
  (func (export "load16_s") (param i32) (result i32) (i32.load16_s (local.get 0)))
  (func (export "load32_s64") (param i32) (result i64) (i64.load32_s (local.get 0)))
  (func (export "load32_u64") (param i32) (result i64) (i64.load32_u (local.get 0)))
  (func (export "store8") (param i32 i32) (i32.store8 (local.get 0) (local.get 1)))
  ;; Explicitly under-aligned access — the immediate is a hint only.
  (func (export "unaligned_load") (param i32) (result i32)
    (i32.load align=1 (local.get 0)))
  (func (export "unaligned_store") (param i32 i32)
    (i32.store align=1 (local.get 0) (local.get 1)))
  (func (export "size") (result i32) (memory.size))
)

;; ── The bound is the END of the access ────────────────────────────────────
;; One page is 65536 bytes, so the last valid 4-byte load starts at 65532.
(assert_return (invoke "i32_load" (i32.const 65532)) (i32.const 0))
(assert_trap (invoke "i32_load" (i32.const 65533)) "out of bounds memory access")
(assert_trap (invoke "i32_load" (i32.const 65534)) "out of bounds memory access")
(assert_trap (invoke "i32_load" (i32.const 65536)) "out of bounds memory access")
;; A single byte is valid right up to the last address.
(assert_return (invoke "load8_u" (i32.const 65535)) (i32.const 0))
(assert_trap (invoke "load8_u" (i32.const 65536)) "out of bounds memory access")
;; 8-byte accesses shift the boundary accordingly.
(assert_return (invoke "i64_load" (i32.const 65528)) (i64.const 0))
(assert_trap (invoke "i64_load" (i32.const 65529)) "out of bounds memory access")

;; The static offset participates in the same check.
(assert_trap (invoke "i32_load_off" (i32.const 0)) "out of bounds memory access")

;; The address is UNSIGNED: -1 is 4294967295, far past the end.
(assert_trap (invoke "i32_load" (i32.const -1)) "out of bounds memory access")
(assert_trap (invoke "load8_u" (i32.const -1)) "out of bounds memory access")

;; ── Little-endian layout, observed byte by byte ──────────────────────────
(invoke "i32_store" (i32.const 0) (i32.const 0x12345678))
(assert_return (invoke "load8_u" (i32.const 0)) (i32.const 0x78))
(assert_return (invoke "load8_u" (i32.const 1)) (i32.const 0x56))
(assert_return (invoke "load8_u" (i32.const 2)) (i32.const 0x34))
(assert_return (invoke "load8_u" (i32.const 3)) (i32.const 0x12))
(assert_return (invoke "load16_u" (i32.const 0)) (i32.const 0x5678))

;; ── loadN_s vs loadN_u over the SAME bytes ───────────────────────────────
(invoke "store8" (i32.const 16) (i32.const 0xff))
(assert_return (invoke "load8_u" (i32.const 16)) (i32.const 255))
(assert_return (invoke "load8_s" (i32.const 16)) (i32.const -1))
(invoke "store8" (i32.const 17) (i32.const 0x80))
(assert_return (invoke "load8_u" (i32.const 17)) (i32.const 128))
(assert_return (invoke "load8_s" (i32.const 17)) (i32.const -128))
(invoke "i32_store" (i32.const 20) (i32.const 0xffffffff))
(assert_return (invoke "load16_u" (i32.const 20)) (i32.const 65535))
(assert_return (invoke "load16_s" (i32.const 20)) (i32.const -1))
(assert_return (invoke "load32_s64" (i32.const 20)) (i64.const -1))
(assert_return (invoke "load32_u64" (i32.const 20)) (i64.const 4294967295))

;; ── Alignment is a hint: misaligned access works and agrees ──────────────
(invoke "unaligned_store" (i32.const 1) (i32.const 0x0a0b0c0d))
(assert_return (invoke "unaligned_load" (i32.const 1)) (i32.const 0x0a0b0c0d))
(assert_return (invoke "i32_load" (i32.const 1)) (i32.const 0x0a0b0c0d))
(assert_return (invoke "load8_u" (i32.const 1)) (i32.const 0x0d))
(invoke "i32_store" (i32.const 3) (i32.const 0x11223344))
(assert_return (invoke "unaligned_load" (i32.const 3)) (i32.const 0x11223344))

;; ── A trapping store must write NOTHING ─────────────────────────────────
;; Seed the last four bytes with a known pattern, then attempt a 4-byte store
;; at 65534 — two bytes in bounds, two past the end. The trap must leave both
;; in-bounds bytes exactly as they were.
(invoke "i32_store" (i32.const 65532) (i32.const 0x04030201))
(assert_return (invoke "load8_u" (i32.const 65534)) (i32.const 0x03))
(assert_return (invoke "load8_u" (i32.const 65535)) (i32.const 0x04))
(assert_trap (invoke "i32_store" (i32.const 65534) (i32.const 0xffffffff))
             "out of bounds memory access")
(assert_return (invoke "load8_u" (i32.const 65534)) (i32.const 0x03))
(assert_return (invoke "load8_u" (i32.const 65535)) (i32.const 0x04))
(assert_return (invoke "i32_load" (i32.const 65532)) (i32.const 0x04030201))

(assert_return (invoke "size") (i32.const 1))

;; ── memory.grow: returns the OLD size, or -1 without growing ─────────────
(module
  (memory 1 2)
  (func (export "size") (result i32) (memory.size))
  (func (export "grow") (param i32) (result i32) (memory.grow (local.get 0)))
  (func (export "load8") (param i32) (result i32) (i32.load8_u (local.get 0)))
  (func (export "store8") (param i32 i32) (i32.store8 (local.get 0) (local.get 1)))
)
(assert_return (invoke "size") (i32.const 1))
;; Beyond the current size but within the eventual maximum: still a trap now.
(assert_trap (invoke "load8" (i32.const 65536)) "out of bounds memory access")
;; Growing returns the size BEFORE the growth.
(assert_return (invoke "grow" (i32.const 1)) (i32.const 1))
(assert_return (invoke "size") (i32.const 2))
;; The new page is addressable and reads as zero.
(assert_return (invoke "load8" (i32.const 65536)) (i32.const 0))
(invoke "store8" (i32.const 131071) (i32.const 7))
(assert_return (invoke "load8" (i32.const 131071)) (i32.const 7))
(assert_trap (invoke "load8" (i32.const 131072)) "out of bounds memory access")
;; Exceeding the declared maximum answers -1 and does NOT grow.
(assert_return (invoke "grow" (i32.const 1)) (i32.const -1))
(assert_return (invoke "size") (i32.const 2))
;; A failed grow leaves existing contents intact.
(assert_return (invoke "load8" (i32.const 131071)) (i32.const 7))
;; Growing by zero is legal and reports the current size.
(assert_return (invoke "grow" (i32.const 0)) (i32.const 2))
