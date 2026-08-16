;; vybe-test: wast/wat_spec_memory/bulk_operations_overlap_and_bounds
;; vybe-test-mode: run
;;
;; From the bulk memory/table instructions in core/exec/instructions.rst.
;;
;; Three rules, each with a specific way of going wrong:
;;
;; 1. `memory.copy` is defined to behave correctly when source and destination
;;    OVERLAP, in either direction — it is memmove, not memcpy. A forward
;;    byte loop gives the wrong answer when dst > src, which is exactly the
;;    case a naive implementation gets wrong and a non-overlapping test never
;;    reaches.
;; 2. Bounds are checked BEFORE anything is written. A copy/fill/init that
;;    runs off the end traps and leaves memory completely untouched — not
;;    partially applied up to the boundary.
;; 3. A ZERO-length operation at an otherwise out-of-bounds offset is still
;;    bounds-checked: `n = 0` at exactly the end is fine, one past the end
;;    traps. This is the boundary case where implementations tend to
;;    short-circuit on `n == 0` and skip the check.
;;
;; `data.drop` / `elem.drop` then make the segment behave as empty: a later
;; `memory.init` from a dropped segment traps for any non-zero length.

(module
  (memory 1)
  (data $seg "ABCDEFGH")
  (data $other "xyz")
  (func (export "fill") (param i32 i32 i32) (memory.fill (local.get 0) (local.get 1) (local.get 2)))
  (func (export "copy") (param i32 i32 i32) (memory.copy (local.get 0) (local.get 1) (local.get 2)))
  (func (export "init") (param i32 i32 i32) (memory.init $seg (local.get 0) (local.get 1) (local.get 2)))
  (func (export "drop_seg") (data.drop $seg))
  (func (export "load8") (param i32) (result i32) (i32.load8_u (local.get 0)))
  (func (export "store8") (param i32 i32) (i32.store8 (local.get 0) (local.get 1)))
)

;; Seed 0..7 with 1,2,3,4,5,6,7,8.
(invoke "store8" (i32.const 0) (i32.const 1))
(invoke "store8" (i32.const 1) (i32.const 2))
(invoke "store8" (i32.const 2) (i32.const 3))
(invoke "store8" (i32.const 3) (i32.const 4))
(invoke "store8" (i32.const 4) (i32.const 5))
(invoke "store8" (i32.const 5) (i32.const 6))
(invoke "store8" (i32.const 6) (i32.const 7))
(invoke "store8" (i32.const 7) (i32.const 8))

;; ── Overlapping copy, destination ABOVE source ───────────────────────────
;; copy(dst=2, src=0, n=4): bytes 0..3 (1,2,3,4) land at 2..5. A forward loop
;; would read back what it just wrote and produce 1,2,1,2 instead.
(invoke "copy" (i32.const 2) (i32.const 0) (i32.const 4))
(assert_return (invoke "load8" (i32.const 2)) (i32.const 1))
(assert_return (invoke "load8" (i32.const 3)) (i32.const 2))
(assert_return (invoke "load8" (i32.const 4)) (i32.const 3))
(assert_return (invoke "load8" (i32.const 5)) (i32.const 4))
;; Bytes outside the destination range are untouched.
(assert_return (invoke "load8" (i32.const 0)) (i32.const 1))
(assert_return (invoke "load8" (i32.const 1)) (i32.const 2))
(assert_return (invoke "load8" (i32.const 6)) (i32.const 7))

;; ── Overlapping copy, destination BELOW source ───────────────────────────
;; Reseed, then copy(dst=0, src=2, n=4).
(invoke "store8" (i32.const 0) (i32.const 1))
(invoke "store8" (i32.const 1) (i32.const 2))
(invoke "store8" (i32.const 2) (i32.const 3))
(invoke "store8" (i32.const 3) (i32.const 4))
(invoke "store8" (i32.const 4) (i32.const 5))
(invoke "store8" (i32.const 5) (i32.const 6))
(invoke "copy" (i32.const 0) (i32.const 2) (i32.const 4))
(assert_return (invoke "load8" (i32.const 0)) (i32.const 3))
(assert_return (invoke "load8" (i32.const 1)) (i32.const 4))
(assert_return (invoke "load8" (i32.const 2)) (i32.const 5))
(assert_return (invoke "load8" (i32.const 3)) (i32.const 6))

;; A copy with src == dst is a no-op, not a corruption.
(invoke "copy" (i32.const 0) (i32.const 0) (i32.const 4))
(assert_return (invoke "load8" (i32.const 0)) (i32.const 3))
(assert_return (invoke "load8" (i32.const 3)) (i32.const 6))

;; ── fill ─────────────────────────────────────────────────────────────────
(invoke "fill" (i32.const 100) (i32.const 0xab) (i32.const 3))
(assert_return (invoke "load8" (i32.const 100)) (i32.const 0xab))
(assert_return (invoke "load8" (i32.const 102)) (i32.const 0xab))
;; Exactly the requested length, no more.
(assert_return (invoke "load8" (i32.const 103)) (i32.const 0))
;; Only the low byte of the value is used.
(invoke "fill" (i32.const 110) (i32.const 0x1234ff) (i32.const 2))
(assert_return (invoke "load8" (i32.const 110)) (i32.const 0xff))

;; ── memory.init from a data segment ─────────────────────────────────────
(invoke "init" (i32.const 200) (i32.const 0) (i32.const 8))
(assert_return (invoke "load8" (i32.const 200)) (i32.const 65))  ;; 'A'
(assert_return (invoke "load8" (i32.const 207)) (i32.const 72))  ;; 'H'
;; A slice from the middle of the segment.
(invoke "init" (i32.const 210) (i32.const 2) (i32.const 3))
(assert_return (invoke "load8" (i32.const 210)) (i32.const 67))  ;; 'C'
(assert_return (invoke "load8" (i32.const 212)) (i32.const 69))  ;; 'E'

;; ── Bounds are checked BEFORE any write ─────────────────────────────────
;; A fill running one byte past the end must leave everything alone.
(invoke "store8" (i32.const 65535) (i32.const 42))
(assert_trap (invoke "fill" (i32.const 65534) (i32.const 0xcd) (i32.const 3))
             "out of bounds memory access")
(assert_return (invoke "load8" (i32.const 65535)) (i32.const 42))
(assert_return (invoke "load8" (i32.const 65534)) (i32.const 0))

;; Same for copy and init.
(assert_trap (invoke "copy" (i32.const 65534) (i32.const 0) (i32.const 4))
             "out of bounds memory access")
(assert_return (invoke "load8" (i32.const 65535)) (i32.const 42))
(assert_trap (invoke "init" (i32.const 65534) (i32.const 0) (i32.const 8))
             "out of bounds memory access")
(assert_return (invoke "load8" (i32.const 65535)) (i32.const 42))
;; Reading past the END of the data segment traps too.
(assert_trap (invoke "init" (i32.const 0) (i32.const 0) (i32.const 9))
             "out of bounds memory access")
(assert_trap (invoke "init" (i32.const 0) (i32.const 8) (i32.const 1))
             "out of bounds memory access")

;; ── Zero length is still bounds-checked ────────────────────────────────
;; n = 0 exactly AT the end is in bounds...
(invoke "fill" (i32.const 65536) (i32.const 0) (i32.const 0))
(invoke "copy" (i32.const 65536) (i32.const 0) (i32.const 0))
;; ...but one past it is not, even with nothing to write.
(assert_trap (invoke "fill" (i32.const 65537) (i32.const 0) (i32.const 0))
             "out of bounds memory access")
(assert_trap (invoke "copy" (i32.const 65537) (i32.const 0) (i32.const 0))
             "out of bounds memory access")
;; A zero-length init at the very end of the segment is fine.
(invoke "init" (i32.const 0) (i32.const 8) (i32.const 0))

;; ── data.drop makes the segment behave as empty ────────────────────────
(invoke "drop_seg")
;; Zero length from a dropped segment at offset 0 is still allowed.
(invoke "init" (i32.const 0) (i32.const 0) (i32.const 0))
;; Any actual read from it traps.
(assert_trap (invoke "init" (i32.const 0) (i32.const 0) (i32.const 1))
             "out of bounds memory access")
;; Dropping is idempotent, not an error.
(invoke "drop_seg")
