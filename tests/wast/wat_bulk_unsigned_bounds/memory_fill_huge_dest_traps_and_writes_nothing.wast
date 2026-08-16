;; vybe-test: wast/wat_bulk_unsigned_bounds/memory_fill_huge_dest_traps_and_writes_nothing
;; origin: proposals/spec/test/core/memory_fill.wast (spec-compliance regression)
;; vybe-test-mode: run
;;
;; A `memory.fill` destination is an UNSIGNED address: 0xfffffff0 is
;; 4294967280, far past a one-page memory, so the instruction traps.
;;
;; Read as a signed -16 and clamped to 0, the very same instruction became a
;; legal one-byte write at address 0 — an out-of-bounds store silently
;; redirected to the start of memory. So "did it trap" is only half the
;; question; the second `assert_return` is the half that catches the clamp,
;; because a clamped fill traps nowhere and leaves 7 at address 0.

(module
  (memory 1)
  (func (export "fill-oob")
    (memory.fill (i32.const 0xfffffff0) (i32.const 7) (i32.const 1)))
  (func (export "byte0") (result i32)
    (i32.load8_u (i32.const 0)))
)
(assert_trap (invoke "fill-oob") "out of bounds memory access")
(assert_return (invoke "byte0") (i32.const 0))
