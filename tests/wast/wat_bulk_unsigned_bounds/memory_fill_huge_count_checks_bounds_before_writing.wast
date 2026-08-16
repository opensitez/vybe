;; vybe-test: wast/wat_bulk_unsigned_bounds/memory_fill_huge_count_checks_bounds_before_writing
;; origin: proposals/spec/test/core/memory_fill.wast (spec-compliance regression)
;; vybe-test-mode: run
;;
;; The COUNT is unsigned too, and the bounds check has to happen BEFORE any
;; byte moves. `dst + n > |mem|` is decidable from the operands alone.
;;
;; This is the ORDERING test, which is why the destination is 0 — perfectly in
;; bounds — and only the length runs off the end. An implementation that fills
;; first and discovers the overflow afterwards traps just the same, so a
;; trap-only assertion passes it; what it cannot do is leave address 0 at zero.
;; It also has to build a 4-gigabyte buffer to get there, turning a trap into
;; an allocation.

(module
  (memory 1)
  (func (export "fill-oob")
    (memory.fill (i32.const 0) (i32.const 7) (i32.const 0xfffffff0)))
  (func (export "byte0") (result i32)
    (i32.load8_u (i32.const 0)))
)
(assert_trap (invoke "fill-oob") "out of bounds memory access")
(assert_return (invoke "byte0") (i32.const 0))
