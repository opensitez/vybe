;; vybe-test: wast/wat_bulk_unsigned_bounds/table_fill_huge_count_traps_and_writes_nothing
;; origin: proposals/spec/test/core/table_fill.wast (spec-compliance regression)
;; vybe-test-mode: run
;;
;; `table.fill` carries the same unsigned-count, check-before-write contract as
;; `memory.fill` — a separate instruction with a separate operand path, so it
;; needs its own test rather than inheriting the memory one's verdict.
;;
;; The reference written is a non-null funcref precisely so the no-write half
;; is observable: after the trap, entry 0 must still be null. Clamping the
;; count to 0 makes the fill a no-op that never traps at all.

(module
  (table $t 8 funcref)
  (elem declare func $f)
  (func $f)
  (func (export "fill-oob")
    (table.fill $t (i32.const 0) (ref.func $f) (i32.const 0xfffffff0)))
  (func (export "entry0-is-null") (result i32)
    (ref.is_null (table.get $t (i32.const 0))))
)
(assert_trap (invoke "fill-oob") "out of bounds table access")
(assert_return (invoke "entry0-is-null") (i32.const 1))
