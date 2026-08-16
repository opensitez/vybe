;; vybe-test: wast/wat_bulk_unsigned_bounds/table_copy_huge_count_traps_and_writes_nothing
;; origin: proposals/spec/test/core/table_copy.wast (spec-compliance regression)
;; vybe-test-mode: run
;;
;; `table.copy` reads THREE unsigned operands — destination, source, count —
;; and traps unless BOTH ranges fit, with nothing copied. The count is the one
;; a signed read turns negative, and a clamp to 0 makes the whole instruction
;; disappear without a trap.
;;
;; Source entry 4 holds a function and destination entry 0 is null, so a copy
;; that ran before checking is visible as entry 0 becoming non-null.

(module
  (table $t 8 funcref)
  (elem declare func $f)
  (func $f)
  (func (export "seed")
    (table.set $t (i32.const 4) (ref.func $f)))
  (func (export "copy-oob")
    (table.copy $t $t (i32.const 0) (i32.const 4) (i32.const 0xfffffff0)))
  (func (export "entry0-is-null") (result i32)
    (ref.is_null (table.get $t (i32.const 0))))
)
(assert_return (invoke "seed"))
(assert_trap (invoke "copy-oob") "out of bounds table access")
(assert_return (invoke "entry0-is-null") (i32.const 1))
