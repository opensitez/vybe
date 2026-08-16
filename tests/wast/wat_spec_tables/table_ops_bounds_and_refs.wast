;; vybe-test: wast/wat_spec_tables/table_ops_bounds_and_refs
;; vybe-test-mode: run
;;
;; From the table instructions and reference instructions in
;; core/exec/instructions.rst.
;;
;; Tables mirror memories but count ELEMENTS rather than bytes, and the same
;; three hazards apply: the bound is the end of the access, a zero-length bulk
;; operation is still checked, and `table.copy` must handle overlap in both
;; directions. On top of that:
;;
;; * `table.grow` returns the OLD size or -1, and the -1 case must not grow.
;;   Its fill value initialises the new slots, so growing with a real funcref
;;   leaves the new entries CALLABLE rather than null.
;; * `ref.is_null` distinguishes a null reference from a present one; a table
;;   slot never read is null.
;; Note: `ref.eq` deliberately does NOT appear here. `func` is not a subtype
;; of `eq`, so comparing two funcrefs with it is a validation error, not a
;; runtime false — reference identity for functions is observed through the
;; table and `call_indirect` instead.

(module
  (type $thunk (func (result i32)))
  (table $t 3 8 funcref)
  (elem declare func $ten $twenty)
  (func $ten (type $thunk) (i32.const 10))
  (func $twenty (type $thunk) (i32.const 20))

  (func (export "size") (result i32) (table.size $t))
  (func (export "grow_null") (param i32) (result i32)
    (table.grow $t (ref.null func) (local.get 0)))
  (func (export "grow_with") (param i32) (result i32)
    (table.grow $t (ref.func $ten) (local.get 0)))
  (func (export "set_ten") (param i32) (table.set $t (local.get 0) (ref.func $ten)))
  (func (export "set_twenty") (param i32) (table.set $t (local.get 0) (ref.func $twenty)))
  (func (export "set_null") (param i32) (table.set $t (local.get 0) (ref.null func)))
  (func (export "is_null") (param i32) (result i32)
    (ref.is_null (table.get $t (local.get 0))))
  (func (export "call") (param i32) (result i32)
    (call_indirect $t (type $thunk) (local.get 0)))
  (func (export "fill") (param i32 i32)
    (table.fill $t (local.get 0) (ref.func $twenty) (local.get 1)))
  (func (export "fill_null") (param i32 i32)
    (table.fill $t (local.get 0) (ref.null func) (local.get 1)))
  (func (export "copy") (param i32 i32 i32)
    (table.copy $t $t (local.get 0) (local.get 1) (local.get 2)))
  (func (export "null_is_null") (result i32) (ref.is_null (ref.null func)))
  (func (export "func_is_null") (result i32) (ref.is_null (ref.func $ten)))
)

;; ── Initial state: declared size, every slot null ───────────────────────
(assert_return (invoke "size") (i32.const 3))
(assert_return (invoke "is_null" (i32.const 0)) (i32.const 1))
(assert_return (invoke "is_null" (i32.const 2)) (i32.const 1))
(assert_trap (invoke "call" (i32.const 0)) "uninitialized element")

;; ── ref.is_null / ref.eq ────────────────────────────────────────────────
(assert_return (invoke "null_is_null") (i32.const 1))
(assert_return (invoke "func_is_null") (i32.const 0))

;; ── set / get ───────────────────────────────────────────────────────────
(invoke "set_ten" (i32.const 0))
(invoke "set_twenty" (i32.const 1))
(assert_return (invoke "is_null" (i32.const 0)) (i32.const 0))
(assert_return (invoke "call" (i32.const 0)) (i32.const 10))
(assert_return (invoke "call" (i32.const 1)) (i32.const 20))
;; Writing null back makes the slot uninitialised again.
(invoke "set_null" (i32.const 1))
(assert_return (invoke "is_null" (i32.const 1)) (i32.const 1))
(assert_trap (invoke "call" (i32.const 1)) "uninitialized element")
(invoke "set_twenty" (i32.const 1))

;; ── Bounds: the last valid index is size-1, and the index is unsigned ───
(assert_trap (invoke "is_null" (i32.const 3)) "out of bounds table access")
(assert_trap (invoke "is_null" (i32.const -1)) "out of bounds table access")
(assert_trap (invoke "set_ten" (i32.const 3)) "out of bounds table access")
(assert_trap (invoke "call" (i32.const 3)) "undefined element")

;; ── grow returns the OLD size; the fill value initialises new slots ─────
(assert_return (invoke "grow_null" (i32.const 1)) (i32.const 3))
(assert_return (invoke "size") (i32.const 4))
(assert_return (invoke "is_null" (i32.const 3)) (i32.const 1))
;; Growing with a real reference leaves the new slots callable.
(assert_return (invoke "grow_with" (i32.const 2)) (i32.const 4))
(assert_return (invoke "size") (i32.const 6))
(assert_return (invoke "is_null" (i32.const 4)) (i32.const 0))
(assert_return (invoke "call" (i32.const 5)) (i32.const 10))
;; Past the declared maximum: -1, and no growth.
(assert_return (invoke "grow_null" (i32.const 3)) (i32.const -1))
(assert_return (invoke "size") (i32.const 6))
;; Growing by zero reports the current size.
(assert_return (invoke "grow_null" (i32.const 0)) (i32.const 6))
;; A failed grow left the existing contents alone.
(assert_return (invoke "call" (i32.const 0)) (i32.const 10))

;; ── fill ────────────────────────────────────────────────────────────────
(invoke "fill" (i32.const 2) (i32.const 2))
(assert_return (invoke "call" (i32.const 2)) (i32.const 20))
(assert_return (invoke "call" (i32.const 3)) (i32.const 20))
;; Exactly the requested length.
(assert_return (invoke "call" (i32.const 4)) (i32.const 10))
;; Out of range traps and writes nothing.
(assert_trap (invoke "fill" (i32.const 5) (i32.const 2)) "out of bounds table access")
(assert_return (invoke "call" (i32.const 5)) (i32.const 10))
;; Zero length at exactly the end is in bounds; one past is not.
(invoke "fill" (i32.const 6) (i32.const 0))
(assert_trap (invoke "fill" (i32.const 7) (i32.const 0)) "out of bounds table access")

;; ── Overlapping copy, both directions ──────────────────────────────────
;; Lay out [ten, twenty, ten, twenty, ten, twenty].
(invoke "set_ten" (i32.const 0))
(invoke "set_twenty" (i32.const 1))
(invoke "set_ten" (i32.const 2))
(invoke "set_twenty" (i32.const 3))
(invoke "set_ten" (i32.const 4))
(invoke "set_twenty" (i32.const 5))
;; dst above src: slots 0,1,2 -> 2,3,4 should read ten,twenty,ten.
(invoke "copy" (i32.const 2) (i32.const 0) (i32.const 3))
(assert_return (invoke "call" (i32.const 2)) (i32.const 10))
(assert_return (invoke "call" (i32.const 3)) (i32.const 20))
(assert_return (invoke "call" (i32.const 4)) (i32.const 10))
;; dst below src.
(invoke "set_twenty" (i32.const 4))
(invoke "set_ten" (i32.const 5))
(invoke "copy" (i32.const 3) (i32.const 4) (i32.const 2))
(assert_return (invoke "call" (i32.const 3)) (i32.const 20))
(assert_return (invoke "call" (i32.const 4)) (i32.const 10))
;; Out of range traps.
(assert_trap (invoke "copy" (i32.const 5) (i32.const 0) (i32.const 3))
             "out of bounds table access")
(assert_trap (invoke "copy" (i32.const 0) (i32.const 5) (i32.const 3))
             "out of bounds table access")
;; Zero length at the end is fine.
(invoke "copy" (i32.const 6) (i32.const 6) (i32.const 0))
