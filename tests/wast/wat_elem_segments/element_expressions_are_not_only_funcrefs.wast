;; vybe-test: wast/wat_elem_segments/element_expressions_are_not_only_funcrefs
;; vybe-test-mode: run
;;
;; A WASM 3.0 element segment holds arbitrary element EXPRESSIONS of its
;; reference type — `(item (ref.i31 (i32.const 999)))`, `(item (global.get
;; $g))`, `(item (struct.new $t …))` — not only `ref.func`.
;;
;; The walker reduced every item to a function NAME, so anything that was not
;; a `ref.func` or a `ref.null` contributed nothing: the slot kept the table's
;; default and the read answered a null. `gc/i31.wast` fills an `i31ref` table
;; that way and every one of its reads came back wrong.
;;
;; The funcref spellings are exercised alongside, because that path is the one
;; that already worked and a general fix must not cost it.

(module
  (table $t 4 i31ref)
  (elem (table $t) (i32.const 0) i31ref
    (item (ref.i31 (i32.const 999)))
    (item (ref.i31 (i32.const 888)))
    (item (ref.i31 (i32.const 777))))

  (func (export "get") (param $i i32) (result i32)
    (i31.get_u (table.get $t (local.get $i))))
  (func (export "is_null") (param $i i32) (result i32)
    (ref.is_null (table.get $t (local.get $i)))))

(assert_return (invoke "get" (i32.const 0)) (i32.const 999))
(assert_return (invoke "get" (i32.const 1)) (i32.const 888))
(assert_return (invoke "get" (i32.const 2)) (i32.const 777))
;; Slot 3 was never written — the segment covers three slots, not the table.
(assert_return (invoke "is_null" (i32.const 3)) (i32.const 1))

;; ── A segment mixing nulls, funcrefs and general expressions ────────
;; Every item occupies its slot whatever its form, so a spelling that
;; contributes no value must still ADVANCE the index. Dropping one shifts
;; everything after it down, which is the shape that made an earlier fix
;; necessary for `ref.null`.
(module
  (type $r (func (result i32)))
  (table $t 5 funcref)
  (func $a (result i32) (i32.const 11))
  (func $b (result i32) (i32.const 22))
  (elem (table $t) (i32.const 0) funcref
    (item (ref.func $a))
    (item (ref.null func))
    (item (ref.func $b)))

  (func (export "call") (param $i i32) (result i32)
    (call_indirect $t (type $r) (local.get $i)))
  (func (export "is_null") (param $i i32) (result i32)
    (ref.is_null (table.get $t (local.get $i)))))

(assert_return (invoke "call" (i32.const 0)) (i32.const 11))
(assert_return (invoke "is_null" (i32.const 1)) (i32.const 1))
(assert_return (invoke "call" (i32.const 2)) (i32.const 22))

;; ── An externref table filled from a global ─────────────────────────
(module
  (global $g (ref i31) (ref.i31 (i32.const 55)))
  (table $t 2 i31ref)
  (elem (table $t) (i32.const 0) i31ref (item (global.get $g)))
  (func (export "get") (result i32) (i31.get_u (table.get $t (i32.const 0)))))

(assert_return (invoke "get") (i32.const 55))

;; ── PASSIVE segments hold element expressions too ───────────────────
;; ⛔ EVERY CASE ABOVE IS AN *ACTIVE* SEGMENT, and that is why this file went
;; green while `gc/i31` stayed red. An active segment lowers to `table.set`
;; statements and evaluates its expressions like any other code; a passive one
;; is a segment the VM materializes at instantiation, and it could hold nothing
;; but function references. The two paths share a parser and no lowering, so
;; covering one says nothing about the other.
(module
  (table $t 6 i31ref)
  (elem $e i31ref
    (item (ref.i31 (i32.const 123)))
    (item (ref.i31 (i32.const 456)))
    (item (ref.null i31))
    (item (ref.i31 (i32.const 789))))

  (func (export "init") (param $d i32) (param $s i32) (param $n i32)
    (table.init $t $e (local.get $d) (local.get $s) (local.get $n)))
  (func (export "get") (param $i i32) (result i32)
    (i31.get_u (table.get $t (local.get $i))))
  (func (export "is_null") (param $i i32) (result i32)
    (ref.is_null (table.get $t (local.get $i))))
  (func (export "drop") (elem.drop $e)))

;; Nothing is in the table until the segment is copied in.
(assert_return (invoke "is_null" (i32.const 0)) (i32.const 1))
(invoke "init" (i32.const 0) (i32.const 0) (i32.const 4))
(assert_return (invoke "get" (i32.const 0)) (i32.const 123))
(assert_return (invoke "get" (i32.const 1)) (i32.const 456))
;; The `ref.null` item OCCUPIES its slot: dropping it would shift 789 down.
(assert_return (invoke "is_null" (i32.const 2)) (i32.const 1))
(assert_return (invoke "get" (i32.const 3)) (i32.const 789))

;; Copying from a non-zero source offset — the check that a dropped slot
;; renumbers the segment rather than merely blanking one entry.
(invoke "init" (i32.const 4) (i32.const 3) (i32.const 1))
(assert_return (invoke "get" (i32.const 4)) (i32.const 789))

;; A dropped segment is an EMPTY one, not an error.
(invoke "drop")
(assert_return (invoke "init" (i32.const 0) (i32.const 0) (i32.const 0)))
(assert_trap (invoke "init" (i32.const 0) (i32.const 0) (i32.const 1))
             "out of bounds table access")

;; ── An element expression is evaluated ONCE, at instantiation ───────
;; `(item (array.new_default …))` ALLOCATES. The spec evaluates element
;; expressions during module instantiation, so every read of that slot yields
;; the SAME array — two `array.new_elem`s off one segment must return
;; references that compare equal. Re-evaluating per use would allocate twice
;; and answer 0, which is the shape the spec's own `array_new_elem.wast`
;; checks. It is also why the segment cannot be materialized lazily.
(module
  (type $inner (array (mut i32)))
  (type $arr (array (mut arrayref)))
  (elem $elem arrayref (item (array.new_default $inner (i32.const 3))))
  (func (export "same") (result i32)
    (local $a (ref null $arr))
    (local $b (ref null $arr))
    (local.set $a (array.new_elem $arr $elem (i32.const 0) (i32.const 1)))
    (local.set $b (array.new_elem $arr $elem (i32.const 0) (i32.const 1)))
    (ref.eq (array.get $arr (local.get $a) (i32.const 0))
            (array.get $arr (local.get $b) (i32.const 0))))
  ;; …and the allocated array is a real one of the declared length.
  (func (export "len") (result i32)
    (array.len (array.get $arr
      (array.new_elem $arr $elem (i32.const 0) (i32.const 1))
      (i32.const 0)))))

(assert_return (invoke "same") (i32.const 1))
(assert_return (invoke "len") (i32.const 3))
