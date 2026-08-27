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
