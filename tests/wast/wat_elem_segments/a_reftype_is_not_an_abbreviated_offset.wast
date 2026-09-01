;; vybe-test: wast/wat_elem_segments/a_reftype_is_not_an_abbreviated_offset
;; vybe-test-mode: compile
;;
;; ⛔ AN ELEMENT SEGMENT'S REFTYPE IS NOT ITS MODE.
;;
;; An ACTIVE segment may abbreviate its offset to a bare folded instruction
;; (`(elem (i32.const 0) …)`), and a PASSIVE one may open with the element
;; list's reftype (`(elem $e (ref $t) …)`). Both are a parenthesised form in the
;; same position, so the offset shorthand swallowed the reftype and every such
;; segment was read as ACTIVE — populating table 0, which these modules do not
;; declare, and trapping "out of bounds table access" before any code ran.
;;
;; `(ref …)` and `(ref null …)` are never offsets. `(ref.func $f)` and
;; `(ref.null func)` are not reftypes, so element ITEMS still parse as items.

;; Passive, reftype-led, no table in the module at all.
(module
  (type $bvec (array i8))
  (type $vec (array (ref $bvec)))
  (elem $e (ref $bvec)
    (array.new $bvec (i32.const 7) (i32.const 3))
    (array.new_fixed $bvec 2 (i32.const 1) (i32.const 2))
  )
  (func (export "len") (result i32)
    (array.len (array.new_elem $vec $e (i32.const 0) (i32.const 2)))
  )
)
(assert_return (invoke "len") (i32.const 2))

;; Nullable reftype spelling, same shape.
(module
  (type $bvec (array i8))
  (type $vec (array (ref null $bvec)))
  (elem $e (ref null $bvec) (array.new $bvec (i32.const 1) (i32.const 1)))
  (func (export "len") (result i32)
    (array.len (array.new_elem $vec $e (i32.const 0) (i32.const 1)))
  )
)
(assert_return (invoke "len") (i32.const 1))

;; ⛔ AND THE ABBREVIATION MUST STILL WORK. An active segment led by a bare
;; folded offset still populates its table.
(module
  (table 4 funcref)
  (func $a (result i32) (i32.const 11))
  (func $b (result i32) (i32.const 22))
  (elem (i32.const 1) $a $b)
  (type $r (func (result i32)))
  (func (export "call") (param i32) (result i32)
    (call_indirect (type $r) (local.get 0))
  )
)
(assert_return (invoke "call" (i32.const 1)) (i32.const 11))
(assert_return (invoke "call" (i32.const 2)) (i32.const 22))

;; An active segment whose element list ALSO names a reftype: the mode is read
;; first, the reftype after it.
(module
  (table 2 funcref)
  (func $c (result i32) (i32.const 33))
  (elem (i32.const 0) funcref (ref.func $c) (ref.null func))
  (type $r (func (result i32)))
  (func (export "call") (result i32)
    (call_indirect (type $r) (i32.const 0))
  )
)
(assert_return (invoke "call") (i32.const 33))
