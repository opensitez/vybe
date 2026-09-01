;; vybe-test: wast/wat_array_new/a_segment_operand_is_unsigned
;; vybe-test-mode: compile
;;
;; ⛔ EVERY OFFSET, INDEX AND COUNT IN THESE OPS IS AN UNSIGNED i32.
;;
;; `0x8000_0000` is negative as an i32 and 2147483648 as the unsigned value the
;; spec reads. Clamping the negative to zero turns the most out-of-bounds
;; request there is into the most in-bounds one, so the trap never fires and the
;; op quietly succeeds on an empty range.
;;
;; ⛔ AND THE COUNT IS IN ELEMENTS WHILE THE SEGMENT IS IN BYTES. `(array i32)`
;; needs four bytes per element, so one element off a one-byte segment traps.

(module
  (type $a8  (array (mut i8)))
  (type $a32 (array (mut i32)))
  (data $d "abcd")
  (func $f)
  (elem $e func $f)

  ;; Unsigned operands on the allocating form.
  (func (export "new_data-huge-size") (result i32)
    (array.len (array.new_data $a8 $d (i32.const 0) (i32.const 0x8000_0000))))
  (func (export "new_data-huge-offset") (result i32)
    (array.len (array.new_data $a8 $d (i32.const 0x8000_0000) (i32.const 1))))
  (func (export "new_elem-huge-size") (result i32)
    (array.len (array.new_elem $a8 $e (i32.const 0) (i32.const 0x8000_0000))))

  ;; Unsigned operands on the filling form — the sibling that shares the
  ;; helper and that no passing fixture had ever fed a negative.
  (func (export "init_data-huge-size")
    (array.init_data $a8 $d
      (array.new_default $a8 (i32.const 4))
      (i32.const 0) (i32.const 0) (i32.const 0x8000_0000)))
  (func (export "init_data-huge-src")
    (array.init_data $a8 $d
      (array.new_default $a8 (i32.const 4))
      (i32.const 0) (i32.const 0x8000_0000) (i32.const 1)))
  (func (export "init_elem-huge-size")
    (array.init_elem $a8 $e
      (array.new_default $a8 (i32.const 4))
      (i32.const 0) (i32.const 0) (i32.const 0x8000_0000)))

  ;; Element width: four bytes per i32, so one element needs all four.
  (func (export "i32-one-elem") (result i32)
    (array.len (array.new_data $a32 $d (i32.const 0) (i32.const 1))))
  (func (export "i32-two-elems") (result i32)
    (array.len (array.new_data $a32 $d (i32.const 0) (i32.const 2))))
  (func (export "i32-offset-past") (result i32)
    (array.len (array.new_data $a32 $d (i32.const 1) (i32.const 1))))
)

(assert_trap (invoke "new_data-huge-size") "out of bounds memory access")
(assert_trap (invoke "new_data-huge-offset") "out of bounds memory access")
(assert_trap (invoke "new_elem-huge-size") "out of bounds table access")
(assert_trap (invoke "init_data-huge-size") "out of bounds")
(assert_trap (invoke "init_data-huge-src") "out of bounds")
(assert_trap (invoke "init_elem-huge-size") "out of bounds")

;; Four bytes is exactly one i32 element; two need eight and there are four.
(assert_return (invoke "i32-one-elem") (i32.const 1))
(assert_trap (invoke "i32-two-elems") "out of bounds memory access")
(assert_trap (invoke "i32-offset-past") "out of bounds memory access")
