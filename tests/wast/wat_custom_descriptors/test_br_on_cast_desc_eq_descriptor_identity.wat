;; vybe-test: wast/wat_custom_descriptors/test_br_on_cast_desc_eq_descriptor_identity
;; origin: proposals/custom-descriptors/test/core/custom-descriptors/br_on_cast_desc_eq.wast

;; `(assert_return (invoke "self-nullable-val-desc") (i32.const 1))` and
;; `(assert_return (invoke "self-nullable-val-other") (i32.const 0))`.
;;
;; The branch is taken on descriptor IDENTITY, not on structural equality:
;; `$d2` below is a second descriptor of the same type holding the SAME field
;; value as `$d1`, and it must NOT match. A value compare would answer 1 here.

(module
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)

  (rec
    (type $a (descriptor $a.desc) (struct (field i32)))
    (type $a.desc (describes $a) (struct (field i32)))
  )

  ;; 1 when the branch is taken, 0 on the fall-through.
  (func $same (result i32) (local $d (ref null $a.desc)) (local $o (ref null $a))
    i32.const 7
    struct.new $a.desc
    local.set $d

    i32.const 1
    local.get $d
    struct.new_desc $a
    local.set $o

    (block (result anyref)
      (br_on_cast_desc_eq 0 anyref (ref null $a) (local.get $o) (local.get $d))
      (return (i32.const 0))
    )
    drop
    (return (i32.const 1))
  )

  ;; Same shape, but the descriptor compared against is a DIFFERENT allocation
  ;; carrying an equal field value.
  (func $other (result i32)
    (local $d1 (ref null $a.desc)) (local $d2 (ref null $a.desc)) (local $o (ref null $a))
    i32.const 7
    struct.new $a.desc
    local.set $d1
    i32.const 7
    struct.new $a.desc
    local.set $d2

    i32.const 1
    local.get $d1
    struct.new_desc $a
    local.set $o

    (block (result anyref)
      (br_on_cast_desc_eq 0 anyref (ref null $a) (local.get $o) (local.get $d2))
      (return (i32.const 0))
    )
    drop
    (return (i32.const 1))
  )

  (func (export "_start")
    call $same
    i32.const 1
    call $vybe_check_i32

    call $other
    i32.const 0
    call $vybe_check_i32
  )
)
