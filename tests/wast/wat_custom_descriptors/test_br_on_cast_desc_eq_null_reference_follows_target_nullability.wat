;; vybe-test: wast/wat_custom_descriptors/test_br_on_cast_desc_eq_null_reference_follows_target_nullability
;; origin: proposals/custom-descriptors/test/core/custom-descriptors/br_on_cast_desc_eq.wast

;; `(assert_return (invoke "self-nullable-null-desc") (i32.const 1))` against
;; `(assert_return (invoke "self-nonnullable-null-desc") (i32.const 0))`.
;;
;; A NULL reference with a valid descriptor branches iff the TARGET reftype is
;; nullable — the only difference between the two functions below is
;; `(ref null $a)` vs `(ref $a)`. This is why the nullability of the immediate
;; has to survive lowering instead of being folded down to a heap type.
;;
;; It is also why the null-reference case is answered from the immediate and
;; never reaches `ref.get_desc`, whose result type is non-nullable and which
;; therefore traps on a null input.

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

  (func $nullable_target (result i32) (local $d (ref null $a.desc))
    i32.const 7
    struct.new $a.desc
    local.set $d
    (block (result anyref)
      (br_on_cast_desc_eq 0 anyref (ref null $a) (ref.null none) (local.get $d))
      (return (i32.const 0))
    )
    drop
    (return (i32.const 1))
  )

  (func $nonnullable_target (result i32) (local $d (ref null $a.desc))
    i32.const 7
    struct.new $a.desc
    local.set $d
    (block (result anyref)
      (br_on_cast_desc_eq 0 anyref (ref $a) (ref.null none) (local.get $d))
      (return (i32.const 0))
    )
    drop
    (return (i32.const 1))
  )

  (func (export "_start")
    call $nullable_target
    i32.const 1
    call $vybe_check_i32

    call $nonnullable_target
    i32.const 0
    call $vybe_check_i32
  )
)
