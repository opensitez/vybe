;; vybe-test: wast/wat_custom_descriptors/test_br_on_cast_desc_eq_fail_inverts_the_branch
;; origin: proposals/custom-descriptors/test/core/custom-descriptors/br_on_cast_desc_eq_fail.wast

;; `br_on_cast_desc_eq_fail` branches on the COMPLEMENT: it takes the label
;; when the descriptors do not match and falls through when they do. Both forms
;; carry the reference to the target block either way — for `_fail` the
;; proposal types that as `rt_1 \ rt_2`, which is still the reference.
;;
;; The null-descriptor trap is NOT inverted: it applies to both forms.

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

  ;; Matching descriptor ⇒ `_fail` does NOT branch ⇒ 0.
  (func $matching (result i32) (local $d (ref null $a.desc)) (local $o (ref null $a))
    i32.const 7
    struct.new $a.desc
    local.set $d
    i32.const 1
    local.get $d
    struct.new_desc $a
    local.set $o
    (block (result anyref)
      (br_on_cast_desc_eq_fail 0 anyref (ref null $a) (local.get $o) (local.get $d))
      (return (i32.const 0))
    )
    drop
    (return (i32.const 1))
  )

  ;; Mismatched descriptor ⇒ `_fail` DOES branch ⇒ 1.
  (func $mismatched (result i32)
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
      (br_on_cast_desc_eq_fail 0 anyref (ref null $a) (local.get $o) (local.get $d2))
      (return (i32.const 0))
    )
    drop
    (return (i32.const 1))
  )

  (func (export "_start")
    call $matching
    i32.const 0
    call $vybe_check_i32

    call $mismatched
    i32.const 1
    call $vybe_check_i32
  )
)
