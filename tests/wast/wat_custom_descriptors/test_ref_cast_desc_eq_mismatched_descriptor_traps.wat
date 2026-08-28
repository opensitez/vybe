;; vybe-test: wast/wat_custom_descriptors/test_ref_cast_desc_eq_mismatched_descriptor_traps
;; origin: proposals/custom-descriptors/test/core/custom-descriptors/ref_cast_desc_eq.wast
;; vybe-test-mode: run-fail

;; Two SEPARATELY allocated descriptors of the same type are different
;; references, so the comparison is identity, not structural equality — even
;; though both descriptors here hold the same field value.

(module
  (rec
    (type $a (descriptor $a.desc) (struct (field i32)))
    (type $a.desc (describes $a) (struct (field i32)))
  )

  (func (export "_start") (local $d1 (ref null $a.desc)) (local $d2 (ref null $a.desc))
    i32.const 5
    struct.new $a.desc
    local.set $d1
    i32.const 5
    struct.new $a.desc
    local.set $d2

    i32.const 1
    local.get $d1
    struct.new_desc $a
    local.get $d2
    ref.cast_desc_eq (ref $a)
    drop
  )
)
