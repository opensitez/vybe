;; vybe-test: wast/wat_custom_descriptors/test_ref_get_desc_null_reference_traps
;; origin: proposals/custom-descriptors/test/core/custom-descriptors/ref_get_desc.wast
;; vybe-test-mode: run-fail

;; `ref.get_desc x : (ref null (exact_1 x)) -> (ref (exact_1 y))` — the RESULT
;; is non-nullable, so a null input cannot be passed through. `ref_get_desc.wast`
;; asserts "null reference" six ways; this pins the instruction itself.

(module
  (rec
    (type $a (descriptor $a.desc) (struct))
    (type $a.desc (describes $a) (struct))
  )
  (func (export "_start")
    ref.null none
    ref.get_desc $a
    drop
  )
)
