;; vybe-test: wast/wat_custom_descriptors/test_struct_new_default_desc_null_descriptor_traps
;; origin: proposals/custom-descriptors/test/core/custom-descriptors/struct_new_desc.wast
;; vybe-test-mode: run-fail

;; The same trap on the default-initialising form.

(module
  (rec
    (type $a (descriptor $a.desc) (struct (field i32)))
    (type $a.desc (describes $a) (struct))
  )
  (func (export "_start")
    ref.null none
    struct.new_default_desc $a
    drop
  )
)
