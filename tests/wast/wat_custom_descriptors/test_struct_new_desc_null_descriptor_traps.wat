;; vybe-test: wast/wat_custom_descriptors/test_struct_new_desc_null_descriptor_traps
;; origin: proposals/custom-descriptors/test/core/custom-descriptors/struct_new_desc.wast
;; vybe-test-mode: run-fail

;; `(assert_trap (invoke "new-null") "null descriptor reference")` — the
;; descriptor operand is typed `(ref null (exact y))`, so a null reaches the
;; instruction and traps there rather than being rejected at validation.

(module
  (rec
    (type $a (descriptor $a.desc) (struct (field i32)))
    (type $a.desc (describes $a) (struct))
  )
  (func (export "_start")
    i32.const 1
    ref.null none
    struct.new_desc $a
    drop
  )
)
