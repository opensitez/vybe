;; vybe-test: wast/wast_script_assert_return/assert_return_ref_func
;; origin: languages/wast/tests/wast/test_wast_script_assert_return.rs
;; vybe-test-mode: compile

(module 
  (func $dummy)
  (func (export "f") (result funcref) ref.func $dummy)
)
(assert_return (invoke "f") (ref.func))
