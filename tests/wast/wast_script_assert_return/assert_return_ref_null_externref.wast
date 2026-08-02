;; vybe-test: wast/wast_script_assert_return/assert_return_ref_null_externref
;; origin: languages/wast/tests/wast/test_wast_script_assert_return.rs
;; vybe-test-mode: compile

(module (func (export "f") (result externref) ref.null extern))
(assert_return (invoke "f") (ref.null extern))
