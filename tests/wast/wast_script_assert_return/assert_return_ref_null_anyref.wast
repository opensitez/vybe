;; vybe-test: wast/wast_script_assert_return/assert_return_ref_null_anyref
;; origin: languages/wast/tests/wast/test_wast_script_assert_return.rs
;; vybe-test-mode: compile

(module (func (export "f") (result anyref) ref.null any))
(assert_return (invoke "f") (ref.null any))
