;; vybe-test: wast/wast_script_assert_return/assert_return_empty
;; origin: languages/wast/tests/wast/test_wast_script_assert_return.rs
;; vybe-test-mode: compile

(module (func (export "f")))
(assert_return (invoke "f"))
