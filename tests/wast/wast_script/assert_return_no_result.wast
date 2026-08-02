;; vybe-test: wast/wast_script/assert_return_no_result
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "noop")))
(assert_return (invoke "noop"))
