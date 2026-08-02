;; vybe-test: wast/wast_script/invoke_no_args
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "noop")))
(invoke "noop")
