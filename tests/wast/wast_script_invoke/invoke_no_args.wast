;; vybe-test: wast/wast_script_invoke/invoke_no_args
;; origin: languages/wast/tests/wast/test_wast_script_invoke.rs
;; vybe-test-mode: compile

(module (func (export "f") nop))
(invoke "f")
