;; vybe-test: wast/wast_script_invoke/invoke_with_i64_args
;; origin: languages/wast/tests/wast/test_wast_script_invoke.rs
;; vybe-test-mode: compile

(module (func (export "f") (param i64) nop))
(invoke "f" (i64.const 9999999999))
