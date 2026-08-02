;; vybe-test: wast/wast_script_invoke/invoke_with_f64_args
;; origin: languages/wast/tests/wast/test_wast_script_invoke.rs
;; vybe-test-mode: compile

(module (func (export "f") (param f64) nop))
(invoke "f" (f64.const 1.0))
