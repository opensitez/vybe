;; vybe-test: wast/wast_script_invoke/invoke_with_nan_args
;; origin: languages/wast/tests/wast/test_wast_script_invoke.rs
;; vybe-test-mode: compile

(module (func (export "f") (param f32) nop))
(invoke "f" (f32.const nan))
