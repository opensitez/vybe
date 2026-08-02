;; vybe-test: wast/wast_script_invoke/invoke_with_f32_args
;; origin: languages/wast/tests/wast/test_wast_script_invoke.rs
;; vybe-test-mode: compile

(module (func (export "f") (param f32) nop))
(invoke "f" (f32.const 1.0))
