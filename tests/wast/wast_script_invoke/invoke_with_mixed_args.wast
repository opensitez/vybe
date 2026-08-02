;; vybe-test: wast/wast_script_invoke/invoke_with_mixed_args
;; origin: languages/wast/tests/wast/test_wast_script_invoke.rs
;; vybe-test-mode: compile

(module (func (export "f") (param i32 i64 f32 f64) nop))
(invoke "f" (i32.const 1) (i64.const 2) (f32.const 3.0) (f64.const 4.0))
