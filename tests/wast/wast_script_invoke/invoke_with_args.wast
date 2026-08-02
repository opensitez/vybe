;; vybe-test: wast/wast_script_invoke/invoke_with_args
;; origin: languages/wast/tests/wast/test_wast_script_invoke.rs
;; vybe-test-mode: compile

(module (func (export "f") (param i32 i32) nop))
(invoke "f" (i32.const 10) (i32.const 20))
