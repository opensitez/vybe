;; vybe-test: wast/wast_script/invoke_i32_arg
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "f") (param i32)))
(invoke "f" (i32.const 42))
