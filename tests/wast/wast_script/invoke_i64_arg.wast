;; vybe-test: wast/wast_script/invoke_i64_arg
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "f") (param i64)))
(invoke "f" (i64.const 9999999999))
