;; vybe-test: wast/wast_script/invoke_f64_arg
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "f") (param f64)))
(invoke "f" (f64.const 2.718))
