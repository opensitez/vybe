;; vybe-test: wast/wast_script/invoke_f32_arg
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "f") (param f32)))
(invoke "f" (f32.const 1.5))
