;; vybe-test: wast/wast_script/register_anonymous_module
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "f") (result i32) i32.const 1))
(register "mymod")
