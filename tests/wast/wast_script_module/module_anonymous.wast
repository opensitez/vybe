;; vybe-test: wast/wast_script_module/module_anonymous
;; origin: languages/wast/tests/wast/test_wast_script_module.rs
;; vybe-test-mode: run

(module (func (export "f") (result i32) i32.const 42))
(assert_return (invoke "f") (i32.const 42))
