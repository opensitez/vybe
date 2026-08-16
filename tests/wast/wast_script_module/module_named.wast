;; vybe-test: wast/wast_script_module/module_named
;; origin: languages/wast/tests/wast/test_wast_script_module.rs
;; vybe-test-mode: run

(module $m (func (export "f") (result i32) i32.const 42))
(assert_return (invoke $m "f") (i32.const 42))
