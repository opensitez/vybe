;; vybe-test: wast/wast_script_module/module_multiple_named
;; origin: languages/wast/tests/wast/test_wast_script_module.rs
;; vybe-test-mode: compile

(module $m1 (func (export "f") (result i32) i32.const 42))
(module $m2 (func (export "f") (result i32) i32.const 99))
(assert_return (invoke $m1 "f") (i32.const 42))
(assert_return (invoke $m2 "f") (i32.const 99))
