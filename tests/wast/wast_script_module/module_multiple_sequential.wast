;; vybe-test: wast/wast_script_module/module_multiple_sequential
;; origin: languages/wast/tests/wast/test_wast_script_module.rs
;; vybe-test-mode: compile

(module (func (export "f") (result i32) i32.const 42))
(assert_return (invoke "f") (i32.const 42))
(module (func (export "g") (result i32) i32.const 99))
(assert_return (invoke "g") (i32.const 99))
