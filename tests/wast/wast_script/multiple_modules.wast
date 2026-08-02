;; vybe-test: wast/wast_script/multiple_modules
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module $a (func (export "f") (result i32) i32.const 1))
(module $b (func (export "g") (result i32) i32.const 2))
(assert_return (invoke $a "f") (i32.const 1))
(assert_return (invoke $b "g") (i32.const 2))
