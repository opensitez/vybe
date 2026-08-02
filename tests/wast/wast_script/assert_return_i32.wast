;; vybe-test: wast/wast_script/assert_return_i32
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "forty_two") (result i32) i32.const 42))
(assert_return (invoke "forty_two") (i32.const 42))
