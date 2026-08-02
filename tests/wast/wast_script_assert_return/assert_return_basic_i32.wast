;; vybe-test: wast/wast_script_assert_return/assert_return_basic_i32
;; origin: languages/wast/tests/wast/test_wast_script_assert_return.rs
;; vybe-test-mode: compile

(module (func (export "f") (result i32) i32.const 42))
(assert_return (invoke "f") (i32.const 42))
