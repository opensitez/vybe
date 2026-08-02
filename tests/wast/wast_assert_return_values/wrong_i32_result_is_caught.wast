;; vybe-test: wast/wast_assert_return_values/wrong_i32_result_is_caught
;; origin: languages/wast/tests/wast/test_wast_assert_return_values.rs
;; vybe-test-mode: compile

(module (func (export "f") (result i32) i32.const 41))
           (assert_return (invoke "f") (i32.const 42))
