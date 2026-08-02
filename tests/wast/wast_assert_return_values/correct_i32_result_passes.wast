;; vybe-test: wast/wast_assert_return_values/correct_i32_result_passes
;; origin: languages/wast/tests/wast/test_wast_assert_return_values.rs
;; vybe-test-mode: compile

(module (func (export "f") (result i32) i32.const 42))
          (assert_return (invoke "f") (i32.const 42))
