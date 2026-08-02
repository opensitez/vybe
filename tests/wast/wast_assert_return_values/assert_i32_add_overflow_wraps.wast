;; vybe-test: wast/wast_assert_return_values/assert_i32_add_overflow_wraps
;; origin: languages/wast/tests/wast/test_wast_assert_return_values.rs
;; vybe-test-mode: compile

(module (func (export "add") (param i32 i32) (result i32)
            local.get 0 local.get 1 i32.add))
          (assert_return (invoke "add" (i32.const 2147483647) (i32.const 1))
            (i32.const -2147483648))
