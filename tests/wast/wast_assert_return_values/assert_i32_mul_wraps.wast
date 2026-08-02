;; vybe-test: wast/wast_assert_return_values/assert_i32_mul_wraps
;; origin: languages/wast/tests/wast/test_wast_assert_return_values.rs
;; vybe-test-mode: compile

(module (func (export "mul") (param i32 i32) (result i32)
            local.get 0 local.get 1 i32.mul))
          (assert_return (invoke "mul" (i32.const 65536) (i32.const 65536)) (i32.const 0))
