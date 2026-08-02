;; vybe-test: wast/wast_assert_return_values/assert_i32_lt_u_unsigned_compare
;; origin: languages/wast/tests/wast/test_wast_assert_return_values.rs
;; vybe-test-mode: compile

(module (func (export "lt") (param i32 i32) (result i32)
            local.get 0 local.get 1 i32.lt_u))
          (assert_return (invoke "lt" (i32.const -1) (i32.const 1)) (i32.const 0))
