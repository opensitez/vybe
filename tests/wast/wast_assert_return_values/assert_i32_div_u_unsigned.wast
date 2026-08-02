;; vybe-test: wast/wast_assert_return_values/assert_i32_div_u_unsigned
;; origin: languages/wast/tests/wast/test_wast_assert_return_values.rs
;; vybe-test-mode: compile

(module (func (export "d") (param i32 i32) (result i32)
            local.get 0 local.get 1 i32.div_u))
          (assert_return (invoke "d" (i32.const -2) (i32.const 2)) (i32.const 2147483647))
