;; vybe-test: wast/wast_assert_return_values/assert_i64_arithmetic_exact
;; origin: languages/wast/tests/wast/test_wast_assert_return_values.rs
;; vybe-test-mode: compile

(module (func (export "m") (param i64 i64) (result i64)
            local.get 0 local.get 1 i64.mul))
          (assert_return (invoke "m" (i64.const 1000000000) (i64.const 1000000000))
            (i64.const 1000000000000000000))
