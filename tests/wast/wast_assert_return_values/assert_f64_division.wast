;; vybe-test: wast/wast_assert_return_values/assert_f64_division
;; origin: languages/wast/tests/wast/test_wast_assert_return_values.rs
;; vybe-test-mode: compile

(module (func (export "d") (param f64 f64) (result f64)
            local.get 0 local.get 1 f64.div))
          (assert_return (invoke "d" (f64.const 9.0) (f64.const 2.0)) (f64.const 4.5))
