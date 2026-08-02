;; vybe-test: wast/wast_script_assert_return/assert_return_nan_arithmetic_f64
;; origin: languages/wast/tests/wast/test_wast_script_assert_return.rs
;; vybe-test-mode: compile

(module (func (export "f") (result f64) f64.const nan:arithmetic))
(assert_return (invoke "f") (f64.const nan:arithmetic))
