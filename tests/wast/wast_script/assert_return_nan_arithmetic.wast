;; vybe-test: wast/wast_script/assert_return_nan_arithmetic
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "f") (result f64) f64.const nan))
(assert_return (invoke "f") (f64.const nan:arithmetic))
