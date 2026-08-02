;; vybe-test: wast/wast_script_assert_return/assert_return_nan_arithmetic_f32
;; origin: languages/wast/tests/wast/test_wast_script_assert_return.rs
;; vybe-test-mode: compile

(module (func (export "f") (result f32) f32.const nan:arithmetic))
(assert_return (invoke "f") (f32.const nan:arithmetic))
