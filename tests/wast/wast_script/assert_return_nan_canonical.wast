;; vybe-test: wast/wast_script/assert_return_nan_canonical
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "f") (result f32) f32.const nan))
(assert_return (invoke "f") (f32.const nan:canonical))
