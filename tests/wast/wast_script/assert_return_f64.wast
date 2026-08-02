;; vybe-test: wast/wast_script/assert_return_f64
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "f") (result f64) f64.const 3.14))
(assert_return (invoke "f") (f64.const 3.14))
