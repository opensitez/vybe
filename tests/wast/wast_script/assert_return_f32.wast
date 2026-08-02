;; vybe-test: wast/wast_script/assert_return_f32
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "f") (result f32) f32.const 1.0))
(assert_return (invoke "f") (f32.const 1.0))
