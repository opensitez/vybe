;; vybe-test: wast/wast_script/assert_return_i64
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "f") (result i64) i64.const 100))
(assert_return (invoke "f") (i64.const 100))
