;; vybe-test: wast/wast_script_assert_return/assert_return_basic_i64
;; origin: languages/wast/tests/wast/test_wast_script_assert_return.rs
;; vybe-test-mode: compile

(module (func (export "f") (result i64) i64.const 9999999999))
(assert_return (invoke "f") (i64.const 9999999999))
