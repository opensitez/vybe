;; vybe-test: wast/wast_script/assert_return_multiple_results
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "f") (result i32 i32) i32.const 1 i32.const 2))
(assert_return (invoke "f") (i32.const 1) (i32.const 2))
