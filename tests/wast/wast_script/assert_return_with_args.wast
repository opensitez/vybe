;; vybe-test: wast/wast_script/assert_return_with_args
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "add") (param i32 i32) (result i32) local.get 0 local.get 1 i32.add))
(assert_return (invoke "add" (i32.const 3) (i32.const 4)) (i32.const 7))
