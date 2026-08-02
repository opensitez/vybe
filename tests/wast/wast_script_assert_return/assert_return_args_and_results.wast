;; vybe-test: wast/wast_script_assert_return/assert_return_args_and_results
;; origin: languages/wast/tests/wast/test_wast_script_assert_return.rs
;; vybe-test-mode: compile

(module 
  (func (export "f") (param i32 i32) (result i32 i32) 
    local.get 1 
    local.get 0)
)
(assert_return (invoke "f" (i32.const 10) (i32.const 20)) (i32.const 20) (i32.const 10))
