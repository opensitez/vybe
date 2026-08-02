;; vybe-test: wast/wast_script_assert_return/assert_return_multiple_invokes
;; origin: languages/wast/tests/wast/test_wast_script_assert_return.rs
;; vybe-test-mode: compile

(module 
  (func (export "f") (result i32) i32.const 1)
  (func (export "g") (result i32) i32.const 2)
)
(assert_return (invoke "f") (i32.const 1))
(assert_return (invoke "g") (i32.const 2))
