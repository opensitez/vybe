;; vybe-test: wast/wast_assert_return_values/assert_multiple_invocations_same_module
;; origin: languages/wast/tests/wast/test_wast_assert_return_values.rs
;; vybe-test-mode: compile

(module (func (export "sq") (param i32) (result i32) local.get 0 local.get 0 i32.mul))
          (assert_return (invoke "sq" (i32.const 3)) (i32.const 9))
          (assert_return (invoke "sq" (i32.const 5)) (i32.const 25))
          (assert_return (invoke "sq" (i32.const 12)) (i32.const 144))
