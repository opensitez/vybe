;; vybe-test: wast/wast_assert_return_values/assert_i32_clz
;; origin: languages/wast/tests/wast/test_wast_assert_return_values.rs
;; vybe-test-mode: compile

(module (func (export "c") (param i32) (result i32) local.get 0 i32.clz))
          (assert_return (invoke "c" (i32.const 1)) (i32.const 31))
