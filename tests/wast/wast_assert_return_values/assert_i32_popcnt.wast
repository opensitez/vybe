;; vybe-test: wast/wast_assert_return_values/assert_i32_popcnt
;; origin: languages/wast/tests/wast/test_wast_assert_return_values.rs
;; vybe-test-mode: compile

(module (func (export "p") (param i32) (result i32) local.get 0 i32.popcnt))
          (assert_return (invoke "p" (i32.const 255)) (i32.const 8))
