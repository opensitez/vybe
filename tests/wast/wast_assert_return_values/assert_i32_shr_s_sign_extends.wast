;; vybe-test: wast/wast_assert_return_values/assert_i32_shr_s_sign_extends
;; origin: languages/wast/tests/wast/test_wast_assert_return_values.rs
;; vybe-test-mode: compile

(module (func (export "s") (param i32 i32) (result i32)
            local.get 0 local.get 1 i32.shr_s))
          (assert_return (invoke "s" (i32.const -8) (i32.const 1)) (i32.const -4))
