;; vybe-test: wast/wast_script_assert_invalid/assert_invalid_type_mismatch
;; origin: languages/wast/tests/wast/test_wast_script_assert_invalid.rs
;; vybe-test-mode: compile

(assert_invalid (module (func (result i32) f32.const 1.0)) "type mismatch")
