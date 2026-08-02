;; vybe-test: wast/wast_script_assert_invalid/assert_invalid_unknown_global
;; origin: languages/wast/tests/wast/test_wast_script_assert_invalid.rs
;; vybe-test-mode: compile

(assert_invalid (module (func (result i32) global.get 0)) "unknown global")
