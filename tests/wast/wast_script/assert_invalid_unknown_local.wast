;; vybe-test: wast/wast_script/assert_invalid_unknown_local
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(assert_invalid (module (func (result i32) i32.const 1)) "unknown local")
