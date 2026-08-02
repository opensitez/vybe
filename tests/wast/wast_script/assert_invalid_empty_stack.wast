;; vybe-test: wast/wast_script/assert_invalid_empty_stack
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(assert_invalid (module (func (result i32) nop)) "type mismatch")
