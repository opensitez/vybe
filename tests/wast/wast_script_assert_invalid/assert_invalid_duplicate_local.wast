;; vybe-test: wast/wast_script_assert_invalid/assert_invalid_duplicate_local
;; origin: languages/wast/tests/wast/test_wast_script_assert_invalid.rs
;; vybe-test-mode: compile

(assert_invalid (module (func (local $a i32) (local $a i32))) "duplicate local")
