;; vybe-test: wast/wast_script_assert_invalid/assert_invalid_duplicate_export
;; origin: languages/wast/tests/wast/test_wast_script_assert_invalid.rs
;; vybe-test-mode: compile

(assert_invalid (module (func (export "a")) (func (export "a"))) "duplicate export name")
