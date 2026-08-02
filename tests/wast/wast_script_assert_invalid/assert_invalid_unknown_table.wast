;; vybe-test: wast/wast_script_assert_invalid/assert_invalid_unknown_table
;; origin: languages/wast/tests/wast/test_wast_script_assert_invalid.rs
;; vybe-test-mode: compile

(assert_invalid (module (func (result funcref) i32.const 0 table.get 0)) "unknown table")
