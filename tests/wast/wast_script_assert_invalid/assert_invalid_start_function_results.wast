;; vybe-test: wast/wast_script_assert_invalid/assert_invalid_start_function_results
;; origin: languages/wast/tests/wast/test_wast_script_assert_invalid.rs
;; vybe-test-mode: compile

(assert_invalid (module (func $start (result i32) i32.const 0) (start $start)) "start function")
