;; vybe-test: wast/wast_script_assert_invalid/assert_invalid_start_function_params
;; origin: languages/wast/tests/wast/test_wast_script_assert_invalid.rs
;; vybe-test-mode: compile

(assert_invalid (module (func $start (param i32)) (start $start)) "start function")
