;; vybe-test: wast/wast_script_assert_invalid/assert_invalid_constant_expression
;; origin: languages/wast/tests/wast/test_wast_script_assert_invalid.rs
;; vybe-test-mode: compile

(assert_invalid (module (global i32 (i32.add (i32.const 1) (i32.const 2)))) "constant expression required")
