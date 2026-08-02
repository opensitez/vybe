;; vybe-test: wast/wast_script_assert_invalid/assert_invalid_import_after_func
;; origin: languages/wast/tests/wast/test_wast_script_assert_invalid.rs
;; vybe-test-mode: compile

(assert_invalid (module (func) (import "a" "b" (func))) "import after function")
