;; vybe-test: wast/wast_script_invoke/invoke_fails_on_missing_export
;; origin: languages/wast/tests/wast/test_wast_script_invoke.rs
;; vybe-test-mode: compile

(module)
(assert_trap (invoke "f") "unknown export")
