;; vybe-test: wast/wast_script/assert_unlinkable_missing_import
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(assert_unlinkable (module (import "env" "missing" (func))) "unknown import")
