;; vybe-test: wast/wast_script_assert_malformed/assert_malformed_invalid_character
;; origin: languages/wast/tests/wast/test_wast_script_assert_malformed.rs
;; vybe-test-mode: run

(assert_malformed (module quote "(module (func @))") "unexpected token")
