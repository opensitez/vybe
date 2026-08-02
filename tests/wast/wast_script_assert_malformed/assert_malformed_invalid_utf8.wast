;; vybe-test: wast/wast_script_assert_malformed/assert_malformed_invalid_utf8
;; origin: languages/wast/tests/wast/test_wast_script_assert_malformed.rs
;; vybe-test-mode: compile

(assert_malformed (module quote "(module (data \"\\ff\"))") "invalid utf-8 encoding")
