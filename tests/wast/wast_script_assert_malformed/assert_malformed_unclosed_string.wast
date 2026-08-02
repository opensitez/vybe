;; vybe-test: wast/wast_script_assert_malformed/assert_malformed_unclosed_string
;; origin: languages/wast/tests/wast/test_wast_script_assert_malformed.rs
;; vybe-test-mode: compile

(assert_malformed (module quote "(module (func (export \"f\"))") "unexpected token")
