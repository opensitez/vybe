;; vybe-test: wast/wast_script_assert_malformed/assert_malformed_invalid_integer
;; origin: languages/wast/tests/wast/test_wast_script_assert_malformed.rs
;; vybe-test-mode: compile

(assert_malformed (module quote "(module (func i32.const 9999999999999999999999999))") "constant out of range")
