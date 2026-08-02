;; vybe-test: wast/wast_script/assert_malformed_quote
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(assert_malformed (module quote "(module (func (result i32)))") "unexpected token")
