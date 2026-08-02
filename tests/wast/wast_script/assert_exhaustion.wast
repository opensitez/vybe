;; vybe-test: wast/wast_script/assert_exhaustion
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func $inf (export "inf") call $inf))
(assert_exhaustion (invoke "inf") "call stack exhausted")
