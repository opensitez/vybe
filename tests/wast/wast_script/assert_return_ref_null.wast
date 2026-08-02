;; vybe-test: wast/wast_script/assert_return_ref_null
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "f") (result funcref) ref.null func))
(assert_return (invoke "f") (ref.null func))
