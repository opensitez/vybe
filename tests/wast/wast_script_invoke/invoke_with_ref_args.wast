;; vybe-test: wast/wast_script_invoke/invoke_with_ref_args
;; origin: languages/wast/tests/wast/test_wast_script_invoke.rs
;; vybe-test-mode: compile

(module (func (export "f") (param funcref externref) nop))
(invoke "f" (ref.null func) (ref.null extern))
