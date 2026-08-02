;; vybe-test: wast/wast_script/invoke_on_named_module
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module $m (func (export "f")))
(invoke $m "f")
