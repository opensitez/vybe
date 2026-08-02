;; vybe-test: wast/wast_script_invoke/invoke_on_module_alias
;; origin: languages/wast/tests/wast/test_wast_script_invoke.rs
;; vybe-test-mode: compile

(module $m (func (export "f") nop))
(invoke $m "f")
