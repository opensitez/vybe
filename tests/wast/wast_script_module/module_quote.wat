;; vybe-test: wast/wast_script_module/module_quote
;; origin: languages/wast/tests/wast/test_wast_script_module.rs
;; vybe-test-mode: compile

(module quote "(module (func (export \"f\") (result i32) i32.const 42))")
