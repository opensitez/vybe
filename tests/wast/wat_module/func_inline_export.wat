;; vybe-test: wast/wat_module/func_inline_export
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (func (export "add") (param i32 i32) (result i32) local.get 0 local.get 1 i32.add))
