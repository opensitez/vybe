;; vybe-test: wast/wat_module/func_multiple_inline_exports
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (func (export "add") (export "sum") (param i32 i32) (result i32) local.get 0 local.get 1 i32.add))
