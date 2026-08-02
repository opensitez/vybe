;; vybe-test: wast/wat_module/func_with_locals
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (func (param i32) (result i32) (local i32) local.get 0 local.set 1 local.get 1))
