;; vybe-test: wast/wat_module/func_with_params_result
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.add))
