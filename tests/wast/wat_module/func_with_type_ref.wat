;; vybe-test: wast/wat_module/func_with_type_ref
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (type $t (func (param i32) (result i32))) (func (type $t) local.get 0))
