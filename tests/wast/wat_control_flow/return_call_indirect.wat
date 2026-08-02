;; vybe-test: wast/wat_control_flow/return_call_indirect
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module (type $t (func (param i32) (result i32))) (table 1 funcref) (func (param i32 i32) (result i32) local.get 0 local.get 1 return_call_indirect (type $t)))
