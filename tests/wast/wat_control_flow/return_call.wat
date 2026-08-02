;; vybe-test: wast/wat_control_flow/return_call
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module (func $f (param i32) (result i32) local.get 0) (func (param i32) (result i32) local.get 0 return_call $f))
