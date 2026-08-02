;; vybe-test: wast/wat_control_flow/br_if_conditional
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module (func (param i32) (block $b local.get 0 br_if $b)))
