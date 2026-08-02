;; vybe-test: wast/wat_control_flow/if_else_plain
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module (func (param i32) (result i32) local.get 0 if (result i32) i32.const 1 else i32.const 0 end))
