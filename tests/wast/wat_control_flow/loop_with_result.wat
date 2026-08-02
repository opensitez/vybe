;; vybe-test: wast/wat_control_flow/loop_with_result
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module (func (result i32) (loop (result i32) i32.const 42)))
