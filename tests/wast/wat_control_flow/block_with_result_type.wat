;; vybe-test: wast/wat_control_flow/block_with_result_type
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module (func (result i32) (block (result i32) i32.const 1)))
