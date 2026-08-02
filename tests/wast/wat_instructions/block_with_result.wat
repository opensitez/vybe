;; vybe-test: wast/wat_instructions/block_with_result
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (result i32) block (result i32) i32.const 1 end))
