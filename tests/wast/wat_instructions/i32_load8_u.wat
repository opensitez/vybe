;; vybe-test: wast/wat_instructions/i32_load8_u
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (memory 1) (func (param i32) (result i32) local.get 0 i32.load8_u))
