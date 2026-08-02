;; vybe-test: wast/wat_instructions/i32_extend8_s
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i32) (result i32) local.get 0 i32.extend8_s))
