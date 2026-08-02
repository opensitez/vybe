;; vybe-test: wast/wat_instructions/i32_add
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.add))
