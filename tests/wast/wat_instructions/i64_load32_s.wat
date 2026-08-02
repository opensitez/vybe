;; vybe-test: wast/wat_instructions/i64_load32_s
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (memory 1) (func (param i32) (result i64) local.get 0 i64.load32_s))
