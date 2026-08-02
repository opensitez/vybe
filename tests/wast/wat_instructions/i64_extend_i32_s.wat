;; vybe-test: wast/wat_instructions/i64_extend_i32_s
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i32) (result i64) local.get 0 i64.extend_i32_s))
