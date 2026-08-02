;; vybe-test: wast/wat_instructions/i32_wrap_i64
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i64) (result i32) local.get 0 i32.wrap_i64))
