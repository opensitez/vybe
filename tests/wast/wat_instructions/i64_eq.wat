;; vybe-test: wast/wat_instructions/i64_eq
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i64 i64) (result i32) local.get 0 local.get 1 i64.eq))
