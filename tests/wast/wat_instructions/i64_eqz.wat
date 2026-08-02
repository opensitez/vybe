;; vybe-test: wast/wat_instructions/i64_eqz
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i64) (result i32) local.get 0 i64.eqz))
