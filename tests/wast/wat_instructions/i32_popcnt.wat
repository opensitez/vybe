;; vybe-test: wast/wat_instructions/i32_popcnt
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i32) (result i32) local.get 0 i32.popcnt))
