;; vybe-test: wast/wat_instructions/folded_add
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i32 i32) (result i32) (i32.add (local.get 0) (local.get 1))))
