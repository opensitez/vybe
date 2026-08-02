;; vybe-test: wast/wat_instructions/deeply_nested_folded
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i32 i32 i32) (result i32) (i32.add (i32.mul (local.get 0) (local.get 1)) (local.get 2))))
