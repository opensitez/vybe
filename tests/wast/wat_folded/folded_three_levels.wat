;; vybe-test: wast/wat_folded/folded_three_levels
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (func (param i32 i32 i32) (result i32) (i32.add (i32.mul (local.get 0) (local.get 1)) (local.get 2))))
