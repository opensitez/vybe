;; vybe-test: wast/wat_folded/folded_four_levels
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (func (param i32) (result i32) (i32.add (i32.mul (i32.add (local.get 0) (i32.const 1)) (i32.const 2)) (i32.const 3))))
