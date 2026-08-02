;; vybe-test: wast/wat_folded/folded_i32_eq
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (func (param i32 i32) (result i32) (i32.eq (local.get 0) (local.get 1))))
