;; vybe-test: wast/wat_folded/folded_i32_load
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (memory 1) (func (param i32) (result i32) (i32.load (local.get 0))))
