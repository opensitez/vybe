;; vybe-test: wast/wat_folded/folded_i32_lt_s
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (func (param i32 i32) (result i32) (i32.lt_s (local.get 0) (local.get 1))))
