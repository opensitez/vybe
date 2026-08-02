;; vybe-test: wast/wat_folded/folded_i32_wrap_i64
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (func (param i64) (result i32) (i32.wrap_i64 (local.get 0))))
