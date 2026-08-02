;; vybe-test: wast/wat_folded/folded_i64_extend_i32_s
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (func (param i32) (result i64) (i64.extend_i32_s (local.get 0))))
