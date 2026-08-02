;; vybe-test: wast/wat_instructions/f64_lt
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param f64 f64) (result i32) local.get 0 local.get 1 f64.lt))
