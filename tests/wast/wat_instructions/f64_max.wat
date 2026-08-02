;; vybe-test: wast/wat_instructions/f64_max
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param f64 f64) (result f64) local.get 0 local.get 1 f64.max))
