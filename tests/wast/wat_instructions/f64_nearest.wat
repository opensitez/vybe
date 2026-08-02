;; vybe-test: wast/wat_instructions/f64_nearest
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param f64) (result f64) local.get 0 f64.nearest))
