;; vybe-test: wast/wat_instructions/f32_neg
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param f32) (result f32) local.get 0 f32.neg))
