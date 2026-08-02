;; vybe-test: wast/wat_instructions/f32_demote_f64
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param f64) (result f32) local.get 0 f32.demote_f64))
