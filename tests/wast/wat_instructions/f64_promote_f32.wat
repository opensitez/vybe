;; vybe-test: wast/wat_instructions/f64_promote_f32
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param f32) (result f64) local.get 0 f64.promote_f32))
