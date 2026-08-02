;; vybe-test: wast/wat_instructions/f64_convert_i32_u
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i32) (result f64) local.get 0 f64.convert_i32_u))
