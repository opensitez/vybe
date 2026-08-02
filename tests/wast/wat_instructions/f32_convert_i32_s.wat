;; vybe-test: wast/wat_instructions/f32_convert_i32_s
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i32) (result f32) local.get 0 f32.convert_i32_s))
