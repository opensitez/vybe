;; vybe-test: wast/wat_instructions/i32_trunc_f32_s
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param f32) (result i32) local.get 0 i32.trunc_f32_s))
