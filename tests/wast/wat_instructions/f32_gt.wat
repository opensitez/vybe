;; vybe-test: wast/wat_instructions/f32_gt
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param f32 f32) (result i32) local.get 0 local.get 1 f32.gt))
