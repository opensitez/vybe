;; vybe-test: wast/wat_instructions/i32_clz
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i32) (result i32) local.get 0 i32.clz))
