;; vybe-test: wast/wat_instructions/folded_if_then
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i32) (if (local.get 0) (then nop))))
