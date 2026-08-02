;; vybe-test: wast/wat_instructions/br_if_instr
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i32) block local.get 0 br_if 0 end))
