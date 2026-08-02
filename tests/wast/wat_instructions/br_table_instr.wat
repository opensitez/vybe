;; vybe-test: wast/wat_instructions/br_table_instr
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i32) block block local.get 0 br_table 0 1 end end))
