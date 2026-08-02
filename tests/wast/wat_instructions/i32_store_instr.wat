;; vybe-test: wast/wat_instructions/i32_store_instr
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (memory 1) (func (param i32 i32) local.get 0 local.get 1 i32.store))
