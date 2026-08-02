;; vybe-test: wast/wat_instructions/memory_grow_instr
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (memory 1) (func (param i32) (result i32) local.get 0 memory.grow))
