;; vybe-test: wast/wat_instructions/memory_size_instr
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (memory 1) (func (result i32) memory.size))
