;; vybe-test: wast/wat_folded/folded_memory_grow
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (memory 1) (func (param i32) (result i32) (memory.grow (local.get 0))))
