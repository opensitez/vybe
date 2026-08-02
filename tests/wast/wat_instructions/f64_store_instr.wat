;; vybe-test: wast/wat_instructions/f64_store_instr
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (memory 1) (func (param i32 f64) local.get 0 local.get 1 f64.store))
