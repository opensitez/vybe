;; vybe-test: wast/wat_instructions/call_instr
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func $f (param i32) (result i32) local.get 0) (func (result i32) i32.const 5 call $f))
