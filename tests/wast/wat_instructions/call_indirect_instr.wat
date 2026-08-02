;; vybe-test: wast/wat_instructions/call_indirect_instr
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (type $t (func (param i32) (result i32))) (table 1 funcref) (func (param i32 i32) (result i32) local.get 0 local.get 1 call_indirect (type $t)))
