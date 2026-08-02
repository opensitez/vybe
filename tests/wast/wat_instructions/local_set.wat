;; vybe-test: wast/wat_instructions/local_set
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i32) (local i32) local.get 0 local.set 1))
