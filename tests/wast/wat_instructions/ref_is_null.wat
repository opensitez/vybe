;; vybe-test: wast/wat_instructions/ref_is_null
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param funcref) (result i32) local.get 0 ref.is_null))
