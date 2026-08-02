;; vybe-test: wast/wat_control_flow/br_to_block
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module (func (block $b br $b)))
