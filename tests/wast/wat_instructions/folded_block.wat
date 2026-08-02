;; vybe-test: wast/wat_instructions/folded_block
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (block $b (br $b))))
