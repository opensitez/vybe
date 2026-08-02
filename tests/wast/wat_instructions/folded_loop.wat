;; vybe-test: wast/wat_instructions/folded_loop
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (loop $l (br $l))))
