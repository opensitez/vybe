;; vybe-test: wast/wat_instructions/nested_block_comment
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (; outer (; inner ;) outer ;) (func))
