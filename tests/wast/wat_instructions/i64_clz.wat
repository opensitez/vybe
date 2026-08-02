;; vybe-test: wast/wat_instructions/i64_clz
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i64) (result i64) local.get 0 i64.clz))
