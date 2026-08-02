;; vybe-test: wast/wat_instructions/global_get
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (global $g i32 (i32.const 0)) (func (result i32) global.get $g))
