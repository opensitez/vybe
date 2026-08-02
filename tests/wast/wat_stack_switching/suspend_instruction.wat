;; vybe-test: wast/wat_stack_switching/suspend_instruction
;; origin: languages/wast/tests/wast/test_wat_stack_switching.rs
;; vybe-test-mode: compile

(module (tag $yield (param i32))
          (func (export "_start") i32.const 5 suspend $yield))
