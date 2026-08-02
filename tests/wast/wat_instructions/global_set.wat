;; vybe-test: wast/wat_instructions/global_set
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (global $g (mut i32) (i32.const 0)) (func i32.const 1 global.set $g))
