;; vybe-test: wast/wat_types/invalid_global_immutable_write
;; origin: languages/wast/tests/wast/test_wat_types.rs
;; vybe-test-mode: compile-fail

(module (global $g i32 (i32.const 0)) (func global.set $g (i32.const 1)))
