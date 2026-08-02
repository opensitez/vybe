;; vybe-test: wast/wat_folded/folded_global_set
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (global $g (mut i32) (i32.const 0)) (func (global.set $g (i32.const 1))))
