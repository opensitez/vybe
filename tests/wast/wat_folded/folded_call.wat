;; vybe-test: wast/wat_folded/folded_call
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (func $f (param i32) (result i32) (local.get 0)) (func (result i32) (call $f (i32.const 5))))
