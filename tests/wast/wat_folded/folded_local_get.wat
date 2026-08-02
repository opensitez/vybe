;; vybe-test: wast/wat_folded/folded_local_get
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (func (param $x i32) (result i32) (local.get $x)))
