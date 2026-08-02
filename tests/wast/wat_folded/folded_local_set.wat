;; vybe-test: wast/wat_folded/folded_local_set
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (func (param i32) (local i32) (local.set 1 (local.get 0))))
