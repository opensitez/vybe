;; vybe-test: wast/wat_folded/folded_local_tee
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (func (param i32) (result i32) (local i32) (local.tee 1 (local.get 0))))
