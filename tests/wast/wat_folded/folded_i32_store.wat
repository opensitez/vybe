;; vybe-test: wast/wat_folded/folded_i32_store
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (memory 1) (func (param i32 i32) (i32.store (local.get 0) (local.get 1))))
