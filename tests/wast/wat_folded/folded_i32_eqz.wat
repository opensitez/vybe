;; vybe-test: wast/wat_folded/folded_i32_eqz
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (func (param i32) (result i32) (i32.eqz (local.get 0))))
