;; vybe-test: wast/wat_folded/folded_br_if
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (func (param i32) (block $b (br_if $b (local.get 0)))))
