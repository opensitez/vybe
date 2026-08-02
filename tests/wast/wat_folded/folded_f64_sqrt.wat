;; vybe-test: wast/wat_folded/folded_f64_sqrt
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (func (param f64) (result f64) (f64.sqrt (local.get 0))))
