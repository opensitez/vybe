;; vybe-test: wast/wat_folded/folded_f64_mul
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (func (param f64 f64) (result f64) (f64.mul (local.get 0) (local.get 1))))
