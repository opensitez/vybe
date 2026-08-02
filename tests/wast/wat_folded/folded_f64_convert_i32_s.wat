;; vybe-test: wast/wat_folded/folded_f64_convert_i32_s
;; origin: languages/wast/tests/wast/test_wat_folded.rs
;; vybe-test-mode: compile

(module (func (param i32) (result f64) (f64.convert_i32_s (local.get 0))))
