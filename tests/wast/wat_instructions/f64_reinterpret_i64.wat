;; vybe-test: wast/wat_instructions/f64_reinterpret_i64
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param i64) (result f64) local.get 0 f64.reinterpret_i64))
