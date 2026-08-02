;; vybe-test: wast/wat_execution/f64_neg_executed
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f64.const 5.0
    f64.neg
    call $log))
