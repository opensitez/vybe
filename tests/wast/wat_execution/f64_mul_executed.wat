;; vybe-test: wast/wat_execution/f64_mul_executed
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f64.const 3.0
    f64.const 4.0
    f64.mul
    call $log))
