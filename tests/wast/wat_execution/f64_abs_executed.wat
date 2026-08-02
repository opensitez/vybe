;; vybe-test: wast/wat_execution/f64_abs_executed
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f64.const -7.0
    f64.abs
    call $log))
