;; vybe-test: wast/wat_execution/f64_add_executed
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f64.const 1.5
    f64.const 2.5
    f64.add
    call $log))
