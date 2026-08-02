;; vybe-test: wast/wat_execution/i32_eq_executed
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 42
    i32.const 42
    i32.eq
    call $log))
