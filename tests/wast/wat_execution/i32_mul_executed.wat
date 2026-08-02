;; vybe-test: wast/wat_execution/i32_mul_executed
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 6
    i32.const 7
    i32.mul
    call $log))
