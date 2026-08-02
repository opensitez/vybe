;; vybe-test: wast/wat_execution/i32_shl_executed
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 1
    i32.const 3
    i32.shl
    call $log))
