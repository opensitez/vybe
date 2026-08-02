;; vybe-test: wast/wat_execution/multiple_log_calls_in_order
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 1 call $log
    i32.const 2 call $log
    i32.const 3 call $log
    i32.const 4 call $log
    i32.const 5 call $log))
