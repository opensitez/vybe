;; vybe-test: wast/wat_execution/i32_rem_s_executed
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 17
    i32.const 5
    i32.rem_s
    call $log))
