;; vybe-test: wast/wat_execution/i32_add_executed
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 30
    i32.const 12
    i32.add
    call $log))
