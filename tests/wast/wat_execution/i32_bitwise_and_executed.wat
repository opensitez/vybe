;; vybe-test: wast/wat_execution/i32_bitwise_and_executed
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 0xFF
    i32.const 0x0F
    i32.and
    call $log))
