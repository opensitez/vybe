;; vybe-test: wast/wat_execution/i32_shr_u_executed
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 128
    i32.const 2
    i32.shr_u
    call $log))
