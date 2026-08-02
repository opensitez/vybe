;; vybe-test: wast/wat_execution/i32_div_s_executed
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 100
    i32.const 4
    i32.div_s
    call $log))
