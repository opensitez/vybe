;; vybe-test: wast/wat_execution/nop_is_transparent
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    nop
    nop
    i32.const 5
    nop
    call $log))
