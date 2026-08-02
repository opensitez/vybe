;; vybe-test: wast/wat_execution/global_immutable_read
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (global $c i32 (i32.const 42))
  (func (export "_start")
    global.get $c
    call $log))
